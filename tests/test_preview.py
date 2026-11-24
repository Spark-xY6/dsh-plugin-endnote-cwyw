"""Synthetic dry-run cases for the public preview CLI."""

from contextlib import closing, redirect_stdout
import io
import json
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest
from unittest.mock import patch

import fitz
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "skills" / "cite-from-pdfs" / "scripts"))
import preview_import
from sync_to_endnote import sha256_file


def record(key, **changes):
    return {"pdf_id": key, "first_author_family": "Example", "year": "2020",
            "title": "Detecting variation in example catalogue records", "doi": "10.5555/demo",
            "authors": [{"family": "Example"}], "journal": "Sample Journal",
            "volume": "2", "pages": "11-17", **changes}


class PreviewTests(unittest.TestCase):
    def test_foreign_number_and_duplicate_in_target(self):
        old = record("existing:1:1", original_id=8, library_ids=["foreign"], reference_type="Journal Article")
        unrelated = record("unrelated", title="A different work", doi="10.5555/other")
        same = record("existing:1:2", original_id=3, library_ids=["target"], reference_type="Journal Article")
        report = preview_import.forecast([old, same], [record("P1")], {1: record("a"), 3: same, 8: unrelated}, "target")
        self.assertEqual(report["planned_new_works"], 0)
        self.assertEqual([row["target_record_id"] for row in report["rows"]], [3, 3, 3])
        self.assertTrue(report["rows"][0]["requires_rebind"])
        collision = next(c for c in report["rows"][0]["candidates"] if c["id"] == 8)
        self.assertEqual(collision["match"], "different")
        self.assertIn("doi", collision["different_fields"])
        self.assertFalse(report["rows"][1]["requires_rebind"])

    def test_alias_and_conflict_diagnostics(self):
        first = record("P1", doi="")
        report = preview_import.forecast([], [first, record("P2")], {}, "target")
        self.assertEqual(report["planned_new_works"], 1)
        self.assertEqual(report["rows"][0]["canonical_input"], "P2")
        conflict = record("P3", doi="10.5555/conflict")
        report = preview_import.forecast([], [record("P1"), conflict], {}, "target")
        self.assertEqual(report["status"], "REVIEW")
        self.assertIsNone(report["planned_new_works"])
        candidate = report["rows"][0]["candidates"][0]
        self.assertEqual((candidate["id"], candidate["match"]), ("P3", "conflict"))
        self.assertEqual(candidate["normalized_fields"]["doi"], "10.5555/conflict")

    def test_missing_nonjournal_and_target_conflict_return_review(self):
        book = record("existing:1:1", original_id=77, library_ids=["elsewhere"], reference_type="Book")
        report = preview_import.forecast([book], [], {}, "target")
        self.assertEqual(report["status"], "REVIEW")
        self.assertIn("Book", report["review_reasons"][0])
        report = preview_import.forecast([], [record("P1")], {1: record("a", doi="10.5555/other")}, "target")
        self.assertEqual(report["rows"][0]["candidates"][0]["source"], "target_library")
        self.assertEqual(report["status"], "REVIEW")

    def test_cli_private_read_only_snapshot_and_quote_gate(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            docx, enl = root / "draft.docx", root / "sample.enl"
            doc = Document()
            doc.add_paragraph("An illustrative preservation claim.")
            doc.save(docx)
            enl.write_bytes(b"synthetic marker")
            db = enl.with_suffix(".Data") / "sdb" / "sdb.eni"
            db.parent.mkdir(parents=True)
            with closing(sqlite3.connect(db)) as conn:
                conn.executescript("CREATE TABLE refs(id INTEGER, title TEXT);"
                                   "CREATE TABLE misc(code INTEGER, subcode INTEGER, value TEXT);"
                                   "INSERT INTO misc VALUES (14,0,'example-fixture');")
            pdf = root / "paper.pdf"
            quote = "An illustrative preservation claim."
            with fitz.open() as paper:
                paper.new_page().insert_text((30, 40), quote)
                paper.save(pdf)
            source = {"pdf_id": "P1", "page": 1, "quote": quote}
            placement = {"anchor_text": quote, "pdf_ids": ["P1"], "sources": [source],
                         "rationale": "The example states the claim directly."}
            catalog, plan, work = root / "catalog.json", root / "plan.json", root / "work"
            catalog.write_text(json.dumps({"items": [record("P1", pdf_path=str(pdf), pdf_sha256=sha256_file(pdf))]}))
            originals = {path: sha256_file(path) for path in (docx, enl, db, pdf, catalog)}
            for authorized, status in ((quote, "PASS"), ("Missing excerpt", "REVIEW")):
                source["quote"] = authorized
                plan.write_text(json.dumps({"insertions": [placement]}))
                output = root / (status + ".json")
                args = ["--docx", str(docx), "--enl", str(enl), "--plan", str(plan),
                        "--catalog", str(catalog), "--work-dir", str(work), "--output", str(output)]
                console = io.StringIO()
                with patch.object(preview_import, "require_closed_endnote"), redirect_stdout(console):
                    self.assertEqual(preview_import.main(args), int(status != "PASS"))
                report = json.loads(output.read_text(encoding="utf-8"))
                self.assertEqual(report["status"], status)
                self.assertNotIn(quote, console.getvalue())
                self.assertNotIn("catalogue", console.getvalue())
                self.assertEqual(list(work.iterdir()), [])
                self.assertEqual(originals, {path: sha256_file(path) for path in originals})
                if status == "PASS":
                    self.assertEqual(report["planned_new_works"], 1)
                    self.assertEqual(report["verified_pages"][0]["page"], 1)
                    self.assertEqual(report["placements"][0]["paragraph"], 1)
                    self.assertIn("PDF 第 1 页", preview_import.explain(report, "zh"))
                else:
                    self.assertIn("quotation", report["review_reasons"][0])

    def test_shared_locator_ambiguous_stale_and_explicit_occurrence(self):
        maps = [("Same claim.", []), ("Same claim.", [])]
        item = {"anchor_text": "Same claim.", "pdf_ids": ["P1"]}
        with self.assertRaisesRegex(ValueError, "Ambiguous"):
            preview_import.locate_placements(maps, [item])
        with self.assertRaisesRegex(ValueError, "not found"):
            preview_import.locate_placements(maps, [dict(item, anchor_text="Stale.")])
        with self.assertRaisesRegex(ValueError, "out of range"):
            preview_import.locate_placements(maps, [dict(item, occurrence=3)])
        for invalid in (1.9, True):
            with self.subTest(invalid=invalid), self.assertRaisesRegex(ValueError, "integer"):
                preview_import.locate_placements(maps, [dict(item, occurrence=invalid)])
        op = preview_import.locate_placements(maps, [dict(item, occurrence=2)])[0]
        self.assertEqual((op["paragraph_index"], op["char_offset"]), (1, 10))

    def test_fields_across_paragraphs_and_textboxes_are_not_prose(self):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "draft.docx"
            doc = Document()
            start = OxmlElement("w:fldChar")
            start.set(qn("w:fldCharType"), "begin")
            doc.add_paragraph().add_run()._r.append(start)
            doc.add_paragraph("This belongs to a bibliography.")
            end = OxmlElement("w:fldChar")
            end.set(qn("w:fldCharType"), "end")
            doc.add_paragraph().add_run()._r.append(end)
            text = OxmlElement("w:t")
            text.text = "Text box."
            box = OxmlElement("w:txbxContent")
            para = OxmlElement("w:p")
            para.append(text)
            box.append(para)
            doc.add_paragraph().add_run()._r.append(box)
            doc.add_paragraph("Citable.")
            doc.save(path)
            _, maps = preview_import.read_manuscript(path)
            self.assertEqual([text.replace("\x00", "") for text, _ in maps], ["", "", "", "", "Citable."])

    def test_formatted_runs_match_but_structural_boundaries_and_wrappers_do_not(self):
        from docx.oxml import OxmlElement, parse_xml
        from docx.oxml.ns import nsdecls
        from inject_placeholders import paragraph_text_map
        normal = '<w:r><w:rPr><w:b/></w:rPr><w:t>Cell</w:t></w:r>'
        end = '<w:r><w:t>line.</w:t></w:r>'
        field = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
        item = {"pdf_ids": ["P1"], "anchor_text": "Cellline."}
        for boundary in ("", "<w:r><w:tab/></w:r>", "<w:r><w:br/></w:r>", field,
                         '<w:fldSimple w:instr="DATE"><w:r><w:t>2026</w:t></w:r></w:fldSimple>',
                         "<w:r><w:drawing/></w:r>"):
            p = parse_xml(f'<w:p {nsdecls("w")}>{normal}{boundary}{end}</w:p>')
            maps = [paragraph_text_map(p, {"depth": 0})]
            with self.subTest(boundary=boundary):
                if boundary:
                    with self.assertRaisesRegex(ValueError, "not found"):
                        preview_import.locate_placements(maps, [item])
                else:
                    self.assertEqual(preview_import.locate_placements(maps, [item])[0]["char_offset"], 8)
        for wrapper in ("hyperlink", "sdt", "ins"):
            p = OxmlElement("w:p")
            p.append(parse_xml(f'<w:{wrapper} {nsdecls("w")}>{normal}{end}</w:{wrapper}>'))
            p.append(parse_xml(f'<w:r {nsdecls("w")}><w:t>Plain.</w:t></w:r>'))
            maps = [paragraph_text_map(p, {"depth": 0})]
            with self.subTest(wrapper=wrapper), self.assertRaisesRegex(ValueError, "not found"):
                preview_import.locate_placements(maps, [item])
            self.assertEqual(len(preview_import.locate_placements(maps, [dict(item, anchor_text="Plain.")])), 1)

    def test_v3_curly_quotes_comments_bib_exclusion_and_nature_auto_mode(self):
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        from inject_placeholders import paragraph_text_map
        import fitz
        from placement_evidence import page_text, verify_plan_evidence

        # Curly quotes, em-dash & Word comment range markers bridge cleanly while static References are ignored
        with tempfile.TemporaryDirectory() as directory:
            tmp = Path(directory)
            draft = tmp / "rich.docx"
            p = parse_xml(
                f'<w:p {nsdecls("w")}>'
                '<w:r><w:t>“Cell”</w:t></w:r>'
                '<w:commentRangeStart w:id="1"/><w:r><w:t>–</w:t></w:r>'
                '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/><w:t>fate.</w:t></w:r>'
                '</w:p>'
            )
            maps = [paragraph_text_map(p, {"depth": 0})]
            placed = preview_import.locate_placements(
                maps,
                [{"insertion_id": "C01", "anchor_text": '"Cell"-fate.', "pdf_ids": ["P1"]}],
                style="Nature",
            )
            self.assertEqual(placed[0]["char_offset"], 12)
            self.assertEqual(placed[0]["position"], "after_anchor")

            doc = Document()
            doc.add_paragraph('Main claim with "alpha" response.')
            doc.add_paragraph("References")
            doc.add_paragraph('Main claim with "alpha" response.')
            doc.save(draft)
            _, doc_maps = preview_import.read_manuscript(draft)
            single = preview_import.locate_placements(
                doc_maps, [{"anchor_text": 'Main claim with "alpha" response.', "pdf_ids": ["P1"]}], style="Nature"
            )
            self.assertEqual(len(single), 1)
            self.assertEqual(single[0]["paragraph_index"], 0)

            # Consecutive page span "1-2" and normalized dash/quote verification
            self.assertEqual(page_text("“long-range”—coherence"), '"long-range"-coherence')
            pdf_path = tmp / "span.pdf"
            with fitz.open() as pdf_doc:
                pdf_doc.new_page().insert_text((72, 72), "Neural computation shows")
                pdf_doc.new_page().insert_text((72, 72), "long-range coherence across layers.")
                pdf_doc.save(pdf_path)
            verified = verify_plan_evidence(
                {
                    "insertions": [
                        {
                            "insertion_id": "C01",
                            "pdf_ids": ["P1"],
                            "sources": [{"pdf_id": "P1", "page": "1-2", "quote": "shows long-range coherence"}],
                            "rationale": "Cross-page evidence span test.",
                        }
                    ]
                },
                [{"pdf_id": "P1", "pdf_path": str(pdf_path)}],
            )
            self.assertEqual(verified[0]["total_pages"], 2)
            self.assertEqual(verified[0]["page"], "1-2")

            # Span 1-3 + Literature Cited ending before Appendix + Nature Medicine 0-space after_anchor
            pdf_3p = tmp / "span3.pdf"
            with fitz.open() as pdf_doc:
                pdf_doc.new_page().insert_text((72, 72), "Part one.")
                pdf_doc.new_page().insert_text((72, 72), "Part two")
                pdf_doc.new_page().insert_text((72, 72), "Part three.")
                pdf_doc.save(pdf_3p)
            verified3 = verify_plan_evidence(
                {
                    "insertions": [
                        {
                            "insertion_id": "C02",
                            "pdf_ids": ["P1"],
                            "sources": [{"pdf_id": "P1", "page": "1-3", "quote": "Part two Part three."}],
                            "rationale": "Three-page span check.",
                        }
                    ]
                },
                [{"pdf_id": "P1", "pdf_path": str(pdf_3p)}],
            )
            self.assertEqual(verified3[0]["page"], "1-3")

            doc_app = Document()
            doc_app.add_paragraph("References")
            doc_app.add_paragraph("Ignored bibliography line.")
            doc_app.add_paragraph("Appendix A")
            doc_app.add_paragraph("Appendix finding confirmed here.")
            app_docx = tmp / "appendix.docx"
            doc_app.save(app_docx)
            _, app_maps = preview_import.read_manuscript(app_docx)
            app_placed = preview_import.locate_placements(
                app_maps,
                [{"insertion_id": "C03", "anchor_text": "Appendix finding confirmed here.", "pdf_ids": ["P1"]}],
                style="Nature Medicine",
            )
            self.assertEqual(app_placed[0]["position"], "after_anchor")

    def test_preview_non_aborting_mixed_errors_and_settings_style_fallback(self):
        from contextlib import closing
        import sqlite3
        from types import SimpleNamespace
        from zipfile import ZIP_DEFLATED, ZipFile

        with tempfile.TemporaryDirectory() as td:
            root = Path(td).resolve()
            doc = Document()
            doc.add_paragraph("First verified claim.")
            doc.add_paragraph("Second verified claim.")
            base_docx = root / "base.docx"
            doc.save(base_docx)
            styled_docx = root / "styled.docx"
            settings_xml = (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                b'<w:customXml/><w:docVars>'
                b'<w:docVar w:name="EN.Layout" w:val="&lt;Style&gt;Nature&lt;/Style&gt;"/>'
                b'</w:docVars></w:settings>'
            )
            with ZipFile(base_docx) as zin, ZipFile(styled_docx, "w", ZIP_DEFLATED) as zout:
                for n in zin.namelist():
                    if n != "word/settings.xml":
                        zout.writestr(n, zin.read(n))
                zout.writestr("word/settings.xml", settings_xml)
            enl = root / "lib.enl"
            enl.write_bytes(b"synthetic marker")
            db = enl.with_suffix(".Data") / "sdb" / "sdb.eni"
            db.parent.mkdir(parents=True)
            with closing(sqlite3.connect(db)) as conn:
                conn.executescript(
                    "CREATE TABLE refs(id INTEGER, title TEXT);"
                    "CREATE TABLE misc(code INTEGER, subcode INTEGER, value TEXT);"
                    "INSERT INTO misc VALUES (14,0,'example-fixture');"
                )
            pdf = root / "paper.pdf"
            import fitz, hashlib
            with fitz.open() as pdf_doc:
                pdf_doc.new_page().insert_text((72, 72), "Exact quote matches here.")
                pdf_doc.save(pdf)
            sha = hashlib.sha256(pdf.read_bytes()).hexdigest().upper()
            catalog = root / "cat.json"
            catalog.write_text(json.dumps({
                "status": "PASS",
                "items": [{
                    "pdf_id": "P001", "pdf_path": str(pdf), "pdf_sha256": sha,
                    "title": "Sample Study", "authors": [{"family": "Lee", "given": "A."}],
                    "first_author_family": "Lee", "year": "2024", "journal": "Nature",
                }],
            }), encoding="utf-8")
            plan = root / "plan.json"
            plan.write_text(json.dumps({
                "insertions": [
                    {"anchor_text": "First verified claim.", "pdf_ids": ["P001"],
                     "sources": [{"pdf_id": "P001", "page": 1, "quote": "Exact quote matches here."}], "rationale": "ok1"},
                    {"anchor_text": "Second verified claim.", "pdf_ids": ["P001"],
                     "sources": [{"pdf_id": "P001", "page": 1, "quote": "Exact quote matches here."}], "rationale": "ok2"},
                    {"anchor_text": "Missing sentence.", "pdf_ids": ["P999"],
                     "sources": [{"pdf_id": "P999", "page": 1, "quote": "nope"}], "rationale": "bad"},
                ]
            }), encoding="utf-8")
            args = SimpleNamespace(enl=enl, docx=styled_docx, catalog=catalog, plan=plan, work_dir=root / "work", output=root / "out.json")
            report = preview_import.make_preview(args)
            self.assertEqual(report["status"], "REVIEW")
            self.assertEqual(len(report["placements"]), 2)
            self.assertEqual(report["placements"][0]["position"], "after_anchor")
            self.assertTrue(any("C03" in r and "Missing sentence" in r for r in report["review_reasons"]))
            self.assertTrue(any("C03" in r and "P999" in r for r in report["review_reasons"]))


if __name__ == "__main__":
    unittest.main()
