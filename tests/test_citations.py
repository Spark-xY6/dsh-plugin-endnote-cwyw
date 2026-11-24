"""Public synthetic regression cases. All I/O uses a per-test temporary directory."""

import io
import json
import sqlite3
import sys
import tempfile
import unittest
from contextlib import closing, redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import patch

from docx import Document
import fitz

SKILLS = Path(__file__).resolve().parents[1] / "skills"
sys.path.insert(0, str(SKILLS / "cite-from-pdfs" / "scripts"))

import apply_citations
import index_pdf_folder
from _common import check_work_path, normalize_doi
from inject_placeholders import inject_into_docx
from placement_evidence import verify_plan_evidence
from reference_identity import IdentityReview, RecordIndex
from rebind_cwyw import document_citations
from sync_to_endnote import plan_import, selected_pdfs, sha256_file, sync_candidate


TITLE = "Archiving intermediate outputs in a small demonstration workflow"
CLAIM = "Archiving intermediate outputs preserves the transformation history for later inspection."


def article(key, **changes):
    return {
        "pdf_id": key, "title": TITLE, "authors": [{"family": "Example", "given": "A"}],
        "year": "2024", "first_author_family": "Example", "journal": "Illustrative Journal",
        "volume": "2", "pages": "10-18", "doi": "10.5555/demo-a", **changes,
    }


class CitationTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        # Windows runners may return an 8.3 spelling for the temporary directory.
        # Production paths are resolved before the isolated library is created.
        self.root = Path(temporary.name).resolve(strict=True)

    def test_doi_alias_does_not_change_significant_suffix(self):
        self.assertEqual(normalize_doi("https://doi.org/10.5555%2FDemo-A?view=1"), "10.5555/demo-a")
        self.assertNotEqual(normalize_doi("10.5555/demo.pdf"), normalize_doi("10.5555/demo"))

    def test_stale_catalog_number_is_never_identity(self):
        unrelated = article("other", title="A different paper", doi="10.5555/demo-b")
        resolved, pending, _ = plan_import([article("P1", record_id=42)], {42: unrelated, 9: article("target")})
        self.assertEqual(resolved, {"P1": 9})
        self.assertEqual(pending, {})

    def test_variant_without_doi_is_one_import(self):
        alias = article("P2", title=TITLE.replace("intermediate", "intermedate"), doi="")
        _, pending, aliases = plan_import([article("P1"), alias], {})
        self.assertEqual(list(pending), ["P1"])
        self.assertEqual(aliases, {"P1": "P1", "P2": "P1"})

    def test_prefer_existing_citation_and_repeat_without_new_import(self):
        library = {3: article("first"), 8: article("cited")}
        resolved, pending, _ = plan_import([article("P1"), article("P2", doi="")], library, preferred={8})
        self.assertEqual(resolved, {"P1": 8, "P2": 8})
        self.assertEqual(pending, {})

    def test_same_title_conflicting_identifiers_need_review(self):
        for change in ({"doi": "10.5555/demo-b"}, {"pmid": "222"}):
            with self.subTest(change=change), self.assertRaises(IdentityReview):
                RecordIndex({8: article("old", pmid="111")}).match(article("new", **change))

    def test_ambiguous_variant_without_corroborating_journal(self):
        old = article("old", doi="", title=TITLE.replace("intermediate", "intermedate"))
        with self.assertRaises(IdentityReview):
            RecordIndex({8: old}).match(article("new", doi="", journal="Another Journal", pages="30", volume="7"))

    def test_title_only_bridge_cannot_hide_two_conflicting_dois(self):
        for ordering in (("weak", "first", "other"), ("first", "weak", "other"), ("other", "first", "weak")):
            entries = {"weak": article("weak", doi=""), "first": article("first"),
                       "other": article("other", doi="10.5555/demo-b")}
            with self.subTest(ordering=ordering), self.assertRaises(IdentityReview):
                plan_import([entries[key] for key in ordering], {})
            with self.subTest(ordering=ordering, existing=True), self.assertRaises(IdentityReview):
                plan_import([entries[key] for key in ordering], {32: article("in library", doi="")})

    def test_wrong_shared_doi_even_with_same_author(self):
        reused = article("wrong", title="Documenting coastal tide measurements collected during a storm")
        with self.assertRaises(IdentityReview):
            RecordIndex({4: reused}).match(article("query"))

    def test_subtitle_with_shared_doi_checks_attribution(self):
        extended = article("extended", title=TITLE + ": a worked example", year="2025")
        self.assertEqual(RecordIndex({4: article("old")}).match(extended), 4)
        conflicting = {**extended, "first_author_family": "Different", "year": "2005"}
        with self.assertRaises(IdentityReview):
            RecordIndex({4: article("old")}).match(conflicting)

    def test_short_titles_with_reused_identifier(self):
        for identity in ({"doi": "10.5555/demo-a"}, {"doi": "", "pmid": "1234"}):
            original = article("old", title="Sleep and memory", first_author_family="Smith", **identity)
            unrelated = article("new", title="Cardiac fibrosis", first_author_family="Jones", **identity)
            with self.subTest(identity=identity), self.assertRaises(IdentityReview):
                plan_import([unrelated], {4: original})
            for alternate in ("Sleep and memory: an overview", "Sleep-and-memory"):
                variant = {**original, "pdf_id": "variant", "title": alternate}
                with self.subTest(identity=identity, alternate=alternate):
                    self.assertEqual(RecordIndex({4: original}).match(variant), 4)

    def test_generic_title_prefix_cannot_hide_reused_identifier(self):
        for identity in ({"doi": "10.5555/demo-a"}, {"doi": "", "pmid": "1234"}):
            old = article("old", title="Sleep", first_author_family="Smith", **identity)
            for author in ("Jones", ""):
                new = article("different", title="Sleep quality after surgery",
                              first_author_family=author, journal="Another Journal", **identity)
                with self.subTest(identity=identity, author=author), self.assertRaises(IdentityReview):
                    plan_import([new], {8: old})

    def test_existing_metadata_and_pdf_alias_keep_attachment(self):
        selected = article("pdf", doi="", pdf_path="illustrative.pdf", pdf_sha256="sample-digest")
        _, pending, aliases = plan_import([selected, article("embedded")], {})
        self.assertEqual(aliases["pdf"], "embedded")
        self.assertEqual(pending["embedded"]["pdf_path"], selected["pdf_path"])
        self.assertEqual(pending["embedded"]["doi"], article("embedded")["doi"])

    def test_evidence_must_be_on_the_claimed_physical_page_for_every_alias(self):
        path = self.root / "pages.pdf"
        with fitz.open() as pdf:
            pdf.new_page().insert_text((40, 70), "Context page")
            pdf.new_page().insert_text((40, 70), "An archived transformation can be inspected later.")
            pdf.save(path)
        selected = [article(key, pdf_path=str(path)) for key in ("A", "B")]
        proof = {"pdf_id": "A", "page": 2, "quote": "transformation can be inspected"}
        placement = {"insertion_id": "C1", "pdf_ids": ["A", "B"], "sources": [proof, {**proof, "pdf_id": "B"}],
                     "rationale": "The cited methods describe inspecting the archived output."}
        ok = verify_plan_evidence({"insertions": [placement]}, selected)
        self.assertEqual([(row["pdf_id"], row["page"]) for row in ok], [("A", 2), ("B", 2)])
        for sources in ([{**proof, "page": 1}, {**proof, "pdf_id": "B"}], [proof],
                        [{**proof, "quote": "invented evidence"}, {**proof, "pdf_id": "B"}]):
            with self.subTest(sources=sources), self.assertRaises(ValueError):
                verify_plan_evidence({"insertions": [{**placement, "sources": sources}]}, selected)

    def test_evidence_and_attachment_share_the_indexed_bytes(self):
        original = self.root / "article.pdf"
        with fitz.open() as pdf:
            pdf.new_page().insert_text((40, 70), "Words from the indexed version.")
            pdf.save(original)
        digest = sha256_file(original)
        catalog = self.root / "catalog.json"
        catalog.write_text(json.dumps({"items": [article("P", pdf_path=str(original), pdf_sha256=digest)]}))
        selected = selected_pdfs(catalog, ["P"], self.root / "scratch")
        original.write_bytes(b"replaced original after selection")
        self.assertEqual(sha256_file(Path(selected[0]["pdf_path"])), digest)
        placement = {"pdf_ids": ["P"], "rationale": "Illustrative proof.", "sources": [
            {"pdf_id": "P", "page": 1, "quote": "Words from the indexed version"}]}
        self.assertEqual(verify_plan_evidence({"insertions": [placement]}, selected)[0]["page"], 1)
        with self.assertRaisesRegex(RuntimeError, "PDF changed"):
            selected_pdfs(catalog, ["P"], self.root / "new-scratch")

    def test_private_pdf_index_and_page_search(self):
        papers = self.root / "pdfs"
        papers.mkdir()
        with fitz.open() as pdf:
            page = pdf.new_page()
            page.insert_text((40, 65), TITLE)
            page.insert_textbox((40, 110, 550, 300), CLAIM)
            pdf.set_metadata({"title": TITLE, "subject": "DOI: 10.5555/demo-a"})
            pdf.save(papers / "paper.pdf")
        catalog = self.root / "catalog.json"
        parser = index_pdf_folder.build_parser()
        args = parser.parse_args(["index", "--pdf-dir", str(papers), "--work-dir", str(self.root / "index"),
                                  "--output", str(catalog)])
        report = self.root / "search.json"
        with (patch.object(index_pdf_folder, "enrich_metadata", side_effect=AssertionError("unexpected lookup")),
              redirect_stdout(io.StringIO())):
            self.assertEqual(index_pdf_folder.cmd_index(args), 0)
            search = parser.parse_args(["search", "--catalog", str(catalog), "--query", "transformation history",
                                        "--report", str(report)])
            self.assertEqual(index_pdf_folder.cmd_search(search), 0)
        hits = json.loads(report.read_text(encoding="utf-8"))["hits"]
        self.assertEqual(hits[0]["pdf_id"], "P001")
        self.assertTrue(any(row["page"] == "1" and CLAIM in row["text"] for row in hits[0]["top_passages"]))

    def test_two_pdf_aliases_at_one_anchor_make_one_token(self):
        source, target = self.root / "draft.docx", self.root / "prepared.docx"
        doc = Document()
        doc.add_paragraph(CLAIM)
        doc.save(source)
        frozen = sha256_file(source)
        entry = {"record_id": 7, "citation_token": "Example, 2024 #7"}
        inject_into_docx(source, target, {"insertions": [{"anchor_text": CLAIM, "pdf_ids": ["P1", "P2"]}]},
                         {"by_pdf_id": {"P1": entry, "P2": entry}})
        self.assertEqual(document_citations(target, prepared=True)[1][0][2], (7,))
        self.assertEqual(Document(target).paragraphs[0].text.count("#7"), 1)
        self.assertEqual(sha256_file(source), frozen)

    def test_changed_receipt_is_rejected(self):
        (self.root / "ready.json").write_text("{}", encoding="utf-8")
        args = ["apply", "--publish-prepared", str(self.root), "--ready-sha256", "0" * 64]
        with (patch.object(sys, "argv", args), patch.object(apply_citations, "finalize") as finalizer,
              redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as error):
            apply_citations.main()
        self.assertEqual(error.exception.code, 2)
        finalizer.assert_not_called()

    def test_prevent_work_inside_library(self):
        with self.assertRaises(ValueError):
            check_work_path(self.root / "Library.Data" / "sdb" / "report.json", ".json")

    def test_interrupted_import_only_alters_unpublished_copy(self):
        source = self.root / "sample.enl"
        source.write_bytes(b"synthetic library marker")
        sibling = source.with_suffix(".Data") / "sdb"
        sibling.mkdir(parents=True)
        db = sibling / "sdb.eni"
        with closing(sqlite3.connect(db)) as conn:
            conn.executescript("""CREATE TABLE refs (id INTEGER, title TEXT);
                CREATE TABLE misc (code INTEGER, subcode INTEGER, value TEXT);
                INSERT INTO misc VALUES (14,0,'synthetic-example');""")
        original = [sha256_file(source), sha256_file(db)]
        destination = self.root / "delivery"

        def interrupt(candidate, _import_file):
            self.assertTrue(candidate.resolve(strict=True).is_relative_to(self.root), "candidate is outside the test root")
            self.assertNotEqual(candidate, source)
            candidate.write_bytes(b"interrupted candidate")
            raise RuntimeError("injected import interruption")

        def import_only(_args, work, stage):
            sync_candidate(source, stage / "library", [article("P1")], work)

        argv = ["apply", "--prepare-only", "--enl", str(source), "--docx", "unused", "--catalog", "unused",
                "--plan", "unused", "--work-dir", str(self.root / "work"), "--output-dir", str(destination)]
        with (patch.object(sys, "argv", argv), patch.object(apply_citations, "apply", side_effect=import_only),
              patch("sync_to_endnote.require_closed_endnote"), redirect_stderr(io.StringIO()),
              patch("sync_to_endnote.run_endnote_import", side_effect=interrupt)):
            self.assertEqual(apply_citations.main(), 1)
        self.assertFalse(destination.exists())
        failure = next((self.root / "work").glob("citation-run-*/failure.json"))
        self.assertIn("injected import interruption", json.loads(failure.read_text(encoding="utf-8"))["error"])
        self.assertEqual([sha256_file(source), sha256_file(db)], original)

    def test_v3_accuracy_normalization_offline_authors_columns_and_nature_spacing(self):
        import _common
        from reference_identity import Identity, compare
        from sync_to_endnote import _truncate_abstract, format_enw_entry
        import fitz

        # 1. Multi-line DOI, unbalanced parens, and .pdf suffix sensitivity
        self.assertEqual(_common.normalize_doi("PII S0092\r10.1016/j.cell.2024.01.001."), "10.1016/j.cell.2024.01.001")
        self.assertEqual(_common.normalize_doi("(doi: 10.1038/s41586-024-00001-1)"), "10.1038/s41586-024-00001-1")
        self.assertNotEqual(_common.normalize_doi("10.5555/demo.pdf"), _common.normalize_doi("10.5555/demo"))

        # 2. Greek + HTML title normalization and author particle/suffix parsing
        self.assertEqual(
            _common.normalize_title("<i>NF-κB</i> &amp; β-catenin pathway"),
            _common.normalize_title("NF-kappaB & beta-catenin pathway"),
        )
        parsed = _common.parse_endnote_authors("Ludwig van Beethoven Jr.\nSmith, John, III")
        self.assertEqual(parsed[0], {"family": "van Beethoven", "given": "Ludwig"})
        self.assertEqual(parsed[1], {"family": "Smith", "given": "John"})

        # 3. Journal abbreviation + Epub +/- 1 year match without spurious ambiguity
        a = Identity.from_record({
            "title": "Single-cell chromatin accessibility mapping across mouse development",
            "first_author_family": "Zhang", "year": "2024", "journal": "Biol. Cybern.",
        })
        b = Identity.from_record({
            "title": "Single-cell chromatin accessibility mapping across mouse development",
            "first_author_family": "Zhang", "year": "2023", "journal": "Biological Cybernetics",
        })
        self.assertEqual(compare(a, b), "title_author_year")

        # 4. Sentence/word-safe abstract truncation in .enw export
        long_abs = ("This sentence stays intact. " * 115) + "Trailing incomplete fragment."
        enw = format_enw_entry({"title": "Demo", "abstract": long_abs, "authors": [{"family": "Zhang", "given": "Y."}]})
        abs_line = [ln[3:] for ln in enw.splitlines() if ln.startswith("%X ")][0]
        self.assertLessEqual(len(abs_line), 3000)
        self.assertTrue(abs_line.endswith("."))

        # 5. Offline two-column PDF indexing preserves XMP/byline authors + flags scanned PDFs
        pdf_dir = self.root / "pdfs_v3"
        pdf_dir.mkdir()
        col_pdf = pdf_dir / "two_col.pdf"
        with fitz.open() as doc:
            page = doc.new_page(width=612, height=792)
            page.insert_text((60, 70), "Neural Autocorrelation Models of Acoustic Processing in Mammals", fontsize=15)
            page.insert_text((60, 95), "Alice Smith, Bob van der Berg", fontsize=11)
            page.insert_textbox(fitz.Rect(50, 140, 280, 210), "Left column first paragraph discussing auditory modeling and filterbanks in full detail.")
            page.insert_textbox(fitz.Rect(50, 220, 280, 290), "Left column second paragraph reporting empirical response latencies and thresholds.")
            page.insert_textbox(fitz.Rect(320, 140, 560, 210), "Right column first paragraph continuing the statistical validation across species.")
            page.insert_textbox(fitz.Rect(320, 220, 560, 290), "Right column second paragraph with conclusion and computational parameters.")
            doc.save(col_pdf)
        scan_pdf = pdf_dir / "scanned.pdf"
        with fitz.open() as doc:
            doc.new_page()
            doc.save(scan_pdf)
        bad_pdf = pdf_dir / "corrupt.pdf"
        bad_pdf.write_bytes(b"not a real pdf")

        catalog = self.root / "catalog_v3.json"
        parser = index_pdf_folder.build_parser()
        args = parser.parse_args(["index", "--pdf-dir", str(pdf_dir), "--work-dir", str(self.root / "idxwork"), "--output", str(catalog)])
        with redirect_stdout(io.StringIO()):
            self.assertEqual(index_pdf_folder.cmd_index(args), 0)
        cat_data = json.loads(catalog.read_text(encoding="utf-8"))
        self.assertEqual(cat_data["skipped_pdf_count"], 1)
        self.assertEqual(cat_data["max_pages"], 30)
        by_name = {Path(i["pdf_path"]).name: i for i in cat_data["items"]}
        self.assertEqual(by_name["two_col.pdf"]["first_author_family"], "Smith")
        self.assertTrue(by_name["two_col.pdf"]["authors"])
        self.assertTrue(by_name["scanned.pdf"]["needs_ocr"])
        text_chunk = by_name["two_col.pdf"]["passages"][0]["text"]
        self.assertLess(text_chunk.index("Left column second"), text_chunk.index("Right column first"))

        # 6. Nature style places citation after English period with ZERO leading space, and before Chinese period
        en_zh = self.root / "nature_place.docx"
        en_zh_out = self.root / "nature_placed.docx"
        doc = Document()
        doc.add_paragraph("Record the exact steps used to produce each computational result.")
        doc.add_paragraph("记录每次计算所需的全部步骤。")
        doc.save(en_zh)
        entry = {"record_id": 7, "citation_token": "Example, 2024 #7"}
        inject_into_docx(
            en_zh,
            en_zh_out,
            {
                "style": "Nature",
                "insertions": [
                    {"insertion_id": "C01", "anchor_text": "Record the exact steps used to produce each computational result.", "pdf_ids": ["P1"]},
                    {"insertion_id": "C02", "anchor_text": "记录每次计算所需的全部步骤。", "pdf_ids": ["P1"]},
                ],
            },
            {"by_pdf_id": {"P1": entry}},
        )
        out_paras = [p.text for p in Document(en_zh_out).paragraphs]
        self.assertEqual(out_paras[0], "Record the exact steps used to produce each computational result.{Example, 2024 #7}")
        self.assertEqual(out_paras[1], "记录每次计算所需的全部步骤{Example, 2024 #7}。")

        # 7. Multi-author PDF Info parsing + surname particles + uncorroborated arXiv rejection
        from _common import parse_pdf_meta_authors
        self.assertEqual(
            [a["family"] for a in parse_pdf_meta_authors("Alice Smith; Bob Jones")],
            ["Smith", "Jones"],
        )
        self.assertEqual(
            [a["family"] for a in parse_pdf_meta_authors("Alice Smith, Bob Jones")],
            ["Smith", "Jones"],
        )
        self.assertEqual(
            parse_pdf_meta_authors("Ludwig van Beethoven"),
            [{"family": "van Beethoven", "given": "Ludwig"}],
        )
        self.assertEqual(
            compare(
                Identity.from_record({"title": "Ninth Symphony Acoustics Study in Europe", "first_author_family": "van Beethoven", "year": "2024"}),
                Identity.from_record({"title": "Ninth Symphony Acoustics Study in Europe", "first_author_family": "Beethoven", "year": "2024"}),
            ),
            "title_author_year",
        )

        # Uncorroborated arXiv lookup must not overwrite the PDF's real title/authors
        arxiv_dir = self.root / "pdfs_arxiv"
        arxiv_dir.mkdir()
        p_arxiv = arxiv_dir / "citing_preprint.pdf"
        with fitz.open() as doc:
            pg = doc.new_page()
            pg.insert_text((60, 70), "Original Experimental Study on Yeast Heat Shock Proteins", fontsize=15)
            pg.insert_text((60, 95), "Jane Doe", fontsize=11)
            pg.insert_text((60, 140), "Earlier methods were proposed in arXiv:2301.12345.", fontsize=11)
            doc.save(p_arxiv)
        fake_oa = (
            {
                "title": "Unrelated Machine Learning Preprint Mentioned In Passing",
                "authors": [{"family": "WrongAuthor", "given": "X."}],
                "first_author_family": "WrongAuthor",
                "year": "2023",
                "journal": "arXiv",
                "volume": "",
                "issue": "",
                "pages": "",
                "doi": "10.48550/arXiv.2301.12345",
                "abstract": "",
                "is_retracted": False,
            },
            "openalex_arxiv",
            "direct_no_proxy",
        )
        cat_arxiv = self.root / "catalog_arxiv.json"
        args_arxiv = index_pdf_folder.build_parser().parse_args([
            "index", "--pdf-dir", str(arxiv_dir), "--work-dir", str(self.root / "idxwork_arxiv"),
            "--output", str(cat_arxiv), "--enrich-missing-metadata",
        ])
        from unittest.mock import patch
        with patch.object(index_pdf_folder, "enrich_metadata", return_value=fake_oa):
            with redirect_stdout(io.StringIO()):
                self.assertEqual(index_pdf_folder.cmd_index(args_arxiv), 0)
        item_ax = json.loads(cat_arxiv.read_text(encoding="utf-8"))["items"][0]
        self.assertIn("not corroborated", item_ax["identity_review"])
        self.assertIn("Yeast Heat Shock Proteins", item_ax["title"])
        self.assertEqual(item_ax["first_author_family"], "Doe")


if __name__ == "__main__":
    unittest.main()
