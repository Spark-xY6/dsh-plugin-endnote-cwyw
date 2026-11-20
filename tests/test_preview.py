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
                else:
                    self.assertIn("quotation", report["review_reasons"][0])


if __name__ == "__main__":
    unittest.main()
