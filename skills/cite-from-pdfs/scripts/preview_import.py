# /// script
# requires-python = ">=3.10"
# dependencies = ["lxml>=5", "PyMuPDF>=1.24"]
# ///
"""Preview a citation import against a closed library. The JSON report is private."""

import argparse
from dataclasses import asdict
import json
from pathlib import Path
import sqlite3
import sys
import tempfile
from zipfile import BadZipFile

from _common import (check_work_path, copy_sdb_snapshot, read_endnote_active_records,
                     read_library_id, sdb_path_for_enl, write_json_atomic)
from placement_evidence import verify_plan_evidence
from reference_identity import Identity, IdentityReview, RecordIndex, compare
from sync_to_endnote import plan_import, require_closed_endnote, selected_pdfs, sha256_file

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "endnote-cwyw" / "scripts"))
from rebind_cwyw import read_citations


def comparison(query, other, source, key, outcome):
    left, right = asdict(query), asdict(other)
    return {
        "source": source, "id": key, "match": outcome,
        "normalized_fields": right,
        "different_fields": [field for field in left if left[field] != right[field]],
    }


def forecast(existing, selected, records, library_id):
    items = existing + selected
    if not items:
        raise ValueError("No existing citations or selected PDFs")
    preferred = {row["original_id"] for row in existing if row["library_ids"] == [library_id]}
    index = RecordIndex(records)
    issues = []
    resolved, aliases = {}, {}
    try:
        resolved, _, aliases = plan_import(items, records, preferred)
        for row in existing:
            if row["pdf_id"] not in resolved and row["reference_type"] != "Journal Article":
                issues.append(f"{row['pdf_id']}: missing {row['reference_type'] or 'untyped'} record requires explicit migration")
    except (IdentityReview, ValueError) as error:
        issues.append(str(error))
    accepted = not issues
    batch = [(item["pdf_id"], Identity.from_record(item)) for item in items]
    rows = []
    for item, (key, identity) in zip(items, batch):
        old = item.get("original_id")
        candidates = []
        for rid, candidate in index.entries:
            outcome = compare(identity, candidate)
            if outcome != "different" or old == rid:
                candidates.append(comparison(identity, candidate, "target_library", rid, outcome))
        if not accepted or key in aliases:
            for peer, candidate in batch:
                if peer == key or (accepted and peer != aliases[key]):
                    continue
                outcome = compare(identity, candidate)
                if outcome != "different":
                    candidates.append(comparison(identity, candidate, "batch", peer, outcome))
        row = {
            "id": key, "source": "existing_citation" if old is not None else "selected_pdf",
            "normalized_fields": asdict(identity), "candidates": candidates,
        }
        if old is not None:
            row.update(original_record_id=old, original_library_ids=item["library_ids"])
        if not accepted:
            row["decision"] = "unresolved_batch"
        elif key in resolved:
            row.update(decision="reuse", target_record_id=resolved[key])
        elif aliases[key] == key:
            row["decision"] = "new"
        else:
            row.update(decision="alias", canonical_input=aliases[key])
        if accepted and old is not None:
            row["requires_rebind"] = old != resolved.get(key) or item["library_ids"] != [library_id]
        rows.append(row)
    return {
        "status": "PASS" if accepted else "REVIEW", "scope": "import_preview",
        "review_reasons": issues, "rows": rows, "active_library_records": len(records),
        "target_library_id": library_id,
        "planned_new_works": sum(row["decision"] == "new" for row in rows) if accepted else None,
    }


def make_preview(args):
    source = args.enl.resolve(strict=True)
    docx = args.docx.resolve(strict=True)
    if source.suffix.casefold() != ".enl" or docx.suffix.casefold() != ".docx":
        raise ValueError("Expected an .enl and a .docx input")
    database = sdb_path_for_enl(source)
    work = check_work_path(args.work_dir)
    require_closed_endnote()
    hashes = {path: sha256_file(path) for path in (docx, source, database, args.catalog, args.plan)}
    plan = json.loads(args.plan.read_text(encoding="utf-8"))
    pdf_ids = [key for placement in plan["insertions"] for key in placement["pdf_ids"]]
    groups = read_citations(docx)
    existing = [row for group in groups for row in group.records]
    work.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="preview-", dir=work) as scratch:
        selected = selected_pdfs(args.catalog, pdf_ids, Path(scratch))
        evidence = verify_plan_evidence(plan, selected)
        snapshot = copy_sdb_snapshot(source, Path(scratch) / "library.eni")
        if sha256_file(snapshot) != hashes[database]:
            raise RuntimeError("Library changed while taking the snapshot")
        report = forecast(existing, selected, read_endnote_active_records(snapshot), read_library_id(snapshot))
    require_closed_endnote()
    for suffix in ("-wal", "-journal"):
        sidecar = database.with_name(database.name + suffix)
        if sidecar.exists() and sidecar.stat().st_size:
            raise RuntimeError("Source has an uncheckpointed database after preview")
    if any(sha256_file(path) != digest for path, digest in hashes.items()):
        raise RuntimeError("A source changed during the preview")
    report["verified_pages"] = evidence
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("docx", "enl", "catalog", "plan", "work-dir", "output"):
        parser.add_argument("--" + name, type=Path, required=True)
    args = parser.parse_args(argv)
    output = check_work_path(args.output, ".json")
    if output.exists():
        parser.error("Use a new report path")
    try:
        report = make_preview(args)
    except (KeyError, TypeError, ValueError, RuntimeError, OSError, sqlite3.Error, BadZipFile) as error:
        report = {"status": "REVIEW", "scope": "import_preview", "review_reasons": [str(error)],
                  "rows": [], "planned_new_works": None}
    write_json_atomic(output, report)
    print(json.dumps({"status": report["status"], "planned_new_works": report["planned_new_works"]}))
    return 0 if report["status"] == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
