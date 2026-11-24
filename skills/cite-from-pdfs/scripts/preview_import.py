# /// script
# requires-python = ">=3.10"
# dependencies = ["lxml>=5", "PyMuPDF>=1.24", "python-docx>=1.1"]
# ///
"""Preview a citation import against a closed library. The JSON report is private."""

import argparse
from dataclasses import asdict
import json
import os
from pathlib import Path
import sqlite3
import sys
import tempfile
from zipfile import BadZipFile, ZipFile

from _common import (check_work_path, copy_sdb_snapshot, read_endnote_active_records,
                     read_library_id, sdb_path_for_enl, write_json_atomic)
from placement_evidence import verify_plan_evidence
from inject_placeholders import locate_placements, read_manuscript
from reference_identity import Identity, IdentityReview, RecordIndex, compare
from sync_to_endnote import plan_import, require_closed_endnote, selected_pdfs, sha256_file

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "endnote-cwyw" / "scripts"))
from audit_cwyw_docx import parse_settings
from rebind_cwyw import read_citations


def _configure_utf8_streams():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def _write_text_atomic_new(path: Path, text_value: str) -> None:
    if path.exists():
        raise FileExistsError(f"Refusing to overwrite existing report: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=path.parent, prefix=f".{path.name}.", suffix=".tmp", delete=False
        ) as handle:
            tmp_path = handle.name
            handle.write(text_value)
        if os.name == "nt":
            os.rename(tmp_path, path)
        else:
            os.link(tmp_path, path)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except OSError:
                pass


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
        for item in items:
            try:
                index.match(item, preferred)
            except IdentityReview as item_err:
                if str(item_err) not in issues:
                    issues.append(str(item_err))
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
    if not isinstance(plan, dict) or not isinstance(plan.get("insertions"), list):
        raise ValueError("Citation plan must be a JSON object containing an 'insertions' list")
    raw_insertions = plan["insertions"]
    insertions = [
        {**item, "insertion_id": item.get("insertion_id") or f"C{idx:02d}"}
        if isinstance(item, dict)
        else item
        for idx, item in enumerate(raw_insertions, 1)
    ]

    _, para_maps = read_manuscript(docx)
    with ZipFile(docx) as package:
        settings = parse_settings(package.read("word/settings.xml")) if "word/settings.xml" in package.namelist() else {}
    style = plan.get("style") or settings.get("style_name")

    # Validate each insertion individually so all bad/ambiguous anchors surface together
    placements = []
    placement_errors = []
    valid_items = []
    for idx, item in enumerate(insertions, 1):
        label = f"C{idx:02d}"
        if not isinstance(item, dict):
            placement_errors.append(f"{label}: insertion entry must be a JSON object")
            continue
        try:
            placements.extend(locate_placements(para_maps, [item], style=style))
            valid_items.append(item)
        except (ValueError, KeyError, TypeError, AttributeError) as err:
            placement_errors.append(str(err))
    if valid_items:
        try:
            placements = locate_placements(para_maps, valid_items, style=style)
        except (ValueError, KeyError, TypeError, AttributeError) as err:
            placement_errors.append(str(err))

    groups = read_citations(docx)
    existing = [row for group in groups for row in group.records]
    work.mkdir(parents=True, exist_ok=True)
    evidence = []
    evidence_errors = []
    selected_by_id: dict[str, dict] = {}
    with tempfile.TemporaryDirectory(prefix="preview-", dir=work) as scratch:
        for idx, ins in enumerate(insertions, 1):
            label = ins.get("insertion_id", f"C{idx:02d}") if isinstance(ins, dict) else f"C{idx:02d}"
            if not isinstance(ins, dict):
                continue
            raw_pids = ins.get("pdf_ids")
            if not isinstance(raw_pids, list) or not raw_pids:
                evidence_errors.append(f"{label}: supply a non-empty 'pdf_ids' list")
                continue
            try:
                ins_pids = [str(pid) for pid in raw_pids]
                ins_scratch = Path(scratch) / f"ins_{idx}"
                ins_scratch.mkdir()
                ins_selected = selected_pdfs(args.catalog, ins_pids, ins_scratch)
                for sel in ins_selected:
                    selected_by_id.setdefault(sel["pdf_id"], sel)
                evidence.extend(verify_plan_evidence({"insertions": [ins]}, ins_selected))
            except (ValueError, KeyError, TypeError, AttributeError, FileNotFoundError, RuntimeError) as err:
                msg = str(err)
                evidence_errors.append(msg if msg.startswith(label) else f"{label}: {msg}")
        selected = list(selected_by_id.values())
        snapshot = copy_sdb_snapshot(source, Path(scratch) / "library.eni")
        if sha256_file(snapshot) != hashes[database]:
            raise RuntimeError("Library changed while taking the snapshot")
        if existing or selected:
            report = forecast(existing, selected, read_endnote_active_records(snapshot), read_library_id(snapshot))
        else:
            report = {"status": "REVIEW", "scope": "import_preview", "review_reasons": ["No existing citations or selected PDFs"], "rows": [], "active_library_records": len(read_endnote_active_records(snapshot)), "target_library_id": read_library_id(snapshot), "planned_new_works": None}
    require_closed_endnote()
    for suffix in ("-wal", "-journal"):
        sidecar = database.with_name(database.name + suffix)
        if sidecar.exists() and sidecar.stat().st_size:
            raise RuntimeError("Source has an uncheckpointed database after preview")
    if any(sha256_file(path) != digest for path, digest in hashes.items()):
        raise RuntimeError("A source changed during the preview")

    all_errors = placement_errors + evidence_errors
    if all_errors:
        report["status"] = "REVIEW"
        report["planned_new_works"] = None
        report["review_reasons"] = list(dict.fromkeys(all_errors + list(report.get("review_reasons") or [])))
    report["verified_pages"] = evidence
    report["placements"] = [
        {"insertion_id": op["insertion_id"], "anchor_text": op["anchor_text"],
         "paragraph": op["paragraph_index"] + 1, "occurrence": op["occurrence"],
         "position": op["position"], "inline_preview": op.get("inline_preview", op["anchor_text"]),
         "pdf_ids": op["pdf_ids"]} for op in placements
    ]
    return report


def explain(report, language):
    zh = language == "zh"
    if report["status"] != "PASS":
        lines = ["需要复核。请先处理下列问题，再预演一次。" if zh else
                 "Review required. Resolve these items and preview again."]
        lines.extend(report["review_reasons"])
        for row in report.get("rows") or []:
            for candidate in row.get("candidates") or []:
                if candidate.get("different_fields"):
                    differences = ", ".join(
                        f"{field}: {row['normalized_fields'][field]} / {candidate['normalized_fields'][field]}"
                        for field in candidate["different_fields"])
                    lines.append(f"{row['id']} ↔ {candidate['source']} {candidate['id']} ({differences})")
        return "\n".join(lines)
    count = report["planned_new_works"]
    lines = [f"预演通过，预计新增 {count} 篇文献。" if zh else f"Preview passed. New works planned: {count}."]
    for row in report["rows"]:
        if row["decision"] == "reuse":
            action = ("复用库中记录" if zh else "reuse library record") + f" #{row['target_record_id']}"
        elif row["decision"] == "alias":
            action = ("与此文件同篇" if zh else "same work as") + f" {row['canonical_input']}"
        else:
            action = "新增" if zh else "add"
        lines.append(f"{row['id']} · {row['normalized_fields']['title']} · {action}")
    for placement in report.get("placements") or []:
        label = "正文段落" if zh else "Body paragraph"
        lines.append(f"{label} {placement['paragraph']} · {placement.get('inline_preview', placement['anchor_text'])}")
        for page in report.get("verified_pages") or []:
            if page["insertion_id"] == placement["insertion_id"]:
                tot = page.get("total_pages")
                pg_str = (
                    (f"PDF 第 {page['page']} 页（共 {tot} 页）" if tot else f"PDF 第 {page['page']} 页")
                    if zh else (f"PDF page {page['page']} of {tot}" if tot else f"PDF page {page['page']}")
                )
                quote = str(page.get("quote") or "").strip()
                q_suffix = f" · 摘录：{quote[:90]}" if (zh and quote) else (f" · Quote: {quote[:90]}" if quote else "")
                lines.append(f"  {page['pdf_id']} · {pg_str}{q_suffix}")
    return "\n".join(lines)


def _format_markdown_report(report: dict) -> str:
    lines = [
        "# Citation Preview Report",
        "",
        f"- **Status**: `{report.get('status')}`",
        f"- **Planned new library entries**: `{report.get('planned_new_works')}`",
        "",
    ]
    if report.get("review_reasons"):
        lines.extend(["## Review Required", ""])
        lines.extend(f"- {reason}" for reason in report["review_reasons"])
        lines.append("")
    if report.get("placements"):
        lines.extend(["## Placements & Evidence", ""])
        pages_by_ins: dict[str, list[dict]] = {}
        for vp in report.get("verified_pages") or []:
            pages_by_ins.setdefault(vp["insertion_id"], []).append(vp)
        for pl in report["placements"]:
            lines.append(f"### {pl['insertion_id']} (Paragraph {pl['paragraph']}, position `{pl.get('position', 'auto')}`)")
            lines.append(f"- Preview: `{pl.get('inline_preview', pl['anchor_text'])}`")
            for vp in pages_by_ins.get(pl["insertion_id"], []):
                lines.append(f"- Source `{vp['pdf_id']}` (p. {vp['page']}/{vp.get('total_pages', '?')}): \"{vp.get('quote', '')}\"")
            lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def main(argv=None):
    _configure_utf8_streams()
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("docx", "enl", "catalog", "plan", "work-dir", "output"):
        parser.add_argument("--" + name, type=Path, required=True)
    parser.add_argument("--explain", choices=("zh", "en"), help="Show a private, readable review with manuscript text in this terminal.")
    parser.add_argument("--report-md", type=Path, help="Optional Markdown preview path for human review.")
    args = parser.parse_args(argv)
    output = check_work_path(args.output, ".json")
    if output.exists():
        parser.error("Use a new report path")
    report_md = check_work_path(args.report_md, ".md") if args.report_md else None
    if report_md and report_md.exists():
        parser.error("Use a new Markdown report path")
    try:
        report = make_preview(args)
    except (KeyError, TypeError, ValueError, AttributeError, IndexError, RuntimeError, OSError, sqlite3.Error, BadZipFile, SyntaxError) as error:
        report = {"status": "REVIEW", "scope": "import_preview", "review_reasons": [str(error)],
                  "rows": [], "planned_new_works": None}
    write_json_atomic(output, report)
    if report_md:
        _write_text_atomic_new(report_md, _format_markdown_report(report))
    if args.explain:
        print(explain(report, args.explain))
    else:
        print(json.dumps({"status": report["status"], "planned_new_works": report["planned_new_works"]}))
    return 0 if report["status"] == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
