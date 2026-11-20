# /// script
# requires-python = ">=3.10"
# dependencies = ["lxml>=5.0.0", "python-docx>=1.1.0", "pywin32>=306"]
# ///

"""Apply an evidence-backed selection, publishing an updated library copy and genuine CWYW DOCX."""

from __future__ import annotations

import argparse
from contextlib import closing
import hashlib
from itertools import zip_longest
import json
from pathlib import Path
import shutil
import sqlite3
import sys
import tempfile
from zipfile import ZipFile

from _common import check_work_path, copy_sdb_snapshot, read_endnote_active_records, read_library_id, sdb_path_for_enl
from inject_placeholders import inject_into_docx
from reference_identity import Identity, RecordIndex, compare
from sync_to_endnote import require_closed_endnote, selected_pdfs, sha256_file, sync_candidate, tree_hashes, verify_library_contents


CWYW = Path(__file__).resolve().parents[2] / "endnote-cwyw" / "scripts"
sys.path.insert(0, str(CWYW))
from audit_cwyw_docx import audit, parse_settings, restore_media_from_base
from rebind_cwyw import document_citations, prepare_rebind, read_citations


EQUIVALENT = {"doi", "pmid", "title_author_year", "title_variant"}


def write_json(path, data):
    with path.open("x", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def package_media(path):
    with ZipFile(path) as package:
        return {
            name: package.read(name) for name in package.namelist()
            if name.startswith("word/media/")
        }


def check_delivery(docx, source, sync, positions, candidate_snapshot, style, require_superscript):
    record_ids = sorted({rid for _, _, group in positions for rid in group})
    result = audit(
        docx, base=source, expect_main=len(positions), expect_references=len(record_ids),
        expect_record_ids=record_ids, expect_style=style, require_superscript=require_superscript,
    )
    if result["status"] != "PASS":
        raise RuntimeError(f"Final CWYW audit failed: {[key for key, ok in result['checks'].items() if not ok]}")
    records = read_endnote_active_records(candidate_snapshot)
    final_groups, actual = document_citations(docx)
    for group in final_groups:
        for row in group.records:
            rid = row["original_id"]
            if row["library_ids"] != [sync["library_id"]] or rid not in records:
                raise RuntimeError(f"Final Word record #{rid} is not bound to the delivered library")
            if compare(Identity.from_record(row), Identity.from_record(records[rid])) not in EQUIVALENT:
                raise RuntimeError(f"Final Word metadata for #{rid} differs from the delivered library")
    if actual != positions:
        raise RuntimeError("Final CWYW positions or record groups differ from the verified plan")
    return result


def apply(args, work, stage):
    require_closed_endnote()
    original_docx = args.docx.resolve(strict=True)
    source = work / "frozen.docx"
    source_enl = args.enl.resolve(strict=True)
    source_hash = sha256_file(original_docx)
    shutil.copyfile(original_docx, source)
    if source_hash != sha256_file(source):
        raise RuntimeError("Word source changed while it was copied")
    plan = json.loads(args.plan.read_text(encoding="utf-8"))
    with ZipFile(source) as package:
        settings = parse_settings(package.read("word/settings.xml")) if "word/settings.xml" in package.namelist() else {}
    style = plan.get("style") or settings.get("style_name")
    if not style:
        raise ValueError("Supply the required EndNote output style in the citation plan")
    plan["style"] = style
    require_superscript = args.require_superscript or style.casefold() == "nature"
    plan.setdefault("space_before_citation", True)
    groups = read_citations(source)
    existing = [row for group in groups for row in group.records]
    selected = selected_pdfs(
        args.catalog, [pdf_id for insertion in plan["insertions"] for pdf_id in insertion["pdf_ids"]],
    )
    if not existing and not selected:
        raise ValueError("No genuine or proposed citations to format")

    frozen = copy_sdb_snapshot(source_enl, work / "planning.eni")
    signature = (sha256_file(source_enl), sha256_file(frozen))
    library_id = read_library_id(frozen)
    preferred = {row["original_id"] for row in existing if row["library_ids"] == [library_id]}
    source_index = read_endnote_active_records(frozen)
    index = RecordIndex(source_index)
    for row in existing:
        if index.match(row, preferred) is None and row["reference_type"] != "Journal Article":
            raise ValueError(f"Existing non-journal record requires explicit migration: {row['title']}")
    sync = sync_candidate(
        source_enl, stage / "library", existing + selected, work, preferred, signature,
    )
    write_json(work / "resolved.json", sync)
    rebound = work / "rebound.docx"
    rebinding = prepare_rebind(source, rebound, groups, sync)
    base = rebound if rebinding else source
    pre_cwyw = work / "prepared.docx"
    if plan["insertions"]:
        injection = inject_into_docx(base, pre_cwyw, plan, sync)
        write_json(work / "injection.json", injection)
    else:
        pre_cwyw = base
    _, expected_positions = document_citations(pre_cwyw, prepared=True)
    if len(expected_positions) != len(groups) + len(plan["insertions"]):
        raise RuntimeError("Prepared Word does not contain the expected number of citation placements")
    record_ids = sorted({rid for _, _, group in expected_positions for rid in group})
    candidate = Path(sync["enl_path"])
    before_format = work / ("after.eni" if sync["new_record_ids"] else "before.eni")
    candidate_pdf_hashes = tree_hashes(candidate.with_suffix(".Data") / "PDF")
    native_input = work.parent / f"{work.name}-input.docx"
    with native_input.open("xb") as target, pre_cwyw.open("rb") as original:
        shutil.copyfileobj(original, target)
    native_hash = sha256_file(native_input)
    formatted = work.parent / f"{work.name}-word.docx"
    command = [
        sys.executable, "-X", "utf8", str(CWYW / "format_cwyw.py"),
        "--enl", str(candidate), "--input", str(native_input), "--output", str(formatted),
        "--style", style, "--expect-main", str(len(expected_positions)),
        "--expect-references", str(len(record_ids)), "--expect-record-ids", ",".join(map(str, record_ids)),
        "--expected-source-sha256", native_hash,
        "--timeout", str(args.timeout), "--report", str(work / "format.json"),
    ]
    if require_superscript:
        command.append("--require-superscript")
    ready = {
        "original_docx": str(original_docx), "source_enl": str(source_enl),
        "source_hash": source_hash, "signature": signature, "library_id": library_id,
        "stage": str(stage), "destination": str(args.output_dir.resolve()),
        "candidate": str(candidate.relative_to(stage)),
        "native_input_sha256": native_hash, "style": style,
        "require_superscript": require_superscript, "positions": expected_positions,
        "rebound_existing_citations": rebinding, "candidate_pdf_hashes": candidate_pdf_hashes,
        "before_format": before_format.name, "before_format_sha256": sha256_file(before_format),
        "resolved_sha256": sha256_file(work / "resolved.json"),
        "format_command": command,
    }
    write_json(work / "ready.json", ready)
    return ready


def verify_record_table(before, after):
    with (
        closing(sqlite3.connect(before.as_uri() + "?mode=ro&immutable=1", uri=True)) as initial,
        closing(sqlite3.connect(after.as_uri() + "?mode=ro&immutable=1", uri=True)) as final,
    ):
        sql = "SELECT * FROM refs ORDER BY id"
        if any(left != right for left, right in zip_longest(initial.execute(sql), final.execute(sql))):
            raise RuntimeError("Word use changed the EndNote library's reference table")


def finalize(work, ready):
    require_closed_endnote()
    stage = check_work_path(Path(ready["stage"]))
    destination = check_work_path(Path(ready["destination"]))
    if stage.parent != destination.parent or not stage.name.startswith(f".{destination.name}.pending-"):
        raise RuntimeError("Pending library directory does not belong to this release")
    if destination.exists():
        raise FileExistsError(destination)
    candidate = (stage / ready["candidate"]).resolve(strict=True)
    if candidate.parent != stage / "library" or candidate.suffix.casefold() != ".enl":
        raise RuntimeError("Candidate library path escaped staging")
    original_docx = Path(ready["original_docx"]).resolve(strict=True)
    source_enl = Path(ready["source_enl"]).resolve(strict=True)
    if any(path.is_relative_to(stage) for path in (original_docx, source_enl)):
        raise RuntimeError("An original input is inside the proposed delivery")
    source = work / "frozen.docx"
    source_hash = ready["source_hash"]
    signature = tuple(ready["signature"])
    library_id = ready["library_id"]
    resolution = work / "resolved.json"
    if sha256_file(resolution) != ready["resolved_sha256"]:
        raise RuntimeError("Prepared reference resolution changed")
    sync = json.loads(resolution.read_text(encoding="utf-8"))
    if sync["library_id"] != library_id or Path(sync["enl_path"]).resolve() != candidate:
        raise RuntimeError("Resolved library does not match the prepared session")
    if ready["before_format"] not in {"after.eni", "before.eni"}:
        raise RuntimeError("Invalid prepared library snapshot")
    before_format = work / ready["before_format"]
    if sha256_file(before_format) != ready["before_format_sha256"]:
        raise RuntimeError("Prepared candidate database snapshot changed")
    formatted = work.parent / f"{work.name}-word.docx"
    native_input = work.parent / f"{work.name}-input.docx"
    report = json.loads((work / "format.json").read_text(encoding="utf-8"))
    if (
        report["status"] != "PASS" or report["output_sha256"] != sha256_file(formatted)
        or report["preflight"]["source_sha256"] != sha256_file(native_input)
        or report["preflight"]["source_sha256"] != ready["native_input_sha256"]
        or Path(report["endnote_library"]["enl"]).resolve() != candidate
        or Path(report["preflight"]["source"]).resolve() != native_input
        or Path(report["preflight"]["output"]).resolve() != formatted
    ):
        raise RuntimeError("Word output is not the verified native result for this session")
    if source_hash != sha256_file(source) or source_hash != sha256_file(original_docx) or signature != (
        sha256_file(source_enl), sha256_file(sdb_path_for_enl(source_enl))
    ):
        raise RuntimeError("Original Word or EndNote source changed during the run")
    frozen_final = copy_sdb_snapshot(candidate, work / "final.eni")
    if read_library_id(frozen_final) != library_id:
        raise RuntimeError("Candidate library ID changed during formatting")
    verify_record_table(before_format, frozen_final)
    verify_library_contents(before_format, frozen_final, candidate, {}, ready["candidate_pdf_hashes"])
    final_docx = stage / "manuscript_cwyw.docx"
    source_media, generated_media = package_media(source), package_media(formatted)
    if source_media.keys() != generated_media.keys():
        raise RuntimeError("Word changed the set of embedded media")
    if source_media != generated_media:
        restore_media_from_base(formatted, source, final_docx)
    else:
        with final_docx.open("xb") as output, formatted.open("rb") as original:
            shutil.copyfileobj(original, output)
    if package_media(final_docx) != source_media:
        raise RuntimeError("Final Word does not retain original media bytes")
    expected_positions = [(p, offset, tuple(ids)) for p, offset, ids in ready["positions"]]
    proof = check_delivery(
        final_docx, source, sync, expected_positions, frozen_final,
        ready["style"], ready["require_superscript"],
    )
    if source_hash != sha256_file(original_docx) or signature != (
        sha256_file(source_enl), sha256_file(sdb_path_for_enl(source_enl))
    ):
        raise RuntimeError("Original Word or EndNote source changed before publication")
    proof["docx"] = final_docx.name
    write_json(stage / "audit.json", {
        "status": "PASS", "library": str(candidate.relative_to(stage)), "library_id": library_id,
        "new_record_ids": sync["new_record_ids"], "before_active_count": sync["before_active_count"],
        "after_active_count": sync["after_active_count"],
        "rebound_existing_citations": ready["rebound_existing_citations"],
        "source_sha256": source_hash, "source_library_signature": signature, "word": proof,
    })


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--enl", type=Path)
    parser.add_argument("--docx", type=Path)
    parser.add_argument("--catalog", type=Path)
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--work-dir", type=Path)
    parser.add_argument("--output-dir", type=Path)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--prepare-only", action="store_true")
    mode.add_argument("--publish-prepared", type=Path, help="Publish an audited Word result from a prepared work directory")
    parser.add_argument("--ready-sha256", help="The SHA-256 emitted during --prepare-only")
    parser.add_argument("--timeout", type=int, default=90)
    parser.add_argument("--require-superscript", action="store_true")
    args = parser.parse_args()
    if args.publish_prepared is not None:
        work = check_work_path(args.publish_prepared)
        receipt = (work / "ready.json").read_bytes()
        if not args.ready_sha256 or hashlib.sha256(receipt).hexdigest().upper() != args.ready_sha256.upper():
            parser.error("Preparation receipt did not match its independently retained SHA-256")
        ready = json.loads(receipt)
        destination = check_work_path(Path(ready["destination"]))
        finalize(work, ready)
        check_work_path(Path(ready["stage"])).rename(destination)
        print(json.dumps({"status": "PASS", "published": str(destination), "work": str(work)}, ensure_ascii=False))
        return 0
    if any(getattr(args, key) is None for key in ("enl", "docx", "catalog", "plan", "work_dir", "output_dir")):
        parser.error("Supply --enl, --docx, --catalog, --plan, --work-dir and --output-dir")
    destination = args.output_dir.resolve()
    work_root = args.work_dir.resolve()
    if destination.exists() or args.timeout < 10:
        parser.error("Use a new output directory and a timeout of at least 10 seconds")
    for path in (destination, work_root):
        if any(parent.name.casefold().endswith(".data") for parent in (path, *path.parents)):
            parser.error("Output and work directories must be outside EndNote .Data directories")
    work_root.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="citation-run-", dir=work_root))
    destination.parent.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix=f".{destination.name}.pending-", dir=destination.parent))
    try:
        ready = apply(args, work, stage)
    except Exception as exc:
        failure = {"status": "FAIL", "error": str(exc), "unpublished_candidate": str(stage), "work": str(work)}
        write_json(work / "failure.json", failure)
        print(json.dumps(failure, ensure_ascii=False), file=sys.stderr)
        return 1
    print(json.dumps({
        "status": "READY", "work": str(work), "format_command": ready["format_command"],
        "ready_sha256": sha256_file(work / "ready.json"),
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
