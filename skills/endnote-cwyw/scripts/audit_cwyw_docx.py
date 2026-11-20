# /// script
# requires-python = ">=3.10"
# dependencies = []
# ///

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
W = f"{{{W_NS}}}"
TEMPORARY_CITATION_RE = re.compile(
    r"\{[^{}#\r\n]+?#\d+(?:\s*;[^{}#\r\n]*?#\d+)*\}"
)
BIBLIOGRAPHY_HEADINGS = {
    "references",
    "bibliography",
    "works cited",
    "literature cited",
    "参考文献",
}


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def read_package(path: Path) -> tuple[list[str], dict[str, bytes], str | None]:
    with zipfile.ZipFile(path, "r") as archive:
        names = archive.namelist()
        bad_member = archive.testzip()
        members = {name: archive.read(name) for name in names}
    return names, members, bad_member


def relationship_target(rels_name: str, target: str) -> str:
    if rels_name == "_rels/.rels":
        base = ""
    else:
        source_name = rels_name.replace("/_rels/", "/")[: -len(".rels")]
        base = posixpath.dirname(source_name)
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base, target))


def validate_relationships(members: dict[str, bytes]) -> list[dict[str, str]]:
    missing: list[dict[str, str]] = []
    for name, data in members.items():
        if not name.endswith(".rels"):
            continue
        root = ET.fromstring(data)
        for relationship in root.findall(f"{{{REL_NS}}}Relationship"):
            if relationship.get("TargetMode") == "External":
                continue
            target = relationship_target(name, relationship.get("Target", ""))
            if target not in members:
                missing.append(
                    {
                        "rels": name,
                        "id": relationship.get("Id", ""),
                        "target": target,
                    }
                )
    return missing


def decode_field_data(text: str) -> str:
    if not text.strip():
        return ""
    try:
        return base64.b64decode("".join(text.split()), validate=True).decode(
            "utf-8", errors="replace"
        )
    except Exception:
        return ""


def record_ids(text: str) -> set[int]:
    values = re.findall(r"<(?:RecNum|rec-number)>(\d+)</(?:RecNum|rec-number)>", text)
    return {int(value) for value in values}


def parse_fields(document_xml: bytes) -> dict[str, object]:
    root = ET.fromstring(document_xml)
    stack: list[dict[str, object]] = []
    completed: list[dict[str, object]] = []
    visible_parts: list[str] = []
    deleted_parts: list[str] = []
    paragraph_index = 0

    for paragraph in root.iter(W + "p"):
        paragraph_index += 1
        for run in paragraph.iter(W + "r"):
            run_properties = run.find(W + "rPr")
            superscript = False
            if run_properties is not None:
                align = run_properties.find(W + "vertAlign")
                superscript = align is not None and align.get(W + "val") == "superscript"
            for child in list(run):
                if child.tag == W + "fldChar":
                    kind = child.get(W + "fldCharType")
                    if kind == "begin":
                        encoded = "".join(
                            node.text or "" for node in child.findall(".//" + W + "fldData")
                        )
                        stack.append(
                            {
                                "code_parts": [],
                                "data": decode_field_data(encoded),
                                "separated": False,
                                "result_parts": [],
                                "result_paragraphs": set(),
                                "result_paragraph_texts": {},
                                "result_superscript": [],
                            }
                        )
                    elif kind == "separate" and stack:
                        stack[-1]["separated"] = True
                    elif kind == "end" and stack:
                        completed.append(stack.pop())
                elif child.tag == W + "instrText" and stack:
                    stack[-1]["code_parts"].append(child.text or "")
                elif child.tag == W + "t":
                    text = child.text or ""
                    visible_parts.append(text)
                    for field in stack:
                        if field["separated"]:
                            field["result_parts"].append(text)
                            field["result_paragraph_texts"].setdefault(
                                paragraph_index, []
                            ).append(text)
                            if text.strip():
                                field["result_paragraphs"].add(paragraph_index)
                                field["result_superscript"].append(superscript)
                elif child.tag == W + "delText":
                    deleted_parts.append(child.text or "")

    fields = []
    for field in completed:
        code = "".join(field["code_parts"]).strip()
        data = str(field["data"])
        metadata = code + data
        ids = sorted(record_ids(metadata))
        ordered_paragraphs = [
            "".join(parts).strip()
            for _, parts in sorted(field["result_paragraph_texts"].items())
            if "".join(parts).strip()
        ]
        fields.append(
            {
                "code": code,
                "result": "".join(field["result_parts"]),
                "result_paragraph_count": len(field["result_paragraphs"]),
                "result_paragraphs_ordered": ordered_paragraphs,
                "all_nonempty_result_runs_superscript": bool(field["result_superscript"])
                and all(field["result_superscript"]),
                "has_endnote_xml": "<EndNote>" in metadata
                and "<record>" in metadata,
                "record_ids": ids,
            }
        )

    main = [
        field
        for field in fields
        if field["code"].startswith("ADDIN EN.CITE")
        and not field["code"].startswith("ADDIN EN.CITE.DATA")
    ]
    data_fields = [field for field in fields if field["code"] == "ADDIN EN.CITE.DATA"]
    reference_lists = [field for field in fields if field["code"] == "ADDIN EN.REFLIST"]
    bibliography_heading_detected = False
    bibliography_entry_count = 0
    if reference_lists:
        ref_paragraphs = list(reference_lists[0]["result_paragraphs_ordered"])
        if ref_paragraphs and ref_paragraphs[0].strip().casefold() in BIBLIOGRAPHY_HEADINGS:
            bibliography_heading_detected = True
            bibliography_entry_count = max(len(ref_paragraphs) - 1, 0)
        else:
            bibliography_entry_count = reference_lists[0]["result_paragraph_count"]
    visible_text = "".join(visible_parts)
    deleted_text = "".join(deleted_parts)
    ids_from_fields = sorted(
        {record_id for field in main for record_id in field["record_ids"]}
    )
    return {
        "balanced_field_stack": not stack,
        "all_fields": fields,
        "main_citation_fields": len(main),
        "citation_data_fields": len(data_fields),
        "reference_list_fields": len(reference_lists),
        "bibliography_heading_detected": bibliography_heading_detected,
        "bibliography_entry_count": bibliography_entry_count,
        "citation_results": [field["result"] for field in main],
        "main_fields_with_endnote_metadata": sum(
            bool(field["has_endnote_xml"]) for field in main
        ),
        "all_main_citations_superscript": bool(main)
        and all(field["all_nonempty_result_runs_superscript"] for field in main),
        "record_ids_from_fields": ids_from_fields,
        "temporary_citations_remaining": len(TEMPORARY_CITATION_RE.findall(visible_text)),
        "deleted_temporary_citations": len(TEMPORARY_CITATION_RE.findall(deleted_text)),
    }


def count_notes_temporary_citations(members: dict[str, bytes]) -> int:
    total = 0
    for note_member in ("word/footnotes.xml", "word/endnotes.xml"):
        if note_member not in members:
            continue
        root = ET.fromstring(members[note_member])
        para_texts = []
        for paragraph in root.iter(W + "p"):
            parts = []
            for run in paragraph.iter(W + "r"):
                for child in list(run):
                    if child.tag == W + "t":
                        parts.append(child.text or "")
            if parts:
                para_texts.append("".join(parts))
        total += len(TEMPORARY_CITATION_RE.findall("\n".join(para_texts)))
    return total


def restore_media_from_base(docx: Path, base: Path, out_path: Path) -> list[str]:
    if out_path.exists():
        raise FileExistsError(f"Refusing to overwrite restored output: {out_path}")
    if out_path.resolve() in {docx.resolve(), base.resolve()}:
        raise RuntimeError("Restored output path must differ from docx and base")
    _, base_members, base_bad = read_package(base)
    if base_bad is not None:
        raise RuntimeError(f"Base DOCX has corrupt zip member: {base_bad}")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    restored: list[str] = []
    with zipfile.ZipFile(docx, "r") as src_zip, zipfile.ZipFile(
        out_path, "w", compression=zipfile.ZIP_DEFLATED
    ) as dst_zip:
        for info in src_zip.infolist():
            data = src_zip.read(info.filename)
            if (
                info.filename.startswith("word/media/")
                and info.filename in base_members
                and data != base_members[info.filename]
            ):
                data = base_members[info.filename]
                restored.append(info.filename)
            new_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            new_info.compress_type = zipfile.ZIP_DEFLATED
            dst_zip.writestr(new_info, data)
    return restored


def parse_settings(settings_xml: bytes) -> dict[str, object]:
    root = ET.fromstring(settings_xml)
    variables = {
        node.get(W + "name", ""): node.get(W + "val", "")
        for node in root.iter(W + "docVar")
    }
    layout = variables.get("EN.Layout", "")
    libraries = variables.get("EN.Libraries", "")
    style_match = re.search(r"<Style>(.*?)</Style>", layout, flags=re.DOTALL)
    ids = sorted({int(value) for value in re.findall(r"<item>(\d+)</item>", libraries)})
    return {
        "style_name": style_match.group(1) if style_match else None,
        "record_ids_from_settings": ids,
        "has_endnote_layout": bool(layout),
        "has_endnote_libraries": bool(libraries),
    }


def parse_expected_ids(value: str | None) -> list[int] | None:
    if value is None:
        return None
    ids = sorted({int(item.strip()) for item in value.split(",") if item.strip()})
    if not ids:
        raise ValueError("--expect-record-ids did not contain any IDs")
    return ids


def audit(
    docx: Path,
    *,
    base: Path | None = None,
    expect_main: int | None = None,
    expect_data: int | None = None,
    expect_references: int | None = None,
    expect_record_ids: list[int] | None = None,
    expect_style: str | None = None,
    require_superscript: bool = False,
) -> dict[str, object]:
    names, members, bad_member = read_package(docx)
    required = {"word/document.xml", "word/settings.xml"}
    missing_required = sorted(required - set(members))
    xml_errors: dict[str, str] = {}
    for name, data in members.items():
        if name.endswith(".xml") or name.endswith(".rels"):
            try:
                ET.fromstring(data)
            except Exception as exc:
                xml_errors[name] = f"{type(exc).__name__}: {exc}"
    missing_relationships = validate_relationships(members) if not xml_errors else []
    fields = (
        parse_fields(members["word/document.xml"])
        if "word/document.xml" in members
        else {}
    )
    notes_temp_count = count_notes_temporary_citations(members) if not xml_errors else 0
    if fields:
        fields["notes_temporary_citations_remaining"] = notes_temp_count
    settings = (
        parse_settings(members["word/settings.xml"])
        if "word/settings.xml" in members
        else {}
    )

    field_ids = fields.get("record_ids_from_fields", [])
    settings_ids = settings.get("record_ids_from_settings", [])
    checks: dict[str, bool] = {
        "zip_crc": bad_member is None,
        "required_members": not missing_required,
        "xml_parse": not xml_errors,
        "relationships": not missing_relationships,
        "balanced_fields": bool(fields.get("balanced_field_stack")),
        "has_main_citations": int(fields.get("main_citation_fields", 0)) > 0,
        "one_reference_list": fields.get("reference_list_fields") == 1,
        "no_temporary_citations": (
            fields.get("temporary_citations_remaining") == 0 and notes_temp_count == 0
        ),
        "all_main_fields_have_metadata": fields.get("main_fields_with_endnote_metadata")
        == fields.get("main_citation_fields")
        and int(fields.get("main_citation_fields", 0)) > 0,
        "field_and_settings_record_ids_match": field_ids == settings_ids and bool(field_ids),
    }
    if expect_main is not None:
        checks["expected_main_citations"] = fields.get("main_citation_fields") == expect_main
    if expect_data is not None:
        checks["expected_data_fields"] = fields.get("citation_data_fields") == expect_data
    if expect_references is not None:
        checks["expected_references"] = fields.get("bibliography_entry_count") == expect_references
    if expect_record_ids is not None:
        checks["expected_field_record_ids"] = field_ids == expect_record_ids
        checks["expected_settings_record_ids"] = settings_ids == expect_record_ids
    if expect_style is not None:
        checks["expected_style"] = settings.get("style_name") == expect_style
    if require_superscript:
        checks["all_main_citations_superscript"] = bool(
            fields.get("all_main_citations_superscript")
        )

    package_diff: dict[str, object] | None = None
    if base is not None:
        base_names, base_members, base_bad_member = read_package(base)
        shared = sorted(set(members) & set(base_members))
        changed = [name for name in shared if members[name] != base_members[name]]
        media_names = sorted(name for name in members if name.startswith("word/media/"))
        media_mismatch = [
            name
            for name in media_names
            if name not in base_members or members[name] != base_members[name]
        ]
        package_diff = {
            "base": str(base),
            "base_sha256": sha256_file(base),
            "base_zip_crc_member": base_bad_member,
            "member_order_equal": names == base_names,
            "added_members": sorted(set(members) - set(base_members)),
            "removed_members": sorted(set(base_members) - set(members)),
            "changed_members": changed,
            "media_count": len(media_names),
            "media_mismatch": media_mismatch,
        }

    public_fields = dict(fields)
    public_fields.pop("all_fields", None)
    return {
        "status": "PASS" if checks and all(checks.values()) else "FAIL",
        "docx": str(docx),
        "sha256": sha256_file(docx),
        "package_member_count": len(names),
        "bad_zip_member": bad_member,
        "missing_required_members": missing_required,
        "xml_errors": xml_errors,
        "missing_relationship_targets": missing_relationships,
        "fields": public_fields,
        "settings": settings,
        "checks": checks,
        "package_diff": package_diff,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Audit genuine EndNote CWYW fields and DOCX package integrity."
    )
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--base", type=Path)
    parser.add_argument("--expect-main", type=int)
    parser.add_argument("--expect-data", type=int)
    parser.add_argument("--expect-references", type=int)
    parser.add_argument("--expect-record-ids")
    parser.add_argument("--expect-style")
    parser.add_argument("--require-superscript", action="store_true")
    parser.add_argument("--restore-media-out", type=Path)
    parser.add_argument("--report", type=Path)
    return parser


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    args = build_parser().parse_args()
    docx = args.docx.resolve()
    base = args.base.resolve() if args.base else None
    if not docx.is_file():
        raise FileNotFoundError(docx)
    if base is not None and not base.is_file():
        raise FileNotFoundError(base)
    if args.restore_media_out is not None and base is None:
        raise RuntimeError("--restore-media-out requires --base")
    expected_ids = parse_expected_ids(args.expect_record_ids)
    result = audit(
        docx,
        base=base,
        expect_main=args.expect_main,
        expect_data=args.expect_data,
        expect_references=args.expect_references,
        expect_record_ids=expected_ids,
        expect_style=args.expect_style,
        require_superscript=args.require_superscript,
    )
    if args.restore_media_out is not None:
        mismatched = list((result.get("package_diff") or {}).get("media_mismatch") or [])
        if mismatched:
            restore_out = args.restore_media_out.resolve()
            restored = restore_media_from_base(docx, base, restore_out)
            result["restored_media_docx"] = str(restore_out)
            result["restored_media_files"] = restored
            result["restored_audit"] = audit(
                restore_out,
                base=base,
                expect_main=args.expect_main,
                expect_data=args.expect_data,
                expect_references=args.expect_references,
                expect_record_ids=expected_ids,
                expect_style=args.expect_style,
                require_superscript=args.require_superscript,
            )
        else:
            result["restored_media_files"] = []
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        report = args.report.resolve()
        if report.exists():
            raise FileExistsError(f"Refusing to overwrite report: {report}")
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    return 0 if result["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())

