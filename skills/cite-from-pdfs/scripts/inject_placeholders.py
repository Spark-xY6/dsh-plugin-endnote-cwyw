# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "python-docx>=1.1.0",
#     "lxml>=5.0.0",
# ]
# ///

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree
from _common import check_work_path, write_json_atomic


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W = f"{{{W_NS}}}"
TERMINAL_PUNCT = set(".,;:!?。，；：！？")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def _audit_existing_cwyw(docx_path: Path) -> tuple[int, list[int]]:
    """Inspect existing genuine CWYW citations if any are already present in input DOCX."""
    sibling_scripts = (
        Path(__file__).resolve().parents[2] / "endnote-cwyw" / "scripts"
    )
    if sibling_scripts.is_dir() and str(sibling_scripts) not in sys.path:
        sys.path.insert(0, str(sibling_scripts))
    from audit_cwyw_docx import audit

    fields = audit(docx_path)["fields"]
    return int(fields["main_citation_fields"]), list(fields["record_ids_from_fields"])


def build_docx_from_text(text_file: Path, title: str | None, out_path: Path) -> None:
    raw = text_file.read_text(encoding="utf-8").strip()
    if not raw:
        raise ValueError(f"Text file is empty: {text_file}")
    doc = Document()
    if title:
        doc.add_heading(title.strip(), level=1)
    blocks = [b.strip() for b in re.split(r"\r?\n\s*\r?\n", raw) if b.strip()]
    for block in blocks:
        # Preserve single-paragraph continuity while trimming line wraps
        para_text = " ".join(line.strip() for line in block.splitlines() if line.strip())
        doc.add_paragraph(para_text)
    doc.save(str(out_path))


def iter_eligible_t_nodes(
    node: etree._Element, field_state: dict[str, int]
) -> list[etree._Element]:
    """Collect <w:t> elements outside <w:del> subtrees and outside existing Word fields."""
    if node.tag == W + "del":
        return []
    if node.tag == W + "fldChar":
        kind = node.get(W + "fldCharType")
        if kind == "begin":
            field_state["depth"] += 1
        elif kind == "end" and field_state["depth"] > 0:
            field_state["depth"] -= 1
        return []
    collected: list[etree._Element] = []
    if node.tag == W + "t" and field_state["depth"] == 0:
        collected.append(node)
    for child in node:
        collected.extend(iter_eligible_t_nodes(child, field_state))
    return collected


def paragraph_text_map(
    p_elem: etree._Element,
) -> tuple[str, list[tuple[etree._Element, int, int]]]:
    t_nodes = iter_eligible_t_nodes(p_elem, {"depth": 0})
    parts: list[str] = []
    spans: list[tuple[etree._Element, int, int]] = []
    cursor = 0
    for t_node in t_nodes:
        txt = t_node.text or ""
        parts.append(txt)
        start = cursor
        cursor += len(txt)
        spans.append((t_node, start, cursor))
    return "".join(parts), spans


def locate_span(
    spans: list[tuple[etree._Element, int, int]], char_offset: int
) -> tuple[etree._Element, int]:
    for idx, (t_node, start, end) in enumerate(spans):
        if start < char_offset <= end:
            return t_node, char_offset - start
        if char_offset == 0 and idx == 0 and start == 0:
            return t_node, 0
    # Fallback when char_offset == 0 after an empty leading <w:t>
    for t_node, start, end in spans:
        if start <= char_offset <= end:
            return t_node, char_offset - start
    raise RuntimeError(f"Could not map character offset {char_offset} to any <w:t> run")


def inject_into_docx(
    source_docx: Path,
    out_docx: Path,
    plan: dict[str, object],
    resolved: dict[str, object],
) -> dict[str, object]:
    by_pdf_id: dict[str, dict[str, object]] = dict(resolved.get("by_pdf_id") or {})
    insertions = list(plan.get("insertions") or [])
    if not insertions:
        raise ValueError("citation_plan.json contains zero insertions")

    with zipfile.ZipFile(source_docx, "r") as src_zip:
        doc_bytes = src_zip.read("word/document.xml")

    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(doc_bytes, parser=parser)
    paragraphs = root.findall(".//" + W + "p")
    para_maps = [paragraph_text_map(p) for p in paragraphs]

    resolved_ops: list[dict[str, object]] = []
    used_positions: set[tuple[int, int]] = set()
    used_pdf_ids: set[str] = set()

    for idx, item in enumerate(insertions):
        ins_id = str(item.get("insertion_id") or f"C{idx + 1:02d}")
        anchor_text = str(item.get("anchor_text") or "")
        if not anchor_text:
            raise ValueError(f"Insertion {ins_id} has empty anchor_text")
        pdf_ids = [str(pid) for pid in (item.get("pdf_ids") or [])]
        if not pdf_ids:
            raise ValueError(f"Insertion {ins_id} has empty pdf_ids")
        missing = [pid for pid in pdf_ids if pid not in by_pdf_id]
        if missing:
            raise KeyError(f"Insertion {ins_id} references unresolved pdf_ids: {missing}")

        matches: list[tuple[int, int, int]] = []
        for p_idx, (p_text, _) in enumerate(para_maps):
            search_from = 0
            while True:
                pos = p_text.find(anchor_text, search_from)
                if pos == -1:
                    break
                matches.append((p_idx, pos, pos + len(anchor_text)))
                search_from = pos + 1

        if not matches:
            raise ValueError(f"Anchor text not found for {ins_id}: {anchor_text!r}")
        if "occurrence" not in item and len(matches) > 1:
            raise ValueError(
                f"Ambiguous anchor text ({len(matches)} matches) for {ins_id}; "
                f"specify 'occurrence': {anchor_text!r}"
            )
        occ = int(item.get("occurrence", 1))
        if occ < 1 or occ > len(matches):
            raise IndexError(
                f"Occurrence {occ} out of range (1..{len(matches)}) for {ins_id}"
            )

        p_idx, _, end_off = matches[occ - 1]
        pos_mode = str(item.get("position", "before_terminal_punct"))
        if pos_mode == "before_terminal_punct" and anchor_text[-1] in TERMINAL_PUNCT:
            insert_off = end_off - 1
        elif pos_mode in {"before_terminal_punct", "after_anchor"}:
            insert_off = end_off
        else:
            raise ValueError(f"Unsupported position mode {pos_mode!r} in {ins_id}")

        if (p_idx, insert_off) in used_positions:
            raise ValueError(
                f"Multiple insertions target the exact same offset ({p_idx}, {insert_off}) at {ins_id}; "
                "combine their pdf_ids into one insertion."
            )
        used_positions.add((p_idx, insert_off))
        used_pdf_ids.update(pdf_ids)

        tokens = list(dict.fromkeys(str(by_pdf_id[pid]["citation_token"]) for pid in pdf_ids))
        core = "{" + "; ".join(tokens) + "}"
        p_text = para_maps[p_idx][0]
        use_space = plan.get("space_before_citation", str(plan.get("style")).casefold() != "nature")
        need_space = use_space and insert_off > 0 and not p_text[insert_off - 1].isspace()
        placeholder = (" " + core) if need_space else core

        resolved_ops.append(
            {
                "insertion_id": ins_id,
                "paragraph_index": p_idx,
                "char_offset": insert_off,
                "anchor_text": anchor_text,
                "occurrence": occ,
                "position": pos_mode,
                "pdf_ids": pdf_ids,
                "placeholder": placeholder,
                "record_ids": [int(by_pdf_id[pid]["record_id"]) for pid in pdf_ids],
            }
        )

    # Apply in reverse document order so earlier offsets stay invariant
    for op in sorted(
        resolved_ops,
        key=lambda o: (int(o["paragraph_index"]), int(o["char_offset"])),
        reverse=True,
    ):
        p_idx = int(op["paragraph_index"])
        char_off = int(op["char_offset"])
        _, spans = para_maps[p_idx]
        t_node, local_off = locate_span(spans, char_off)
        old_text = t_node.text or ""
        t_node.text = old_text[:local_off] + str(op["placeholder"]) + old_text[local_off:]
        t_node.set(f"{{{XML_NS}}}space", "preserve")

    new_doc_bytes = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone="yes"
    )
    out_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source_docx, "r") as src_zip, zipfile.ZipFile(
        out_docx, "x", compression=zipfile.ZIP_DEFLATED
    ) as dst_zip:
        for info in src_zip.infolist():
            data = new_doc_bytes if info.filename == "word/document.xml" else src_zip.read(info.filename)
            new_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            new_info.compress_type = zipfile.ZIP_DEFLATED
            dst_zip.writestr(new_info, data)

    existing_main, existing_ids = _audit_existing_cwyw(source_docx)
    new_ids = {rid for op in resolved_ops for rid in op["record_ids"]}
    combined_ids = sorted(set(existing_ids) | new_ids)
    return {
        "existing_main_citations": existing_main,
        "injected_citation_groups": len(resolved_ops),
        "expected_main_citations": existing_main + len(resolved_ops),
        "expected_unique_references": len(combined_ids),
        "expected_record_ids": combined_ids,
        "expected_record_ids_csv": ",".join(str(i) for i in combined_ids),
        "style": plan.get("style"),
        "insertions": resolved_ops,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Losslessly inject EndNote temporary citation placeholders into a DOCX or text manuscript."
    )
    parser.add_argument("--resolved", required=True, type=Path)
    parser.add_argument("--plan", required=True, type=Path)
    src_group = parser.add_mutually_exclusive_group(required=True)
    src_group.add_argument("--text-file", type=Path)
    src_group.add_argument("--docx-in", type=Path)
    parser.add_argument("--docx-out", required=True, type=Path)
    parser.add_argument("--report", type=Path)
    return parser


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    args = build_parser().parse_args()
    resolved_path = args.resolved.resolve()
    plan_path = args.plan.resolve()
    out_docx = check_work_path(args.docx_out, ".docx")
    report_path = check_work_path(args.report, ".json") if args.report else None

    if not resolved_path.is_file():
        raise FileNotFoundError(resolved_path)
    if not plan_path.is_file():
        raise FileNotFoundError(plan_path)
    if out_docx.exists():
        raise FileExistsError(f"Refusing to overwrite output DOCX: {out_docx}")
    if report_path and report_path.exists():
        raise FileExistsError(f"Refusing to overwrite report: {report_path}")

    resolved = json.loads(resolved_path.read_text(encoding="utf-8"))
    plan = json.loads(plan_path.read_text(encoding="utf-8"))

    if args.text_file is not None:
        text_path = args.text_file.resolve()
        if not text_path.is_file():
            raise FileNotFoundError(text_path)
        with tempfile.TemporaryDirectory() as tmpdir:
            base_docx = Path(tmpdir) / "generated_base.docx"
            build_docx_from_text(text_path, plan.get("title"), base_docx)
            summary = inject_into_docx(base_docx, out_docx, plan, resolved)
            input_info = {"mode": "text_file", "text_file": str(text_path)}
    else:
        docx_in = args.docx_in.resolve()
        if not docx_in.is_file():
            raise FileNotFoundError(docx_in)
        if docx_in == out_docx:
            raise RuntimeError("--docx-in and --docx-out must differ")
        summary = inject_into_docx(docx_in, out_docx, plan, resolved)
        input_info = {
            "mode": "docx_in",
            "docx_in": str(docx_in),
            "docx_in_sha256": sha256_file(docx_in),
        }

    result = {
        "status": "PASS",
        "input": input_info,
        "docx_out": str(out_docx),
        "output_sha256": sha256_file(out_docx),
        **summary,
    }
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if report_path:
        write_json_atomic(report_path, result)
    print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
