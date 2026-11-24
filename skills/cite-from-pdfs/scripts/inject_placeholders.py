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
import os
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
ASCII_TERMINAL_PUNCT = set(".,;:!?")
CJK_TERMINAL_PUNCT = set("。，；：！？")
TERMINAL_PUNCT = ASCII_TERMINAL_PUNCT | CJK_TERMINAL_PUNCT
# NUL cannot occur in XML text. It separates spans that cannot share an anchor.
ANCHOR_BOUNDARY = "\x00"
_NORM_1TO1_TRANS = str.maketrans({
    "‘": "'", "’": "'", "‚": "'",
    "“": '"', "”": '"', "„": '"',
    "–": "-", "—": "-", "−": "-",
    "\u00a0": " ",
})
_STATIC_BIB_RE = re.compile(
    r"^(?:(?:\d+(?:\.\d+)*|[一二三四五六七八九十]+)\s*[.、]?\s*)?(?:references|bibliography|literature cited|works cited|参考文献|文献)\s*:?$",
    re.IGNORECASE,
)
_POST_BIB_HEADING_RE = re.compile(
    r"^(?:(?:\d+(?:\.\d+)*|[A-Z]|[一二三四五六七八九十]+)\s*[.、]?\s*)?(?:appendix|supplementary(?:\s+(?:materials|methods|information|notes|data))?|supporting information|acknowledgements?|个人简历|附\s*录|补充材料|致\s*谢)\b",
    re.IGNORECASE,
)
_SKIP_PARA_CHILDREN = {
    W + "pPr", W + "bookmarkStart", W + "bookmarkEnd", W + "proofErr",
    W + "commentRangeStart", W + "commentRangeEnd", W + "permStart", W + "permEnd",
}


def _norm_1to1(value: str) -> str:
    """Length-preserving typography normalization for curly quotes, dashes, and NBSP."""
    out = value.translate(_NORM_1TO1_TRANS)
    if len(out) != len(value):
        raise AssertionError("_norm_1to1 changed string length")
    return out


def _is_cjk_char(ch: str) -> bool:
    return (
        "\u4e00" <= ch <= "\u9fff"
        or "\u3400" <= ch <= "\u4dbf"
        or "\u3000" <= ch <= "\u303f"
        or "\uff00" <= ch <= "\uffef"
    )


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
    blocks = [b.strip() for b in re.split(r"\r?\n(?:[ \t]*\r?\n)+", raw) if b.strip()]
    for block in blocks:
        # Preserve single-paragraph continuity while trimming line wraps
        para_text = " ".join(line.strip() for line in block.splitlines() if line.strip())
        doc.add_paragraph(para_text)
    doc.save(str(out_path))


def paragraph_text_map(
    p_elem: etree._Element, field_state: dict[str, int],
) -> tuple[str, list[tuple[etree._Element, int, int]]]:
    parts: list[str] = []
    spans: list[tuple[etree._Element, int, int]] = []
    cursor = 0
    # The native preparation audit also reads only direct runs in each paragraph.
    for run in p_elem:
        if run.tag in _SKIP_PARA_CHILDREN:
            continue
        if run.tag != W + "r":
            parts.append(ANCHOR_BOUNDARY)
            cursor += 1
            continue
        for child in run:
            if child.tag in {W + "rPr", W + "commentReference"}:
                continue
            if child.tag == W + "t" and not field_state["depth"]:
                txt = child.text or ""
                parts.append(txt)
                spans.append((child, cursor, cursor + len(txt)))
                cursor += len(txt)
                continue
            if child.tag == W + "noBreakHyphen" and not field_state["depth"]:
                parts.append("-")
                cursor += 1
                continue
            if child.tag == W + "fldChar":
                kind = child.get(W + "fldCharType")
                if kind == "begin":
                    field_state["depth"] += 1
                elif kind == "end" and field_state["depth"]:
                    field_state["depth"] -= 1
            parts.append(ANCHOR_BOUNDARY)
            cursor += 1
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


def read_manuscript(source_docx):
    with zipfile.ZipFile(source_docx, "r") as src_zip:
        doc_bytes = src_zip.read("word/document.xml")

    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False, no_network=True)
    root = etree.fromstring(doc_bytes, parser=parser)
    body = root.find(W + "body")
    paragraphs = root.xpath(".//w:body//w:p[not(ancestor::w:txbxContent or ancestor::w:del or ancestor::w:ins or ancestor::w:moveFrom or ancestor::w:moveTo or ancestor::w:sdt or ancestor::w:fldSimple)]",
                            namespaces={"w": W_NS})
    field_state = {"depth": 0}
    para_maps: list[tuple[str, list[tuple[etree._Element, int, int]]]] = []
    in_static_bib = False
    for p in paragraphs:
        mapped = paragraph_text_map(p, field_state)
        visible = mapped[0].replace(ANCHOR_BOUNDARY, "").strip()
        is_top_level = p.getparent() is body
        if not in_static_bib:
            if is_top_level and len(visible) < 40 and _STATIC_BIB_RE.match(visible):
                in_static_bib = True
        elif is_top_level and len(visible) < 80 and _POST_BIB_HEADING_RE.match(visible):
            in_static_bib = False
        para_maps.append(("", []) if in_static_bib else mapped)

    return root, para_maps


def _resolve_position_offset(
    anchor_text: str,
    end_off: int,
    pos_mode: str,
    style: str | None,
) -> tuple[int, str]:
    mode = pos_mode.strip()
    if mode not in {"auto", "before_terminal_punct", "after_anchor"}:
        raise ValueError(f"Unsupported position mode {pos_mode!r}")
    effective_mode = mode
    is_nature = str(style or "").strip().casefold().startswith("nature")
    if mode == "auto":
        if is_nature:
            if anchor_text[-1] in CJK_TERMINAL_PUNCT:
                effective_mode = "before_terminal_punct"
            elif anchor_text[-1] in ASCII_TERMINAL_PUNCT:
                effective_mode = "after_anchor"
            else:
                effective_mode = "after_anchor"
        else:
            effective_mode = "before_terminal_punct"
    if effective_mode == "before_terminal_punct" and anchor_text[-1] in TERMINAL_PUNCT:
        return end_off - 1, effective_mode
    return end_off, effective_mode


def locate_placements(para_maps, insertions, style: str | None = None):
    placements = []
    used_positions: set[tuple[int, int]] = set()
    labels = set()

    norm_paras = [_norm_1to1(p_text) for p_text, _ in para_maps]

    for idx, item in enumerate(insertions):
        ins_id = str(item.get("insertion_id") or f"C{idx + 1:02d}")
        if ins_id in labels:
            raise ValueError(f"Use a distinct insertion_id for each placement: {ins_id}")
        labels.add(ins_id)
        anchor_text = str(item.get("anchor_text") or "")
        if not anchor_text:
            raise ValueError(f"Insertion {ins_id} has empty anchor_text")
        if ANCHOR_BOUNDARY in anchor_text:
            raise ValueError(f"Insertion {ins_id} contains an invalid character")
        pdf_ids = [str(pid) for pid in (item.get("pdf_ids") or [])]
        if not pdf_ids:
            raise ValueError(f"Insertion {ins_id} has empty pdf_ids")

        norm_anchor = _norm_1to1(anchor_text)
        matches: list[tuple[int, int, int]] = []
        for p_idx, p_norm in enumerate(norm_paras):
            search_from = 0
            while True:
                pos = p_norm.find(norm_anchor, search_from)
                if pos == -1:
                    break
                matches.append((p_idx, pos, pos + len(norm_anchor)))
                search_from = pos + 1

        if not matches:
            bridged = any(norm_anchor in p_norm.replace(ANCHOR_BOUNDARY, "") for p_norm in norm_paras if p_norm)
            hint = (
                " (the phrase crosses a hyperlink, tracked change, or content-control boundary; shorten anchor_text or accept revisions)"
                if bridged
                else ""
            )
            raise ValueError(f"Anchor text not found for {ins_id}: {anchor_text!r}{hint}")
        if "occurrence" not in item and len(matches) > 1:
            raise ValueError(
                f"Ambiguous anchor text ({len(matches)} matches) for {ins_id}; "
                f"specify 'occurrence': {anchor_text!r}"
            )
        occ = item.get("occurrence", 1)
        if isinstance(occ, str) and re.fullmatch(r"[0-9]+", occ):
            occ = int(occ)
        if type(occ) is not int:
            raise ValueError(f"Occurrence must be an integer for {ins_id}")
        if occ < 1 or occ > len(matches):
            raise ValueError(
                f"Occurrence {occ} out of range (1..{len(matches)}) for {ins_id}"
            )

        p_idx, start_off, end_off = matches[occ - 1]
        raw_mode = str(item.get("position") or ("auto" if style is not None else "before_terminal_punct"))
        try:
            insert_off, resolved_mode = _resolve_position_offset(anchor_text, end_off, raw_mode, style)
        except ValueError as exc:
            raise ValueError(f"{exc} in {ins_id}") from exc

        if (p_idx, insert_off) in used_positions:
            raise ValueError(
                f"Multiple insertions target the exact same offset ({p_idx}, {insert_off}) at {ins_id}; "
                "combine their pdf_ids into one insertion."
            )
        used_positions.add((p_idx, insert_off))
        p_raw = para_maps[p_idx][0]
        local_rel = insert_off - start_off
        matched_slice = p_raw[start_off:end_off].replace(ANCHOR_BOUNDARY, "")
        inline_preview = f"{matched_slice[:max(0, local_rel)]}[+Cite]{matched_slice[max(0, local_rel):]}"
        placements.append(
            {
                "insertion_id": ins_id,
                "paragraph_index": p_idx,
                "char_offset": insert_off,
                "anchor_text": anchor_text,
                "occurrence": occ,
                "position": resolved_mode,
                "inline_preview": inline_preview,
                "pdf_ids": pdf_ids,
            }
        )
    return placements


def inject_into_docx(
    source_docx: Path,
    out_docx: Path,
    plan: dict[str, object],
    resolved: dict[str, object],
) -> dict[str, object]:
    insertions = list(plan.get("insertions") or [])
    if not insertions:
        raise ValueError("citation_plan.json contains zero insertions")
    style = plan.get("style")
    root, para_maps = read_manuscript(source_docx)
    resolved_ops = locate_placements(para_maps, insertions, style=style)
    by_pdf_id = resolved.get("by_pdf_id") or {}
    use_space = plan.get("space_before_citation", not str(style or "").strip().casefold().startswith("nature"))
    for op in resolved_ops:
        pdf_ids = op["pdf_ids"]
        missing = [pid for pid in pdf_ids if pid not in by_pdf_id]
        if missing:
            raise KeyError(f"Insertion {op['insertion_id']} references unresolved pdf_ids: {missing}")
        tokens = dict.fromkeys(str(by_pdf_id[pid]["citation_token"]) for pid in pdf_ids)
        offset = op["char_offset"]
        paragraph = para_maps[op["paragraph_index"]][0]
        prev_ch = paragraph[offset - 1] if offset > 0 else ""
        space = " " if use_space and prev_ch and not prev_ch.isspace() and not _is_cjk_char(prev_ch) else ""
        op["placeholder"] = space + "{" + "; ".join(tokens) + "}"
        op["record_ids"] = [int(by_pdf_id[pid]["record_id"]) for pid in pdf_ids]

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
    if out_docx.exists():
        raise FileExistsError(f"Refusing to overwrite output DOCX: {out_docx}")
    tmp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb", dir=out_docx.parent, prefix=f".{out_docx.name}.", suffix=".tmp", delete=False
        ) as handle:
            tmp_path = handle.name
        with zipfile.ZipFile(source_docx, "r") as src_zip, zipfile.ZipFile(
            tmp_path, "w", compression=zipfile.ZIP_DEFLATED
        ) as dst_zip:
            for info in src_zip.infolist():
                data = new_doc_bytes if info.filename == "word/document.xml" else src_zip.read(info.filename)
                new_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                new_info.compress_type = zipfile.ZIP_DEFLATED
                dst_zip.writestr(new_info, data)
        if os.name == "nt":
            os.rename(tmp_path, out_docx)
        else:
            os.link(tmp_path, out_docx)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

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
