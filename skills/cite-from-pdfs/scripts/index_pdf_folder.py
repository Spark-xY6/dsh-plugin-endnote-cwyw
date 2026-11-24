# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "PyMuPDF>=1.24.0",
#     "httpx>=0.27.0",
# ]
# ///

from __future__ import annotations

import argparse
from collections import Counter
import difflib
import hashlib
import html
import json
import math
import re
import sys
import urllib.parse
from pathlib import Path
from uuid import uuid4

import fitz
import httpx

from _common import (
    copy_sdb_snapshot,
    check_work_path,
    first_author_family,
    list_process_pids,
    normalize_author_family,
    normalize_doi,
    normalize_title,
    parse_endnote_authors,
    parse_pdf_meta_authors,
    read_endnote_active_records,
    read_library_id,
    split_name_with_particles,
    tokenize_scientific,
    write_json_atomic,
)
from reference_identity import IdentityReview, RecordIndex


DOI_PREF_RE = re.compile(
    r"(?:https?://(?:dx\.)?doi\.org/|doi:\s*)(10\.\d{4,9}/[^\s<>\"']+)",
    re.IGNORECASE,
)
DOI_BARE_RE = re.compile(r"\b(10\.\d{4,9}/[^\s<>\"']+)")
XMP_DOI_RE = re.compile(
    r"<(?:prism:doi|dc:identifier|pdfx:doi)[^>]*>\s*(?:doi:\s*)?(10\.\d{4,9}/[^<\s]+)\s*</",
    re.IGNORECASE,
)
ARXIV_RE = re.compile(r"\barXiv:\s*(\d{4}\.\d{4,5})(?:v\d+)?\b", re.IGNORECASE)
REF_HEADING_RE = re.compile(
    r"^(?:\d+\.?\s*)?(?:references|bibliography|literature cited|works cited)\s*:?$",
    re.IGNORECASE,
)
GARBAGE_TITLE_RE = re.compile(
    r"^(?:untitled|microsoft word|doi:|https?://|arxiv:|page\s+\d+|full\s+article|research\s+article|article|letter)$",
    re.IGNORECASE,
)


def strip_xml_tags(raw: str) -> str:
    cleaned = re.sub(r"<[^>]+>", " ", raw)
    cleaned = html.unescape(cleaned)
    return " ".join(cleaned.split()).strip()


def extract_doi_from_text(text_chunk: str) -> str:
    for pattern in (DOI_PREF_RE, DOI_BARE_RE):
        match = pattern.search(text_chunk)
        if match:
            candidate = normalize_doi(match.group(1))
            if candidate:
                return candidate
    return ""


def extract_arxiv_from_text(text_chunk: str) -> str:
    match = ARXIV_RE.search(text_chunk)
    return match.group(1) if match else ""


def build_overlapping_chunks(
    raw_blocks: list[tuple[int, str]],
    target_chars: int = 850,
    max_chars: int = 1050,
) -> list[dict[str, object]]:
    """Merge non-reference PDF blocks into contextual ~700-1000 char chunks with 1-block overlap."""
    chunks: list[dict[str, object]] = []
    buf: list[tuple[int, str]] = []
    buf_len = 0
    for page_num, block_text in raw_blocks:
        if buf and (buf_len >= target_chars or buf_len + 1 + len(block_text) > max_chars):
            pages = sorted({p for p, _ in buf})
            page_label = str(pages[0]) if len(pages) == 1 else f"{pages[0]}-{pages[-1]}"
            chunks.append({"page": page_label, "text": " ".join(t for _, t in buf)})
            if len(buf) > 1 and len(buf[-1][1]) < max_chars:
                buf = [buf[-1]]
                buf_len = len(buf[0][1])
            else:
                buf = []
                buf_len = 0
        buf.append((page_num, block_text))
        buf_len += (1 if buf_len else 0) + len(block_text)
    if buf:
        pages = sorted({p for p, _ in buf})
        page_label = str(pages[0]) if len(pages) == 1 else f"{pages[0]}-{pages[-1]}"
        chunks.append({"page": page_label, "text": " ".join(t for _, t in buf)})
    return chunks


def is_usable_title(candidate: str, pdf_stem: str) -> bool:
    cleaned = " ".join(candidate.split()).strip()
    if len(cleaned) < 15:
        return False
    if GARBAGE_TITLE_RE.match(cleaned):
        return False
    if cleaned.casefold().endswith((".pdf", ".doc", ".docx")):
        return False
    if re.search(r"\b[a-z]{2,10}\.\d{4,}|\.\.", cleaned, re.IGNORECASE):
        return False
    cjk_count = len(re.findall(r"[\u4e00-\u9fff]", cleaned))
    alpha_words = [w for w in cleaned.split() if re.fullmatch(r"[A-Za-z][A-Za-z\-]{1,}", w)]
    if cjk_count < 4 and len(alpha_words) < 3:
        return False
    if normalize_title(cleaned) == normalize_title(pdf_stem) and len(cleaned.split()) < 4:
        return False
    return True


def heuristic_title_from_page1(page: fitz.Page) -> str:
    """Extract likely article title using consecutive largest-font lines on page 1."""
    data = page.get_text("dict")
    candidates: list[tuple[float, float, str]] = []
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        sized_lines: list[tuple[float, float, str]] = []
        for line in block.get("lines", []):
            line_chars: list[str] = []
            line_sizes: list[float] = []
            for span in line.get("spans", []):
                txt = str(span.get("text", "")).strip()
                if txt:
                    line_chars.append(txt)
                    line_sizes.append(float(span.get("size", 0.0)))
            joined = " ".join(line_chars).strip()
            if joined and line_sizes:
                sized_lines.append((sum(line_sizes) / len(line_sizes), float(line.get("bbox", [0, 0, 0, 0])[1]), joined))
        if not sized_lines:
            continue
        top_size = max(sz for sz, _, _ in sized_lines)
        title_lines = [txt for sz, _, txt in sized_lines if sz >= top_size * 0.88]
        full_text = " ".join(title_lines).strip()
        if top_size >= 11.5 and len(full_text) >= 15 and len(full_text.split()) >= 4:
            candidates.append((top_size, -sized_lines[0][1], full_text))
    if not candidates:
        return ""
    candidates.sort(reverse=True)
    best = " ".join(candidates[0][2].split())
    best = re.sub(r"^(?:Article|Letter|Review|Perspective)\s+https?://\S+\s+", "", best, flags=re.IGNORECASE)
    return best.strip()


def heuristic_year_from_pages(pages_text: list[str]) -> str:
    head = "\n".join(pages_text[:2])
    # Prefer copyright / published / received year patterns
    for pattern in (
        r"(?:published|accepted|©|copyright|\bvol\.|\bnature\b|\bcell\b|\bscience\b)[^\n]{0,60}\b((?:19|20)\d{2})\b",
        r"\b(20\d{2}|19[89]\d)\b",
    ):
        match = re.search(pattern, head, flags=re.IGNORECASE)
        if match:
            return match.group(1)
    return ""


def order_page_blocks(blocks: list[tuple], page_width: float) -> list[tuple]:
    """Order multi-column text blocks in human reading order (top span -> left column -> right column -> bottom)."""
    text_blocks = [b for b in blocks if len(b) >= 7 and b[6] == 0]
    if len(text_blocks) < 4 or page_width <= 0:
        return sorted(text_blocks, key=lambda b: (round(float(b[1]), 1), round(float(b[0]), 1)))
    mid_x = 0.5 * page_width
    spanning = [
        b for b in text_blocks
        if (float(b[2]) - float(b[0])) > 0.58 * page_width
        or (float(b[0]) < 0.35 * page_width and float(b[2]) > 0.65 * page_width)
    ]
    cols = [b for b in text_blocks if b not in spanning]
    left = [b for b in cols if 0.5 * (float(b[0]) + float(b[2])) < mid_x]
    right = [b for b in cols if 0.5 * (float(b[0]) + float(b[2])) >= mid_x]
    if len(left) >= 2 and len(right) >= 2:
        col_top = min(float(b[1]) for b in left + right)
        top_span = sorted([b for b in spanning if float(b[1]) <= col_top], key=lambda b: (round(float(b[1]), 1), round(float(b[0]), 1)))
        bottom_span = sorted([b for b in spanning if float(b[1]) > col_top], key=lambda b: (round(float(b[1]), 1), round(float(b[0]), 1)))
        left_sorted = sorted(left, key=lambda b: (round(float(b[1]), 1), round(float(b[0]), 1)))
        right_sorted = sorted(right, key=lambda b: (round(float(b[1]), 1), round(float(b[0]), 1)))
        return top_span + left_sorted + right_sorted + bottom_span
    return sorted(text_blocks, key=lambda b: (round(float(b[1]), 1), round(float(b[0]), 1)))


def _extract_offline_authors(
    meta: dict[str, object],
    xmp_xml: str,
    page1_text: str,
    title: str,
) -> list[dict[str, str]]:
    """Extract authors offline via PDF Info, XMP dc:creator, or page-1 header byline."""
    from_meta = parse_pdf_meta_authors(meta.get("author") or "")
    if from_meta:
        return from_meta

    if xmp_xml:
        creator_blocks = re.findall(r"<dc:creator[^>]*>(.*?)</dc:creator>", xmp_xml, re.IGNORECASE | re.DOTALL)
        raw_xmp: list[str] = []
        for blk in creator_blocks:
            items = re.findall(r"<rdf:li[^>]*>(.*?)</rdf:li>", blk, re.IGNORECASE | re.DOTALL)
            raw_xmp.extend(strip_xml_tags(i) for i in (items or [blk]) if strip_xml_tags(i))
        if not raw_xmp:
            pdf_auth = re.findall(r"<pdf:Author[^>]*>(.*?)</pdf:Author>", xmp_xml, re.IGNORECASE | re.DOTALL)
            raw_xmp.extend(strip_xml_tags(a) for a in pdf_auth if strip_xml_tags(a))
        if raw_xmp:
            parsed_xmp = (
                parse_endnote_authors("\n".join(raw_xmp))
                if len(raw_xmp) > 1
                else parse_pdf_meta_authors(raw_xmp[0])
            )
            if parsed_xmp:
                return parsed_xmp

    if page1_text:
        norm_title = normalize_title(title)
        affil_re = re.compile(
            r"@|https?://|\b(?:doi|university|institute|department|school|laboratory|college|center|centre|hospital|academy|accepted|received|published|abstract|summary|keywords)\b|\d{4,}",
            re.IGNORECASE,
        )
        past_title = not bool(norm_title)
        for line in [ln.strip() for ln in page1_text.splitlines()[:25] if ln.strip()]:
            if re.fullmatch(r"(?:abstract|summary|introduction)", line, re.IGNORECASE):
                break
            if not past_title and norm_title and normalize_title(line) and (
                normalize_title(line) in norm_title or norm_title in normalize_title(line)
            ):
                past_title = True
                continue
            if affil_re.search(line) or len(line) > 140 or ("." in line and not re.search(r"\b[A-Z]\.", line)):
                continue
            cleaned = re.sub(r"(?<=[A-Za-z])[0-9*†‡§]+(?:,[0-9*†‡§]+)*", "", line)
            chunks = [c.strip() for c in re.split(r"\s*,\s*|\s*;\s*|\s+and\s+", cleaned) if c.strip()]
            particles = {"van", "von", "de", "del", "della", "der", "den", "la", "le", "du", "di", "da", "dos"}
            name_chunks = [
                c for c in chunks
                if 2 <= len(c.split()) <= 5
                and re.match(r"^[A-Z][\w'.-]*$", c.split()[0])
                and re.match(r"^[A-Z][\w'.-]*$", c.split()[-1])
                and all(re.match(r"^[A-Z][\w'.-]*$", w) or w.casefold() in particles for w in c.split())
            ]
            if name_chunks and (past_title or len(name_chunks) >= 2):
                byline_authors = parse_endnote_authors("\n".join(name_chunks))
                if byline_authors:
                    return byline_authors
    return []


def extract_pdf_record(
    pdf_path: Path, pdf_root: Path, pdf_id: str, max_pages_limit: int = 30
) -> dict[str, object]:
    """Parse multi-layer PDF identifiers, heuristic title/year, and overlapping non-reference chunks."""
    with fitz.open(pdf_path) as doc:
        meta = dict(doc.metadata or {})
        xmp_xml = str(doc.get_xml_metadata() or "")
        total_pages = len(doc)
        max_pages = min(total_pages, max(1, int(max_pages_limit)))
        pages_text: list[str] = []
        raw_blocks: list[tuple[int, str]] = []
        stop_body = False
        for page_idx in range(max_pages):
            page = doc[page_idx]
            pages_text.append(page.get_text("text"))
            if stop_body:
                continue
            blocks = order_page_blocks(page.get_text("blocks"), float(page.rect.width))
            for block in blocks:
                raw_block = str(block[4]).strip()
                if not raw_block:
                    continue
                first_line = raw_block.splitlines()[0].strip()
                if page_idx >= 2 and REF_HEADING_RE.match(first_line):
                    stop_body = True
                    break
                cleaned = re.sub(r"-\s*\n\s*", "", raw_block)
                cleaned = " ".join(cleaned.split())
                if len(cleaned) >= 40:
                    raw_blocks.append((page_idx + 1, cleaned))

        passages = build_overlapping_chunks(raw_blocks)

        # Multi-layer identifier extraction (pdf2doi order):
        # Prefer embedded identification and page 1/2 imprints.
        meta_blob = " ".join(
            str(meta.get(k, "") or "") for k in ("subject", "keywords", "title")
        )
        doi = extract_doi_from_text(meta_blob)
        identifier_layer = "pdf_info" if doi else "none"

        if not doi and xmp_xml:
            xmp_match = XMP_DOI_RE.search(xmp_xml)
            doi = normalize_doi(xmp_match.group(1)) if xmp_match else extract_doi_from_text(xmp_xml)
            if doi:
                identifier_layer = "xmp_xml"

        if not doi and pages_text:
            for p_idx in range(min(2, len(pages_text))):
                imprint = re.split(r"(?im)^\s*(?:references|bibliography)\s*$", pages_text[p_idx], maxsplit=1)[0]
                doi = extract_doi_from_text(imprint)
                if doi:
                    identifier_layer = "page_text"
                    break

        arxiv_blob = "\n".join(
            [meta_blob, xmp_xml, pages_text[0] if pages_text else ""]
        )
        arxiv_id = extract_arxiv_from_text(arxiv_blob)
        if not doi and arxiv_id:
            identifier_layer = "arxiv_id"

        meta_title = str(meta.get("title", "") or "").strip()
        if is_usable_title(meta_title, pdf_path.stem):
            title = " ".join(meta_title.split())
        elif total_pages > 0:
            title = heuristic_title_from_page1(doc[0])
        else:
            title = ""

        year = heuristic_year_from_pages(pages_text)
        offline_authors = _extract_offline_authors(
            meta, xmp_xml, pages_text[0] if pages_text else "", title
        )
        total_text_chars = sum(len(" ".join(t.split())) for t in pages_text)
        needs_ocr = bool(total_pages > 0 and total_text_chars < 80)
        abstract_guess = ""
        for p in passages[:6]:
            low = str(p["text"]).casefold()
            if "summary" in low or "abstract" in low or len(str(p["text"])) >= 250:
                abstract_guess = str(p["text"])[:1500]
                break

    stat = pdf_path.stat()
    return {
        "pdf_id": pdf_id,
        "pdf_path": str(pdf_path.resolve()),
        "relative_path": pdf_path.relative_to(pdf_root).as_posix(),
        "file_size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
        "total_pages": total_pages,
        "doi": doi,
        "arxiv_id": arxiv_id,
        "identifier_layer": identifier_layer,
        "heuristic_title": title,
        "heuristic_year": year,
        "first_page_imprint": " ".join(pages_text[0].split())[:600] if pages_text else "",
        "authors": offline_authors,
        "first_author_family": first_author_family(offline_authors),
        "needs_ocr": needs_ocr,
        "heuristic_abstract": abstract_guess,
        "passages": passages[:150],
    }


def fetch_json_with_proxy_policy(
    url: str, params: dict[str, str] | None = None
) -> tuple[dict[str, object] | None, str]:
    """Fetch JSON ignoring proxy env first; fall back to proxy-preserving client only on transport error."""
    headers = {"User-Agent": "EndNoteAutoCite/1.0 (mailto:endnote-autocite@example.org)"}
    try:
        with httpx.Client(trust_env=False, timeout=8.0, follow_redirects=True, headers=headers) as client:
            resp = client.get(url, params=params)
            if resp.status_code == 200:
                return resp.json(), "direct_no_proxy"
            return None, "direct_no_proxy"
    except (httpx.TransportError, httpx.TimeoutException):
        try:
            with httpx.Client(trust_env=True, timeout=15.0, follow_redirects=True, headers=headers) as client:
                resp = client.get(url, params=params)
                if resp.status_code == 200:
                    return resp.json(), "proxy_preserving_fallback"
                return None, "proxy_preserving_fallback"
        except Exception:
            return None, "proxy_preserving_fallback_failed"


def parse_crossref_work(message: dict[str, object]) -> dict[str, object]:
    titles = message.get("title") or []
    title = strip_xml_tags(str(titles[0])) if titles else ""
    authors: list[dict[str, str]] = []
    for raw_a in message.get("author") or []:
        if not isinstance(raw_a, dict):
            continue
        family = str(raw_a.get("family", "") or raw_a.get("name", "")).strip()
        given = str(raw_a.get("given", "")).strip()
        if family:
            authors.append({"family": family, "given": given})
    year = ""
    for key in ("published-print", "published-online", "published", "issued", "created"):
        part = message.get(key)
        if isinstance(part, dict):
            dp = part.get("date-parts") or []
            if dp and isinstance(dp[0], list) and dp[0]:
                year = str(dp[0][0])
                break
    containers = message.get("container-title") or message.get("short-container-title") or []
    journal = strip_xml_tags(str(containers[0])) if containers else ""
    abstract = strip_xml_tags(str(message.get("abstract", "") or ""))
    update_to = message.get("update-to") or []
    crossref_retracted = any(
        isinstance(u, dict) and str(u.get("type", "")).casefold() in {"retraction", "withdrawal"}
        for u in update_to
    ) or title.strip().upper().startswith("RETRACTED:")
    return {
        "title": title,
        "authors": authors,
        "first_author_family": first_author_family(authors),
        "year": year,
        "journal": journal,
        "volume": str(message.get("volume", "") or "").strip(),
        "issue": str(message.get("issue", "") or "").strip(),
        "pages": str(message.get("page") or message.get("article-number") or "").strip(),
        "doi": normalize_doi(message.get("DOI", "")),
        "abstract": abstract,
        "is_retracted": crossref_retracted,
    }


def reconstruct_openalex_abstract(inv_idx: object) -> str:
    if not isinstance(inv_idx, dict) or not inv_idx:
        return ""
    positions: list[tuple[int, str]] = []
    for word, idxs in inv_idx.items():
        if isinstance(idxs, list):
            for pos in idxs:
                positions.append((int(pos), str(word)))
    positions.sort(key=lambda pair: pair[0])
    return " ".join(w for _, w in positions).strip()


def fetch_openalex_work(
    doi: str = "", arxiv_id: str = ""
) -> tuple[dict[str, object] | None, str]:
    if doi:
        url = f"https://api.openalex.org/works/https://doi.org/{urllib.parse.quote(doi, safe='')}"
    elif arxiv_id:
        url = f"https://api.openalex.org/works/https://arxiv.org/abs/{urllib.parse.quote(arxiv_id, safe='')}"
    else:
        return None, "not_needed"
    payload, mode = fetch_json_with_proxy_policy(url)
    if not payload or not isinstance(payload, dict) or not payload.get("id"):
        return None, mode

    title = strip_xml_tags(str(payload.get("title") or payload.get("display_name") or ""))
    authors: list[dict[str, str]] = []
    for authorship in payload.get("authorships") or []:
        if not isinstance(authorship, dict):
            continue
        display_name = str(((authorship.get("author") or {}).get("display_name") or "")).strip()
        if not display_name:
            continue
        if "," in display_name:
            fam, giv = display_name.split(",", 1)
            authors.append({"family": fam.strip(), "given": giv.strip()})
        else:
            authors.append(split_name_with_particles(display_name))

    biblio = payload.get("biblio") or {}
    first_p = str(biblio.get("first_page") or "").strip()
    last_p = str(biblio.get("last_page") or "").strip()
    pages = f"{first_p}-{last_p}" if first_p and last_p and first_p != last_p else first_p
    source_obj = (payload.get("primary_location") or {}).get("source") or {}
    journal = str(source_obj.get("display_name") or "").strip()
    oa_doi = normalize_doi(payload.get("doi") or doi)
    abstract = reconstruct_openalex_abstract(payload.get("abstract_inverted_index"))
    is_retracted = bool(payload.get("is_retracted") is True) or title.strip().upper().startswith("RETRACTED:")
    return (
        {
            "title": title,
            "authors": authors,
            "first_author_family": first_author_family(authors),
            "year": str(payload.get("publication_year") or "").strip(),
            "journal": journal,
            "volume": str(biblio.get("volume") or "").strip(),
            "issue": str(biblio.get("issue") or "").strip(),
            "pages": pages,
            "doi": oa_doi,
            "abstract": abstract,
            "is_retracted": is_retracted,
        },
        mode,
    )


def fetch_pubmed_by_doi(doi: str) -> tuple[dict[str, object] | None, str]:
    esearch, mode = fetch_json_with_proxy_policy(
        "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi",
        params={"db": "pubmed", "term": f"{doi}[LID]", "retmode": "json"},
    )
    if not esearch:
        return None, mode
    ids = ((esearch.get("esearchresult") or {}).get("idlist") or [])
    if len(ids) != 1 or str((esearch.get("esearchresult") or {}).get("count")) != "1":
        return None, mode
    pmid = str(ids[0])
    esummary, mode2 = fetch_json_with_proxy_policy(
        "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi",
        params={"db": "pubmed", "id": pmid, "retmode": "json"},
    )
    if not esummary:
        return None, mode2
    doc = (esummary.get("result") or {}).get(pmid)
    if not isinstance(doc, dict):
        return None, mode2
    returned = {normalize_doi(item.get("value")) for item in doc.get("articleids", []) if item.get("idtype") == "doi"}
    if returned != {doi}:
        return None, mode2
    authors: list[dict[str, str]] = []
    for a in doc.get("authors") or []:
        name = str(a.get("name", "")).strip()
        if not name:
            continue
        parts = name.split()
        if len(parts) >= 2:
            authors.append({"family": " ".join(parts[:-1]), "given": parts[-1]})
        else:
            authors.append({"family": name, "given": ""})
    pubdate = str(doc.get("pubdate", ""))
    year_m = re.search(r"\b((?:19|20)\d{2})\b", pubdate)
    return (
        {
            "title": strip_xml_tags(str(doc.get("title", "")).rstrip(".")),
            "authors": authors,
            "first_author_family": first_author_family(authors),
            "year": year_m.group(1) if year_m else "",
            "journal": str(doc.get("fulljournalname") or doc.get("source") or "").strip(),
            "volume": str(doc.get("volume", "")).strip(),
            "issue": str(doc.get("issue", "")).strip(),
            "pages": str(doc.get("pages", "")).strip(),
            "doi": doi,
            "pmid": pmid,
            "abstract": "",
        },
        mode2,
    )


def enrich_metadata(parsed: dict[str, object]) -> tuple[dict[str, object], str, str]:
    """Query Crossref, OpenAlex (abstract + retraction check), and PubMed fallback."""
    doi = str(parsed.get("doi", ""))
    arxiv_id = str(parsed.get("arxiv_id", ""))
    heuristic_title = str(parsed.get("heuristic_title", ""))
    network_mode = "not_needed"

    cr_meta: dict[str, object] | None = None
    oa_meta: dict[str, object] | None = None

    if doi:
        url = f"https://api.crossref.org/works/{urllib.parse.quote(doi, safe='')}"
        payload, mode = fetch_json_with_proxy_policy(url)
        network_mode = mode
        if payload and isinstance(payload.get("message"), dict):
            cr_meta = parse_crossref_work(payload["message"])

    if doi or arxiv_id:
        oa_meta, oa_mode = fetch_openalex_work(doi=doi, arxiv_id=arxiv_id)
        if oa_mode not in {"not_needed", "direct_no_proxy"}:
            network_mode = oa_mode
        if not doi and oa_meta and oa_meta.get("doi"):
            doi = str(oa_meta["doi"])
            url = f"https://api.crossref.org/works/{urllib.parse.quote(doi, safe='')}"
            payload, mode = fetch_json_with_proxy_policy(url)
            if payload and isinstance(payload.get("message"), dict):
                cr_meta = parse_crossref_work(payload["message"])

    if cr_meta and cr_meta.get("authors") and cr_meta.get("year") and cr_meta.get("title"):
        merged = dict(cr_meta)
        if not merged.get("abstract") and oa_meta and oa_meta.get("abstract"):
            merged["abstract"] = oa_meta["abstract"]
        merged["is_retracted"] = bool(
            cr_meta.get("is_retracted") or (oa_meta and oa_meta.get("is_retracted"))
        )
        return merged, "crossref_doi", network_mode

    if oa_meta and oa_meta.get("authors") and oa_meta.get("year") and oa_meta.get("title"):
        return oa_meta, ("openalex_arxiv" if arxiv_id and not parsed.get("doi") else "openalex_doi"), network_mode

    if doi:
        pm_meta, pm_mode = fetch_pubmed_by_doi(doi)
        if pm_mode != "direct_no_proxy":
            network_mode = pm_mode
        if pm_meta and pm_meta.get("authors") and pm_meta.get("year"):
            if not pm_meta.get("abstract") and oa_meta and oa_meta.get("abstract"):
                pm_meta["abstract"] = oa_meta["abstract"]
            pm_meta["is_retracted"] = bool(
                (oa_meta and oa_meta.get("is_retracted"))
                or str(pm_meta.get("title", "")).strip().upper().startswith("RETRACTED:")
            )
            return pm_meta, "pubmed_doi", network_mode

    if len(heuristic_title.split()) >= 5:
        payload, mode = fetch_json_with_proxy_policy(
            "https://api.crossref.org/works",
            params={"query.bibliographic": heuristic_title, "rows": "3"},
        )
        network_mode = mode
        items = ((payload or {}).get("message") or {}).get("items") or []
        norm_target = normalize_title(heuristic_title)
        for cand in items:
            if not isinstance(cand, dict):
                continue
            parsed_cand = parse_crossref_work(cand)
            ratio = difflib.SequenceMatcher(
                None, norm_target, normalize_title(parsed_cand.get("title", ""))
            ).ratio()
            if ratio >= 0.90 and parsed_cand.get("authors") and parsed_cand.get("year"):
                cand_doi = str(parsed_cand.get("doi") or "")
                if cand_doi and not parsed_cand.get("abstract"):
                    oa_cand, _ = fetch_openalex_work(doi=cand_doi)
                    if oa_cand:
                        if oa_cand.get("abstract"):
                            parsed_cand["abstract"] = oa_cand["abstract"]
                        if oa_cand.get("is_retracted"):
                            parsed_cand["is_retracted"] = True
                return parsed_cand, "crossref_title", network_mode

    return {}, "pdf_heuristic", network_mode


def match_library_record(
    parsed: dict[str, object], records: dict[int, dict[str, object]]
) -> dict[str, object] | None:
    rid = RecordIndex(records).match(parsed)
    return records[rid] if rid is not None else None


def corroborates_pdf(parsed, record):
    title = normalize_title(record.get("title"))
    pdf_title = normalize_title(parsed.get("heuristic_title"))
    imprint = normalize_title(parsed.get("first_page_imprint"))
    author = normalize_author_family(parsed.get("first_author_family"))
    record_author = normalize_author_family(record.get("first_author_family"))
    if author and record_author and author != record_author:
        return False
    years = (str(parsed.get("heuristic_year") or ""), str(record.get("year") or ""))
    if all(year.isdigit() for year in years) and abs(int(years[0]) - int(years[1])) > 1:
        return False
    if len(title) >= 12 and (title == pdf_title or f" {title} " in f" {imprint} "):
        return True
    return bool(
        author and author == record_author and parsed.get("heuristic_year") == record.get("year")
        and min(len(title), len(pdf_title)) >= 28
        and difflib.SequenceMatcher(None, title, pdf_title, autojunk=False).ratio() >= .95
    )


def cmd_index(args: argparse.Namespace) -> int:
    pdf_dir = args.pdf_dir.resolve()
    work_dir = check_work_path(args.work_dir)
    output_path = check_work_path(args.output, ".json")
    if not pdf_dir.is_dir():
        raise FileNotFoundError(f"PDF directory not found: {pdf_dir}")
    if output_path.exists() and not args.overwrite:
        raise FileExistsError(f"Refusing to overwrite catalog: {output_path}")
    work_dir.mkdir(parents=True, exist_ok=True)

    records: dict[int, dict[str, object]] = {}
    library_id = None
    snapshot_path_str: str | None = None
    if args.enl is not None:
        enl_path = args.enl.resolve()
        running_endnote = list_process_pids("EndNote.exe")
        if running_endnote:
            raise RuntimeError(
                f"EndNote.exe is running (PIDs={running_endnote}). Close EndNote cleanly before taking a snapshot."
            )
        snap_path = copy_sdb_snapshot(
            enl_path, work_dir / "snapshots" / f"sdb_{uuid4().hex}.eni"
        )
        snapshot_path_str = str(snap_path)
        records = read_endnote_active_records(snap_path)
        library_id = read_library_id(snap_path)
    record_index = RecordIndex(records)

    cache_path = work_dir / "pdf_parse_cache.json"
    if output_path == cache_path.resolve() or output_path.is_relative_to(pdf_dir):
        raise ValueError("Place the output catalog outside the PDF directory and cache path")
    cache: dict[str, dict[str, object]] = {}
    if cache_path.is_file():
        cache = json.loads(cache_path.read_text(encoding="utf-8"))

    pdf_files = sorted(pdf_dir.rglob("*.pdf"), key=lambda p: p.relative_to(pdf_dir).as_posix())
    items: list[dict[str, object]] = []
    skipped_pdfs: list[dict[str, str]] = []
    cache_hits = 0
    max_pages = max(1, int(getattr(args, "max_pages", 30)))

    for idx, pdf_path in enumerate(pdf_files, start=1):
        pdf_id = f"P{idx:03d}"
        rel_posix = pdf_path.relative_to(pdf_dir).as_posix()
        stat = pdf_path.stat()
        digest = hashlib.sha256()
        with pdf_path.open("rb") as handle:
            for block in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(block)
        sha256 = digest.hexdigest().upper()
        cache_key = f"v4|p{max_pages}|{rel_posix}|{sha256}"
        if cache_key in cache:
            parsed = dict(cache[cache_key])
            parsed["pdf_id"] = pdf_id
            parsed["pdf_path"] = str(pdf_path.resolve())
            cache_hits += 1
        else:
            try:
                parsed = extract_pdf_record(pdf_path, pdf_dir, pdf_id, max_pages_limit=max_pages)
            except Exception as exc:
                skipped_pdfs.append({"pdf_id": pdf_id, "relative_path": rel_posix, "error": str(exc)})
                continue
            cache[cache_key] = parsed
        if pdf_path.stat().st_mtime_ns != stat.st_mtime_ns:
            raise RuntimeError(f"PDF changed during indexing: {pdf_path}")

        identity_review = ""
        try:
            rid = record_index.match(parsed)
        except IdentityReview as exc:
            rid = None
            identity_review = str(exc)
        matched = records.get(rid)
        if matched and parsed.get("identifier_layer") in {"page_text", "xmp_xml", "pdf_info"}:
            if not corroborates_pdf(parsed, matched):
                identity_review = "Extracted DOI lacks corroboration from the PDF title"
                matched = None
        if matched is not None:
            abstract = str(matched.get("abstract") or parsed.get("heuristic_abstract") or "")
            lib_title = str(matched["title"] or parsed["heuristic_title"])
            items.append(
                {
                    "pdf_id": pdf_id,
                    "pdf_path": parsed["pdf_path"],
                    "pdf_sha256": sha256,
                    "relative_path": rel_posix,
                    "in_library": True,
                    "record_id": matched["record_id"],
                    "metadata_source": "endnote_library",
                    "network_mode": "not_needed",
                    "identifier_layer": parsed.get("identifier_layer", "none"),
                    "arxiv_id": parsed.get("arxiv_id", ""),
                    "needs_ocr": bool(parsed.get("needs_ocr", False)),
                    "is_retracted": lib_title.strip().upper().startswith("RETRACTED:"),
                    "title": lib_title,
                    "authors": matched["authors"],
                    "first_author_family": matched["first_author_family"],
                    "year": matched["year"] or parsed["heuristic_year"],
                    "journal": matched["journal"],
                    "volume": matched["volume"],
                    "issue": matched["issue"],
                    "pages": matched["pages"],
                    "doi": matched["doi"] or parsed["doi"],
                    "pmid": matched.get("pmid", ""),
                    "abstract": abstract,
                    "passages": parsed["passages"],
                }
            )
        else:
            enriched: dict[str, object] = {}
            source = "pdf_heuristic"
            network_mode = "not_requested"
            if args.enrich_missing_metadata:
                enriched, source, network_mode = enrich_metadata(parsed)
            found_doi = normalize_doi(enriched.get("doi"))
            pdf_doi = normalize_doi(parsed.get("doi"))
            pdf_title = normalize_title(parsed.get("heuristic_title"))
            online_title = normalize_title(enriched.get("title"))
            if enriched and (
                (found_doi and pdf_doi and found_doi != pdf_doi)
                or (pdf_doi and not corroborates_pdf(parsed, enriched))
            ):
                identity_review = "PDF identification and online metadata disagree"
            elif source == "crossref_title" and not (
                pdf_title == online_title
                and normalize_author_family(parsed.get("first_author_family"))
                and normalize_author_family(parsed.get("first_author_family")) == normalize_author_family(enriched.get("first_author_family"))
                and parsed.get("heuristic_year") == enriched.get("year")
            ):
                identity_review = "Title-search metadata lacks exact title, PDF author and year corroboration"
            elif enriched and not pdf_doi and not corroborates_pdf(parsed, enriched):
                identity_review = "Preprint/arXiv online metadata is not corroborated by this PDF"
            elif enriched:
                try:
                    rid = record_index.match({**parsed, **enriched})
                    identity_review = ""
                except IdentityReview as exc:
                    identity_review = str(exc)
            trust_enriched = bool(enriched and not identity_review)
            accepted_enriched = enriched if trust_enriched else {}
            authors = list(accepted_enriched.get("authors") or parsed.get("authors") or [])
            abstract = str(accepted_enriched.get("abstract") or parsed.get("heuristic_abstract") or "")
            new_title = str(accepted_enriched.get("title") or parsed["heuristic_title"])
            items.append(
                {
                    "pdf_id": pdf_id,
                    "pdf_path": parsed["pdf_path"],
                    "pdf_sha256": sha256,
                    "relative_path": rel_posix,
                    "in_library": rid is not None and not identity_review,
                    "record_id": rid if not identity_review else None,
                    "identity_review": identity_review,
                    "metadata_source": source if trust_enriched else "pdf_heuristic",
                    "network_mode": network_mode,
                    "identifier_layer": parsed.get("identifier_layer", "none"),
                    "arxiv_id": parsed.get("arxiv_id", ""),
                    "needs_ocr": bool(parsed.get("needs_ocr", False)),
                    "is_retracted": bool(
                        accepted_enriched.get("is_retracted")
                        or new_title.strip().upper().startswith("RETRACTED:")
                    ),
                    "title": new_title,
                    "authors": authors,
                    "first_author_family": (
                        accepted_enriched.get("first_author_family")
                        or first_author_family(authors)
                        or parsed.get("first_author_family", "")
                    ),
                    "year": accepted_enriched.get("year") or parsed["heuristic_year"],
                    "journal": accepted_enriched.get("journal") or "",
                    "volume": accepted_enriched.get("volume") or "",
                    "issue": accepted_enriched.get("issue") or "",
                    "pages": accepted_enriched.get("pages") or "",
                    "doi": accepted_enriched.get("doi") or parsed["doi"],
                    "pmid": accepted_enriched.get("pmid") or "",
                    "abstract": abstract,
                    "passages": parsed["passages"],
                }
            )

    write_json_atomic(cache_path, cache, overwrite=True)

    catalog = {
        "status": "PASS",
        "pdf_dir": str(pdf_dir),
        "enl_path": str(args.enl.resolve()) if args.enl else None,
        "library_id": library_id,
        "snapshot_path": snapshot_path_str,
        "library_active_count": len(records),
        "max_pages": max_pages,
        "pdf_count": len(items),
        "in_library_count": sum(1 for item in items if item["in_library"]),
        "new_pdf_count": sum(1 for item in items if not item["in_library"]),
        "skipped_pdf_count": len(skipped_pdfs),
        "skipped_pdfs": skipped_pdfs,
        "cache_hits": cache_hits,
        "items": items,
    }
    write_json_atomic(output_path, catalog, overwrite=args.overwrite)
    summary = {k: v for k, v in catalog.items() if k != "items"}
    summary["items_preview"] = [
        {
            "pdf_id": i["pdf_id"],
            "relative_path": i["relative_path"],
            "in_library": i["in_library"],
            "record_id": i["record_id"],
            "metadata_source": i["metadata_source"],
            "identifier_layer": i.get("identifier_layer", "none"),
            "is_retracted": i.get("is_retracted", False),
            "first_author_family": i["first_author_family"],
            "year": i["year"],
            "doi": i["doi"],
            "title": i["title"],
        }
        for i in items
    ]
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


def cmd_search(args: argparse.Namespace) -> int:
    catalog_path = args.catalog.resolve()
    if not catalog_path.is_file():
        raise FileNotFoundError(catalog_path)
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    items = list(catalog.get("items") or [])
    query_tokens = tokenize_scientific(args.query)
    if not query_tokens or not items:
        out = {"query": args.query, "query_tokens": query_tokens, "hits": []}
        print(json.dumps(out, ensure_ascii=False, indent=2))
        return 0

    doc_tfs: list[Counter[str]] = []
    doc_lens: list[int] = []
    df: Counter[str] = Counter()
    for item in items:
        weighted_tokens = (
            tokenize_scientific(str(item.get("title", ""))) * 3
            + tokenize_scientific(str(item.get("abstract", ""))) * 2
            + [
                tok
                for p in item.get("passages") or []
                for tok in tokenize_scientific(str(p.get("text", "")))
            ]
        )
        tf = Counter(weighted_tokens)
        doc_tfs.append(tf)
        doc_lens.append(sum(tf.values()))
        for tok in tf:
            df[tok] += 1

    n_docs = len(items)
    avgdl = sum(doc_lens) / max(n_docs, 1)
    k1 = 1.5
    b = 0.75
    unique_q = list(dict.fromkeys(query_tokens))

    scored: list[tuple[float, dict[str, object]]] = []
    for item, tf, dl in zip(items, doc_tfs, doc_lens):
        if dl == 0 or avgdl == 0:
            continue
        score = 0.0
        matched_terms: list[str] = []
        for qt in unique_q:
            freq = tf.get(qt, 0)
            if freq == 0:
                continue
            matched_terms.append(qt)
            idf = math.log((n_docs - df[qt] + 0.5) / (df[qt] + 0.5) + 1.0)
            score += idf * (freq * (k1 + 1.0)) / (freq + k1 * (1.0 - b + b * (dl / avgdl)))
        if score <= 0.0:
            continue

        # Rank individual passages to return top 2 verbatim supporting excerpts
        passage_scores: list[tuple[float, dict[str, object]]] = []
        for passage in item.get("passages") or []:
            ptoks = set(tokenize_scientific(str(passage.get("text", ""))))
            overlap = [qt for qt in unique_q if qt in ptoks]
            if not overlap:
                continue
            pscore = sum(math.log((n_docs - df[qt] + 0.5) / (df[qt] + 0.5) + 1.0) for qt in overlap)
            passage_scores.append((pscore, {"page": passage["page"], "matched_terms": overlap, "text": passage["text"]}))
        passage_scores.sort(key=lambda pair: pair[0], reverse=True)
        top_passages = [pair[1] for pair in passage_scores[:2]]

        scored.append(
            (
                score,
                {
                    "pdf_id": item["pdf_id"],
                    "score": round(score, 4),
                    "in_library": item["in_library"],
                    "record_id": item["record_id"],
                    "first_author_family": item["first_author_family"],
                    "year": item["year"],
                    "title": item["title"],
                    "journal": item["journal"],
                    "doi": item["doi"],
                    "is_retracted": bool(item.get("is_retracted", False)),
                    "identity_review": item.get("identity_review", ""),
                    "matched_terms": matched_terms,
                    "top_passages": top_passages,
                },
            )
        )

    scored.sort(key=lambda pair: pair[0], reverse=True)
    result = {
        "query": args.query,
        "query_tokens": unique_q,
        "hits": [pair[1] for pair in scored[: args.top_k]],
    }
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        if args.report.resolve() == args.catalog.resolve():
            raise ValueError("Search report and input catalog must differ")
        write_json_atomic(args.report, result, overwrite=args.overwrite)
    print(rendered)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Index a PDF folder against an EndNote library snapshot and search verbatim claim evidence."
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_idx = sub.add_parser("index", help="Scan PDF folder, align with EndNote snapshot, and enrich metadata.")
    p_idx.add_argument("--pdf-dir", required=True, type=Path)
    p_idx.add_argument("--enl", type=Path)
    p_idx.add_argument("--work-dir", required=True, type=Path)
    p_idx.add_argument("--output", required=True, type=Path)
    p_idx.add_argument("--max-pages", type=int, default=30, help="Maximum pages indexed per PDF (default: 30).")
    p_idx.add_argument("--enrich-missing-metadata", action="store_true")
    p_idx.add_argument("--overwrite", action="store_true")

    p_search = sub.add_parser("search", help="Search indexed PDF catalog by BM25 and extract top passages.")
    p_search.add_argument("--catalog", required=True, type=Path)
    p_search.add_argument("--query", required=True)
    p_search.add_argument("--top-k", type=int, default=5)
    p_search.add_argument("--report", type=Path)
    p_search.add_argument("--overwrite", action="store_true")

    return parser


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    args = build_parser().parse_args()
    if args.command == "index":
        return cmd_index(args)
    if args.command == "search":
        return cmd_search(args)
    raise ValueError(f"Unknown command: {args.command}")


if __name__ == "__main__":
    raise SystemExit(main())
