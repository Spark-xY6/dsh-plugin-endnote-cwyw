"""Check a proposed quotation on its physical PDF page before library import."""

from contextlib import ExitStack
import re
import unicodedata

import fitz


_EVIDENCE_PUNCT_TRANS = str.maketrans({
    "‘": "'", "’": "'", "‚": "'",
    "“": '"', "”": '"', "„": '"',
    "–": "-", "—": "-", "−": "-",
    "\u00a0": " ",
})


def page_text(value):
    value = unicodedata.normalize("NFKC", value).replace("\u00ad", "").translate(_EVIDENCE_PUNCT_TRANS)
    return " ".join(re.sub(r"(?<=\w)-[ \t]*\r?\n[ \t]*(?=\w)", "", value).split())


def _parse_page_spec(page: object) -> tuple[int, int] | None:
    if type(page) is int:
        return (page, page) if page >= 1 else None
    if isinstance(page, str):
        m = re.fullmatch(r"(\d+)-(\d+)", page.strip())
        if m:
            start, end = int(m.group(1)), int(m.group(2))
            if start >= 1 and start <= end <= start + 5:
                return (start, end)
    return None


def verify_plan_evidence(plan, selected):
    articles = {item["pdf_id"]: item for item in selected}
    verified = []
    pages = {}
    with ExitStack() as stack:
        documents = {}
        for index, placement in enumerate(plan["insertions"], 1):
            label = placement.get("insertion_id") or f"C{index:02d}"
            sources = placement.get("sources")
            pdf_ids = set(placement["pdf_ids"])
            rationale = str(placement.get("rationale") or "").strip()
            if not isinstance(sources, list) or not pdf_ids or not rationale:
                raise ValueError(f"{label}: supply a rationale and page evidence for each selected PDF")
            seen = set()
            for source in sources:
                if not isinstance(source, dict) or "pdf_id" not in source or "page" not in source:
                    raise ValueError(f"{label}: source evidence must be an object with 'pdf_id' and 'page'")
                pdf_id, page = source["pdf_id"], source["page"]
                page_span = _parse_page_spec(page)
                quote = source.get("quote")
                if pdf_id not in pdf_ids or pdf_id not in articles or page_span is None or not isinstance(quote, str) or not quote.strip():
                    raise ValueError(f"{label}: invalid source, physical PDF page or quotation")
                if pdf_id not in documents:
                    documents[pdf_id] = stack.enter_context(fitz.open(articles[pdf_id]["pdf_path"]))
                document = documents[pdf_id]
                p_start, p_end = page_span
                if p_end > len(document):
                    raise ValueError(f"{label}: {pdf_id} has no physical page {page}")
                span_parts = []
                for pno in range(p_start, p_end + 1):
                    key = (pdf_id, pno)
                    if key not in pages:
                        pages[key] = page_text(document[pno - 1].get_text("text"))
                    span_parts.append(pages[key])
                haystack = " ".join(span_parts)
                fragment = page_text(quote)
                if not fragment or fragment not in haystack:
                    raise ValueError(f"{label}: quotation not found on {pdf_id} page {page}")
                seen.add(pdf_id)
                verified.append({
                    "insertion_id": label,
                    "pdf_id": pdf_id,
                    "page": page,
                    "total_pages": len(document),
                    "quote": fragment,
                    "rationale": rationale,
                })
            if seen != pdf_ids:
                raise ValueError(f"{label}: missing page evidence for {sorted(pdf_ids - seen)}")
    return verified
