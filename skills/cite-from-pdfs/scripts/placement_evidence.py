"""Check a proposed quotation on its physical PDF page before library import."""

from contextlib import ExitStack
import re
import unicodedata

import fitz


def page_text(value):
    value = unicodedata.normalize("NFKC", value).replace("\u00ad", "")
    return " ".join(re.sub(r"(?<=\w)-[ \t]*\n[ \t]*(?=\w)", "", value).split())


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
            if not isinstance(sources, list) or not pdf_ids or not str(placement.get("rationale") or "").strip():
                raise ValueError(f"{label}: supply a rationale and page evidence for each selected PDF")
            seen = set()
            for source in sources:
                pdf_id, page = source["pdf_id"], source["page"]
                quote = source.get("quote")
                if pdf_id not in pdf_ids or type(page) is not int or page < 1 or not isinstance(quote, str) or not quote.strip():
                    raise ValueError(f"{label}: invalid source, physical PDF page or quotation")
                if pdf_id not in documents:
                    documents[pdf_id] = stack.enter_context(fitz.open(articles[pdf_id]["pdf_path"]))
                document = documents[pdf_id]
                if page > len(document):
                    raise ValueError(f"{label}: {pdf_id} has no physical page {page}")
                key = (pdf_id, page)
                if key not in pages:
                    pages[key] = page_text(document[page - 1].get_text("text"))
                fragment = page_text(quote)
                if not fragment or fragment not in pages[key]:
                    raise ValueError(f"{label}: quotation not found on {pdf_id} page {page}")
                seen.add(pdf_id)
                verified.append({"insertion_id": label, "pdf_id": pdf_id, "page": page})
            if seen != pdf_ids:
                raise ValueError(f"{label}: missing page evidence for {sorted(pdf_ids - seen)}")
    return verified
