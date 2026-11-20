"""Conservative bibliographic identity shared by indexing, sync and Word rebinding."""

from dataclasses import dataclass
from difflib import SequenceMatcher
import re

from _common import normalize_doi, normalize_title, pubmed_id, text


@dataclass(frozen=True)
class Identity:
    doi: str
    pmid: str
    title: str
    author: str
    year: str
    journal: str
    volume: str
    pages: str

    @classmethod
    def from_record(cls, item):
        year = text(item.get("year") or item.get("heuristic_year")).strip()
        return cls(
            normalize_doi(item.get("doi")), pubmed_id(item),
            normalize_title(item.get("title") or item.get("heuristic_title")),
            normalize_title(item.get("first_author_family")).replace(" ", ""),
            year if re.fullmatch(r"(?:18|19|20)\d{2}", year) else "",
            normalize_title(item.get("journal")),
            text(item.get("volume")).strip().casefold(),
            normalize_title(item.get("pages")).split(" ", 1)[0],
        )


def compare(left: Identity, right: Identity) -> str:
    same_doi = bool(left.doi and left.doi == right.doi)
    same_pmid = bool(left.pmid and left.pmid == right.pmid)
    doi_conflict = bool(left.doi and right.doi and not same_doi)
    pmid_conflict = bool(left.pmid and right.pmid and not same_pmid)
    same_author = bool(left.author and left.author == right.author)
    same_year = bool(left.year and left.year == right.year)
    compact_left, compact_right = left.title.replace(" ", ""), right.title.replace(" ", "")
    same_title = bool(compact_left and compact_left == compact_right)
    sufficient_title = min(len(compact_left), len(compact_right)) >= 28
    ratio = 0.0
    if (
        not same_title and sufficient_title
        and (same_author or same_doi or same_pmid)
        and abs(len(compact_left) - len(compact_right)) <= max(len(compact_left), len(compact_right)) * .12
    ):
        ratio = SequenceMatcher(None, compact_left, compact_right, autojunk=False).ratio()
    if same_doi or same_pmid:
        shorter, longer = sorted((left.title, right.title), key=len)
        expanded_title = longer.startswith(shorter + " ")
        incompatible_attribution = (
            left.author and right.author and not same_author and left.year and right.year
            and abs(int(left.year) - int(right.year)) > 1
        )
        if doi_conflict or pmid_conflict or incompatible_attribution or (
            sufficient_title and not same_title and not expanded_title and ratio < .87
        ):
            return "conflict"
        return "doi" if same_doi else "pmid"
    if doi_conflict or pmid_conflict:
        return "conflict" if same_title or (same_author and ratio >= .94) else "different"
    same_location = bool(
        left.volume and left.volume == right.volume and left.pages and left.pages == right.pages
    )
    same_journal = bool(left.journal and left.journal == right.journal)
    contradictory = (
        (left.volume and right.volume and left.volume != right.volume)
        or (left.pages and right.pages and left.pages != right.pages)
        or (left.journal and right.journal and not same_journal and not same_location)
    )
    if contradictory and (same_title or (same_author and ratio >= .91)):
        return "ambiguous"
    if same_title and same_author and same_year and (sufficient_title or same_location):
        return "title_author_year"
    if same_author and same_year and ratio >= .96 and (same_location or same_journal):
        return "title_variant"
    if same_title or (same_author and ratio >= .91):
        return "ambiguous"
    return "different"


class IdentityReview(ValueError):
    pass


class RecordIndex:
    def __init__(self, records):
        self.entries = [(rid, Identity.from_record(item)) for rid, item in records.items()]

    def match(self, item, preferred=()):
        query = Identity.from_record(item)
        matches, conflicts, ambiguous = [], [], []
        for rid, record in self.entries:
            outcome = compare(query, record)
            if outcome == "conflict":
                conflicts.append(rid)
            elif outcome == "ambiguous":
                ambiguous.append(rid)
            elif outcome != "different":
                matches.append((rid, record, outcome))
        if conflicts or (ambiguous and not matches):
            raise IdentityReview(
                f"{item.get('title') or item.get('heuristic_title')!r}: "
                f"conflicting matches {conflicts}; ambiguous matches {ambiguous}"
            )
        if not matches:
            return None
        for idx, (_, record, _) in enumerate(matches):
            for _, other, _ in matches[:idx]:
                if compare(record, other) not in {"doi", "pmid", "title_author_year", "title_variant"}:
                    raise IdentityReview(f"Non-equivalent candidate IDs: {[rid for rid, _, _ in matches]}")
        preferred = set(preferred)
        rank = {"doi": 0, "pmid": 1, "title_author_year": 2, "title_variant": 3}
        return min(matches, key=lambda match: (match[0] not in preferred, rank[match[2]], match[0]))[0]
