"""Read Traveling Library records and prepare a native CWYW rebind on a copy."""

from __future__ import annotations

import base64
from dataclasses import dataclass
import re
from zipfile import ZipFile

from lxml import etree

from audit_cwyw_docx import TEMPORARY_CITATION_RE, parse_fields


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
DOCUMENT = "word/document.xml"
SETTINGS = "word/settings.xml"


def parse_xml(data):
    return etree.fromstring(data, parser=etree.XMLParser(resolve_entities=False, no_network=True))


@dataclass
class CitationGroup:
    runs: list
    records: list
    has_overrides: bool


def traveling_record(cite, key):
    record = cite.find("record")
    if record is None or not cite.findtext("RecNum", "").isdigit():
        raise ValueError("Existing CWYW citation lacks Traveling Library record metadata")
    rid = int(cite.findtext("RecNum"))
    keys = {
        node.get("db-id") for node in record.findall("./foreign-keys/key")
        if node.get("app") == "EN" and node.text == str(rid) and node.get("db-id")
    }
    authors = []
    for node in record.findall("./contributors/authors/author"):
        name = "".join(node.itertext()).strip()
        family, sep, given = name.partition(",")
        if not sep and " " in family:
            given, family = family.rsplit(" ", 1)
        if family.strip():
            authors.append({"family": family.strip(), "given": given.strip()})
    title = record.find("./titles/title")
    return {
        "pdf_id": key, "original_id": rid, "library_ids": sorted(keys),
        "title": "".join(title.itertext()).strip() if title is not None else "",
        "authors": authors, "first_author_family": authors[0]["family"] if authors else cite.findtext("Author", ""),
        "year": record.findtext("./dates/year") or cite.findtext("Year", ""),
        "doi": record.findtext("electronic-resource-num", ""),
        "accession_number": record.findtext("accession-num", ""),
        "name_of_database": record.findtext("remote-database-name", ""),
        "journal": record.findtext("./periodical/full-title") or record.findtext("./titles/secondary-title", ""),
        "volume": record.findtext("volume", ""), "issue": record.findtext("number", ""),
        "pages": record.findtext("pages", ""),
        "reference_type": record.find("ref-type").get("name", "") if record.find("ref-type") is not None else "",
    }


def document_citations(path, *, prepared=False):
    with ZipFile(path) as package:
        data = package.read(DOCUMENT)
        notes = [package.read(part) for part in ("word/footnotes.xml", "word/endnotes.xml") if part in package.namelist()]
    root = parse_xml(data)
    if any(b"EN.CITE" in part for part in notes):
        raise ValueError("CWYW citations in Word notes require separate placement review")
    codes = root.xpath(".//w:instrText/text()|.//w:fldSimple/@w:instr", namespaces={"w": W[1:-1]})
    if any(re.match(r"^\s*(?:CITATION\b|BIBLIOGRAPHY\b|ADDIN\s+(?:ZOTERO|MENDELEY|CSL_CITATION))", code, re.I) for code in codes):
        raise ValueError("Existing Word or third-party citation fields require conversion first")
    baseline = parse_fields(data)
    if not prepared and baseline["temporary_citations_remaining"]:
        raise ValueError("Source contains unformatted citations; resolve them before adding new ones")
    groups = []
    positions = []
    stack = []
    temporary_groups = 0
    for p_number, paragraph in enumerate(root.iter(W + "p"), 1):
        visible_offset = 0
        for run in paragraph.findall(W + "r"):
            for field in stack:
                if field["runs"][-1] is not run:
                    field["runs"].append(run)
            for child in run:
                if child.tag == W + "instrText" and stack:
                    stack[-1]["code"].append(child.text or "")
                if child.tag == W + "t" and not stack:
                    text = child.text or ""
                    offset = 0
                    for match in TEMPORARY_CITATION_RE.finditer(text):
                        visible_offset += len("".join(text[offset:match.start()].split()))
                        ids = tuple(sorted(int(rid) for rid in re.findall(r"#(\d+)", match.group())))
                        positions.append((p_number, visible_offset, ids))
                        temporary_groups += 1
                        offset = match.end()
                    visible_offset += len("".join(text[offset:].split()))
                if child.tag != W + "fldChar":
                    continue
                kind = child.get(W + "fldCharType")
                if kind == "begin":
                    stack.append({
                        "runs": [run], "code": [], "data": child.findtext(W + "fldData", ""),
                        "position": (p_number, visible_offset),
                    })
                elif kind == "end" and stack:
                    field = stack.pop()
                    code = "".join(field["code"]).strip()
                    if not re.match(r"^ADDIN\s+EN\.CITE(?:\s|<|$)", code):
                        continue
                    raw = code
                    if "<EndNote>" not in raw and field["data"]:
                        raw = base64.b64decode("".join(field["data"].split()), validate=True).decode("utf-8")
                    match = re.search(r"<EndNote>.*?</EndNote>", raw, re.S)
                    if not match or stack:
                        raise ValueError("Unsupported existing EndNote citation field")
                    cites = parse_xml(match.group().encode("utf-8")).findall("Cite")
                    if not cites:
                        raise ValueError("Existing EndNote citation group is empty")
                    group_no = len(groups) + 1
                    records = [traveling_record(cite, f"existing:{group_no}:{i}") for i, cite in enumerate(cites, 1)]
                    overriden = any(
                        child.tag not in {"Author", "Year", "RecNum", "DisplayText", "record", "foreign-keys"}
                        for cite in cites for child in cite
                    )
                    groups.append(CitationGroup(field["runs"], records, overriden))
                    positions.append((*field["position"], tuple(sorted(record["original_id"] for record in records))))
    if len(groups) != baseline["main_citation_fields"] or not baseline["balanced_field_stack"]:
        raise ValueError("Existing citations include unsupported field boundaries or tracked revisions")
    if temporary_groups != baseline["temporary_citations_remaining"]:
        raise ValueError("Unformatted citations span unsupported runs or revisions")
    return groups, positions


def read_citations(path):
    return document_citations(path)[0]


def prepare_rebind(source, target, groups, resolved):
    rows = resolved["by_pdf_id"]
    mappings = [(record, rows[record["pdf_id"]]) for group in groups for record in group.records]
    if all(
        record["library_ids"] == [resolved["library_id"]] and record["original_id"] == row["record_id"]
        for record, row in mappings
    ):
        return False
    if any(group.has_overrides for group in groups):
        raise ValueError("Existing citation display overrides require review before rebinding")
    root = groups[0].runs[0].getroottree()
    for group in reversed(groups):
        start, *_, end = group.runs
        parent = start.getparent()
        if parent is not end.getparent():
            raise ValueError("Cross-paragraph citation field requires manual rebinding")
        span = list(parent)[parent.index(start):parent.index(end) + 1]
        if span != group.runs or any(
            child.tag not in {W + tag for tag in ("rPr", "fldChar", "instrText", "t")}
            for run in span for child in run
        ):
            raise ValueError("Citation field contains protected non-citation content")
        if [child.tag for child in start if child.tag != W + "rPr"] != [W + "fldChar"] or [
            child.tag for child in end if child.tag != W + "rPr"
        ] != [W + "fldChar"]:
            raise ValueError("Citation shares its boundary run with other content")
        tokens = list(dict.fromkeys(rows[record["pdf_id"]]["citation_token"] for record in group.records))
        replacement = etree.Element(W + "r")
        etree.SubElement(replacement, W + "t", {XML_SPACE: "preserve"}).text = "{" + "; ".join(tokens) + "}"
        parent.insert(parent.index(start), replacement)
        for run in span:
            parent.remove(run)
    changed = {DOCUMENT: etree.tostring(root, encoding="utf-8", xml_declaration=True, standalone=True)}
    with ZipFile(source) as original, ZipFile(target, mode="x") as output:
        if SETTINGS in original.namelist():
            settings = parse_xml(original.read(SETTINGS))
            for var in settings.iter(W + "docVar"):
                if var.get(W + "name") == "EN.Libraries":
                    var.getparent().remove(var)
                    break
            changed[SETTINGS] = etree.tostring(settings, encoding="utf-8", xml_declaration=True, standalone=True)
        for info in original.infolist():
            output.writestr(info, changed.get(info.filename, original.read(info.filename)))
    return True
