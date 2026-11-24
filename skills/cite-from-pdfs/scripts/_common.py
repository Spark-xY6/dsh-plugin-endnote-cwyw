from __future__ import annotations

import ctypes
from contextlib import closing
from ctypes import wintypes
import html
import os
import re
import shutil
import sqlite3
import tempfile
import unicodedata
import winreg
from pathlib import Path
from urllib.parse import unquote, urlsplit


STOPWORDS = {
    "the", "a", "an", "and", "or", "but", "of", "in", "on", "at", "to", "for",
    "by", "with", "from", "as", "is", "are", "was", "were", "be", "been", "being",
    "that", "this", "these", "those", "it", "its", "their", "our", "we", "they",
    "can", "may", "might", "could", "would", "should", "has", "have", "had",
    "into", "across", "between", "through", "during", "while", "whereas", "which",
}


def text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return str(value)


_GREEK_TO_LATIN = str.maketrans({
    "α": "alpha", "β": "beta", "γ": "gamma", "δ": "delta",
    "ε": "epsilon", "ζ": "zeta", "η": "eta", "θ": "theta",
    "ι": "iota", "κ": "kappa", "λ": "lambda", "μ": "mu", "µ": "mu",
    "ν": "nu", "ξ": "xi", "ο": "omicron", "π": "pi",
    "ρ": "rho", "σ": "sigma", "τ": "tau", "υ": "upsilon",
    "φ": "phi", "χ": "chi", "ψ": "psi", "ω": "omega",
})
_SURNAME_PARTICLES = {"van", "von", "de", "del", "della", "der", "den", "la", "le", "du", "di", "da", "dos", "st."}
_GEN_SUFFIX_RE = re.compile(r"(?:,?\s+(?:Jr\.?|Sr\.?|II|III|IV))$")
_DOI_URL_RE = re.compile(r"https?://(?:dx\.)?doi\.org/[^\s\"'<>]+", re.IGNORECASE)
_DOI_SEARCH_RE = re.compile(r"(?:doi:\s*)?(10\.\d{4,9}/[^\s\"'<>]+)", re.IGNORECASE)


def normalize_doi(value: object) -> str:
    raw = text(value).strip()
    if not raw:
        return ""
    for line in re.split(r"[\r\n]+", raw):
        chunk = line.strip()
        if not chunk:
            continue
        url_match = _DOI_URL_RE.search(chunk)
        candidate = ""
        if url_match:
            candidate = unquote(urlsplit(url_match.group(0)).path.lstrip("/"))
        else:
            match = _DOI_SEARCH_RE.search(chunk)
            if match:
                candidate = match.group(1)
        if not candidate:
            continue
        candidate = candidate.rstrip(".,;]}\"'<>")
        while candidate.endswith(")") and candidate.count(")") > candidate.count("("):
            candidate = candidate[:-1]
        if re.fullmatch(r"10\.\d{4,9}/[^\s\"<>]+", candidate):
            return candidate.casefold()
    return ""


def normalize_title(value: object) -> str:
    raw = html.unescape(text(value))
    raw = re.sub(r"</?[A-Za-z][^>]*>", "", raw).casefold().translate(_GREEK_TO_LATIN)
    normalized = unicodedata.normalize("NFKD", raw)
    return " ".join(
        "".join(
            char if char.isalnum() else " "
            for char in normalized
            if not unicodedata.combining(char)
        ).split()
    )


def split_name_with_particles(full_name: str) -> dict[str, str]:
    cleaned = _GEN_SUFFIX_RE.sub("", full_name.strip()).strip()
    parts = cleaned.split()
    if len(parts) > 1:
        idx = len(parts) - 1
        while idx > 1 and parts[idx - 1].casefold() in _SURNAME_PARTICLES:
            idx -= 1
        return {"family": " ".join(parts[idx:]), "given": " ".join(parts[:idx])}
    return {"family": cleaned, "given": ""}


def normalize_author_family(family: object) -> str:
    norm = normalize_title(family)
    if not norm:
        return ""
    parts = norm.split()
    while len(parts) > 1 and parts[0] in {"van", "von", "de", "del", "della", "der", "den", "la", "le", "du", "di", "da", "dos", "st"}:
        parts = parts[1:]
    return "".join(parts)


def parse_endnote_authors(author_field: object) -> list[dict[str, str]]:
    """Parse EndNote's CR/LF-separated author lines into family/given dicts."""
    raw = text(author_field).replace("\r\n", "\n").replace("\r", "\n")
    authors: list[dict[str, str]] = []
    for line in raw.split("\n"):
        cleaned = _GEN_SUFFIX_RE.sub("", line.strip()).strip()
        if not cleaned:
            continue
        if "," in cleaned:
            family, given = cleaned.split(",", 1)
            authors.append({"family": family.strip(), "given": _GEN_SUFFIX_RE.sub("", given.strip()).strip()})
        else:
            authors.append(split_name_with_particles(cleaned))
    return authors


def parse_pdf_meta_authors(author_field: object) -> list[dict[str, str]]:
    """Parse single-line PDF Info /Author or XMP creator strings containing multiple authors."""
    raw = text(author_field).replace("\r\n", "\n").replace("\r", "\n").strip()
    if not raw:
        return []
    if "\n" in raw or ";" in raw:
        return parse_endnote_authors(raw.replace(";", "\n"))
    if re.search(r"\s+(?:and|&)\s+", raw, re.IGNORECASE) and "," not in raw:
        return parse_endnote_authors(re.sub(r"\s+(?:and|&)\s+", "\n", raw, flags=re.IGNORECASE))
    if "," in raw:
        segments = [s.strip() for s in raw.split(",") if s.strip()]
        # Clean leading "and "/"& " on the final comma segment ("A. Smith, B. Jones, and C. Lee")
        segments = [re.sub(r"^(?:and|&)\s+", "", s, flags=re.IGNORECASE).strip() for s in segments]
        segments = [s for s in segments if s]
        # Case 1: every segment has >= 2 words -> "Alice Smith, Bob Jones"
        if len(segments) >= 2 and all(len(s.split()) >= 2 for s in segments):
            return parse_endnote_authors("\n".join(segments))
        # Case 2: paired "Smith, A., Jones, B." (even count >= 4, second part of each pair looks like initials/given)
        if len(segments) >= 4 and len(segments) % 2 == 0 and all(len(segments[i].split()) <= 2 for i in range(0, len(segments), 2)):
            paired = [f"{segments[i]}, {segments[i + 1]}" for i in range(0, len(segments), 2)]
            return parse_endnote_authors("\n".join(paired))
    return parse_endnote_authors(raw)


def first_author_family(authors: list[dict[str, str]]) -> str:
    if not authors:
        return ""
    return str(authors[0].get("family", "")).strip()


def tokenize_scientific(value: str) -> list[str]:
    """Tokenize scientific text while preserving hyphenated and sub-parts."""
    raw_tokens = re.findall(r"[a-z0-9]+(?:[-'][a-z0-9]+)*", value.casefold())
    tokens: list[str] = []
    for tok in raw_tokens:
        if tok in STOPWORDS or len(tok) <= 1:
            continue
        tokens.append(tok)
        if "-" in tok:
            for sub in tok.split("-"):
                if sub and sub not in STOPWORDS and len(sub) > 1:
                    tokens.append(sub)
    return tokens


def resolve_endnote_exe() -> Path:
    queries = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe"),
        (winreg.HKEY_CLASSES_ROOT, r"EndNotePlus.ENW\shell\open\command"),
    ]
    for hive, subkey in queries:
        try:
            with winreg.OpenKey(hive, subkey) as key:
                raw, _ = winreg.QueryValueEx(key, "")
        except OSError:
            continue
        cleaned = os.path.expandvars(str(raw).strip())
        if not cleaned:
            continue
        if cleaned.startswith('"'):
            parts = cleaned.split('"', 2)
            cand = Path(parts[1]) if len(parts) >= 2 else None
        elif Path(cleaned).is_file():
            cand = Path(cleaned)
        else:
            match = re.match(r"^(.*?\.exe)(?:\s|$)", cleaned, flags=re.IGNORECASE)
            cand = Path(match.group(1)) if match else None
        if cand is not None and cand.is_file():
            return cand
    raise RuntimeError("Could not locate EndNote.exe in Windows Registry")


def list_process_pids(image_name: str) -> list[int]:
    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD),
            ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.c_void_p),
            ("th32ModuleID", wintypes.DWORD),
            ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", wintypes.DWORD),
            ("szExeFile", wintypes.WCHAR * 260),
        ]

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateToolhelp32Snapshot.argtypes = (wintypes.DWORD, wintypes.DWORD)
    kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    for enumerate_process in (kernel32.Process32FirstW, kernel32.Process32NextW):
        enumerate_process.argtypes = (wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W))
        enumerate_process.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
    kernel32.CloseHandle.restype = wintypes.BOOL
    snap = kernel32.CreateToolhelp32Snapshot(0x00000002, 0)
    if snap in (None, ctypes.c_void_p(-1).value):
        raise ctypes.WinError(ctypes.get_last_error())
    target = image_name.casefold()
    pids: list[int] = []
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        ok = kernel32.Process32FirstW(snap, ctypes.byref(entry))
        if not ok:
            raise ctypes.WinError(ctypes.get_last_error())
        while ok:
            if entry.szExeFile.casefold() == target:
                pids.append(int(entry.th32ProcessID))
            ok = kernel32.Process32NextW(snap, ctypes.byref(entry))
        if ctypes.get_last_error() != 18:
            raise ctypes.WinError(ctypes.get_last_error())
    finally:
        kernel32.CloseHandle(snap)
    return pids


def sdb_path_for_enl(enl_path: Path) -> Path:
    resolved = enl_path.resolve()
    return resolved.parent / f"{resolved.stem}.Data" / "sdb" / "sdb.eni"


def check_work_path(path: Path, suffix: str = "") -> Path:
    chosen = path.resolve()
    if any(part.name.casefold().endswith(".data") for part in (path, *path.parents, chosen, *chosen.parents)):
        raise ValueError(f"Work and output paths must be outside EndNote .Data directories: {path}")
    if suffix and (chosen.suffix.casefold() != suffix or path.suffix.casefold() != suffix):
        raise ValueError(f"Expected a {suffix} output: {path}")
    return chosen


def write_json_atomic(path: Path, value, *, overwrite=False):
    import json

    path = check_work_path(path, ".json")
    if path.exists() and not overwrite:
        raise FileExistsError(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, delete=False) as handle:
        temporary = Path(handle.name)
        try:
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        except BaseException:
            handle.close()
            temporary.unlink()
            raise
    try:
        temporary.replace(path) if overwrite else temporary.rename(path)
    finally:
        temporary.unlink(missing_ok=True)


def copy_sdb_snapshot(enl_path: Path, dst_path: Path) -> Path:
    dst_path = check_work_path(dst_path)
    src = sdb_path_for_enl(enl_path)
    if not src.is_file():
        raise FileNotFoundError(f"EndNote database not found: {src}")
    for suffix in ("-wal", "-journal"):
        pending = src.with_name(src.name + suffix)
        if pending.is_file() and pending.stat().st_size:
            raise RuntimeError(f"Close EndNote and checkpoint the pending database file: {pending}")
    dst_path.parent.mkdir(parents=True, exist_ok=True)
    with dst_path.open("xb") as target, src.open("rb") as source:
        shutil.copyfileobj(source, target)
    return dst_path


def read_endnote_active_records(sdb_snapshot: Path) -> dict[int, dict[str, object]]:
    """Read active (non-trashed) records from a frozen sdb.eni snapshot."""
    if sdb_snapshot.parent.name.casefold() == "sdb" and sdb_snapshot.parent.parent.name.casefold().endswith(".data"):
        raise RuntimeError("Refusing live sdb.eni path; copy snapshot into work_dir first.")
    uri = sdb_snapshot.resolve().as_uri() + "?mode=ro&immutable=1"
    conn = sqlite3.connect(uri, uri=True)
    try:
        cur = conn.cursor()
        cols = {row[1] for row in cur.execute('PRAGMA table_info("refs")').fetchall()}
        wanted = [
            "id", "author", "year", "title", "secondary_title",
            "volume", "pages", "number", "electronic_resource_number",
            "abstract", "url", "accession_number", "name_of_database", "database_provider",
        ]
        selected = [c for c in wanted if c in cols]
        where = " WHERE COALESCE(trash_state, 0) = 0" if "trash_state" in cols else ""
        sql = f"SELECT {', '.join(f'\"{c}\"' for c in selected)} FROM refs{where}"
        records: dict[int, dict[str, object]] = {}
        for row in cur.execute(sql).fetchall():
            raw = {col: text(val) for col, val in zip(selected, row)}
            rid = int(row[0])
            authors = parse_endnote_authors(raw.get("author", ""))
            records[rid] = {
                "record_id": rid,
                "authors": authors,
                "first_author_family": first_author_family(authors),
                "year": raw.get("year", "").strip(),
                "title": raw.get("title", "").strip(),
                "title_norm": normalize_title(raw.get("title", "")),
                "journal": raw.get("secondary_title", "").strip(),
                "volume": raw.get("volume", "").strip(),
                "issue": raw.get("number", "").strip(),
                "pages": raw.get("pages", "").strip(),
                "doi": normalize_doi(raw.get("electronic_resource_number", "")),
                "abstract": raw.get("abstract", "").strip(),
                "url": raw.get("url", "").strip(),
                "accession_number": raw.get("accession_number", "").strip(),
                "pmid": pubmed_id(raw),
            }
        return records
    finally:
        conn.close()


def pubmed_id(item: dict[str, object]) -> str:
    pmid = text(item.get("pmid")).strip()
    if re.fullmatch(r"[1-9]\d{0,8}", pmid):
        return pmid
    accession = text(item.get("accession_number")).strip()
    tagged = re.fullmatch(r"PMID\s*:?\s*([1-9]\d{0,8})", accession, re.IGNORECASE)
    if tagged:
        return tagged.group(1)
    database = text(item.get("name_of_database")).casefold()
    if database in {"pubmed", "medline"} and re.fullmatch(r"[1-9]\d{0,8}", accession):
        return accession
    return ""


def read_library_id(sdb_snapshot: Path) -> str:
    """Identify the closed EndNote 21 library from its frozen database."""
    if sdb_snapshot.parent.name.casefold() == "sdb":
        raise RuntimeError("Use an isolated database snapshot")
    with closing(sqlite3.connect(sdb_snapshot.resolve().as_uri() + "?mode=ro&immutable=1", uri=True)) as conn:
        row = conn.execute("SELECT value FROM misc WHERE code=14 AND subcode=0").fetchone()
    library_id = text(row[0]).strip() if row else ""
    if not re.fullmatch(r"[A-Za-z0-9-]{8,100}", library_id):
        raise RuntimeError("Could not establish the EndNote library ID")
    return library_id
