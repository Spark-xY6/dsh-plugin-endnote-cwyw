# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
# ]
# ///

from __future__ import annotations

import argparse
import ctypes
from contextlib import closing
from ctypes import wintypes
import hashlib
import json
import os
import re
import shutil
import sqlite3
import stat
import subprocess
import sys
import time
import tempfile
from pathlib import Path
from uuid import uuid4

from _common import (
    check_work_path,
    copy_sdb_snapshot,
    list_process_pids,
    read_endnote_active_records,
    read_library_id,
    resolve_endnote_exe,
    sdb_path_for_enl,
)
from reference_identity import Identity, IdentityReview, RecordIndex, compare


sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "endnote-cwyw/scripts"))
from endnote_window import TargetWindow, WindowInterference


BM_CLICK = 0x00F5
CB_GETCURSEL = 0x0147
CB_GETLBTEXT = 0x0148
CB_GETLBTEXTLEN = 0x0149
CB_FINDSTRINGEXACT = 0x0158
CB_SETCURSEL = 0x014E
CBN_SELCHANGE = 1
CBN_SELENDOK = 9
WM_COMMAND = 0x0111


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def require_closed_endnote():
    running = list_process_pids("EndNote.exe")
    if running:
        raise RuntimeError(f"Close EndNote cleanly before accessing a library copy (PIDs={running})")


def verify_regular_path(path):
    info = path.lstat()
    if path.is_symlink() or getattr(info, "st_file_attributes", 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT:
        raise ValueError(f"Library contains a symlink or reparse point: {path}")
    if not (stat.S_ISDIR(info.st_mode) or stat.S_ISREG(info.st_mode)):
        raise ValueError(f"Unsupported filesystem entry in library: {path}")


def tree_hashes(root):
    if not root.exists():
        return {}
    verify_regular_path(root)
    files = {}
    for directory, dirs, names in os.walk(root):
        parent = Path(directory)
        for name in dirs + names:
            verify_regular_path(parent / name)
        for name in names:
            path = parent / name
            files[path.relative_to(root).as_posix()] = sha256_file(path)
    return files


def read_contents(snapshot):
    with closing(sqlite3.connect(snapshot.resolve().as_uri() + "?mode=ro&immutable=1", uri=True)) as conn:
        groups = dict((rid, (spec, members)) for rid, spec, members in conn.execute("SELECT group_id, spec, members FROM groups"))
        attachments = set(conn.execute("SELECT refs_id, file_path, file_type, file_pos FROM file_res"))
    return groups, attachments


def verify_library_contents(before, final, candidate, imported, old_pdf_hashes):
    before_groups, before_attachments = read_contents(before)
    final_groups, final_attachments = read_contents(final)
    if before_groups != final_groups or not before_attachments.issubset(final_attachments):
        raise RuntimeError("A pre-existing group or attachment link changed")
    pdf_root = (candidate.with_suffix(".Data") / "PDF").resolve()
    pdf_hashes = tree_hashes(pdf_root)
    if any(pdf_hashes.get(path) != digest for path, digest in old_pdf_hashes.items()):
        raise RuntimeError("A pre-existing PDF attachment changed")
    for rid, item in imported.items():
        if not item.get("pdf_path"):
            continue
        candidates = []
        for owner, relative, kind, _ in final_attachments:
            if owner == rid and kind == 1:
                path = (pdf_root / relative).resolve()
                if path.is_relative_to(pdf_root):
                    candidates.append(pdf_hashes.get(path.relative_to(pdf_root).as_posix()))
        if item["pdf_sha256"] not in candidates:
            raise RuntimeError(f"Imported record #{rid} lacks its verified PDF attachment")


def create_library_candidate(source, candidate_dir):
    require_closed_endnote()
    source = source.resolve(strict=True)
    candidate_dir = check_work_path(candidate_dir)
    data = sdb_path_for_enl(source).parents[1]
    if candidate_dir == data or data in candidate_dir.parents:
        raise ValueError("Candidate must be outside the source .Data directory")
    signature = (sha256_file(source), sha256_file(sdb_path_for_enl(source)))
    tree = tree_hashes(data)
    candidate_dir.mkdir()
    candidate = candidate_dir / f"{source.stem[:35]}-CWYW-{uuid4().hex[:10]}.enl"
    shutil.copy2(source, candidate)
    shutil.copytree(data, candidate.with_suffix(".Data"))
    require_closed_endnote()
    for enl in (source, candidate):
        if signature != (sha256_file(enl), sha256_file(sdb_path_for_enl(enl))):
            raise RuntimeError(f"Library changed during copy: {enl}")
    if tree_hashes(candidate.with_suffix(".Data")) != tree:
        raise RuntimeError("Candidate copy differs from source library files")
    return candidate, signature


def _truncate_abstract(value: object, limit: int = 3000) -> str:
    cleaned = " ".join(str(value or "").split())
    if len(cleaned) <= limit:
        return cleaned
    window = cleaned[: limit + 1]
    for sep in (". ", "! ", "? ", "。"):
        pos = window.rfind(sep)
        if pos >= int(limit * 0.6):
            return window[: pos + len(sep.rstrip())].rstrip()
    pos = window.rfind(" ")
    if pos >= int(limit * 0.5):
        return window[:pos].rstrip()
    return cleaned[:limit].rstrip()


def format_enw_entry(item: dict[str, object]) -> str:
    item = {
        key: " ".join(value.split()) if isinstance(value, str) and key != "pdf_path" else value
        for key, value in item.items()
    }
    lines = ["%0 Journal Article"]
    if item.get("title"):
        lines.append(f"%T {' '.join(str(item['title']).split())}")
    for author in item.get("authors") or []:
        if isinstance(author, dict):
            fam = str(author.get("family", "")).strip()
            giv = str(author.get("given", "")).strip()
            name = f"{fam}, {giv}" if fam and giv else fam
        else:
            name = str(author).strip()
        if name:
            lines.append(f"%A {' '.join(name.split())}")
    if item.get("journal"):
        lines.append(f"%J {str(item['journal']).strip()}")
    if item.get("year"):
        lines.append(f"%D {str(item['year']).strip()}")
    if item.get("volume"):
        lines.append(f"%V {str(item['volume']).strip()}")
    if item.get("issue"):
        lines.append(f"%N {str(item['issue']).strip()}")
    if item.get("pages"):
        lines.append(f"%P {str(item['pages']).strip()}")
    if item.get("doi"):
        lines.append(f"%R {str(item['doi']).strip()}")
    if item.get("pmid"):
        lines.append(f"%M PMID: {str(item['pmid']).strip()}")
    if item.get("abstract"):
        clean_abs = _truncate_abstract(item["abstract"])
        if clean_abs:
            lines.append(f"%X {clean_abs}")
    if item.get("pdf_path"):
        lines.append(f"%> {Path(str(item['pdf_path'])).resolve().as_uri()}")
    return "\r\n".join(lines) + "\r\n\r\n"


def format_ris_entry(item: dict[str, object]) -> str:
    lines = ["TY  - JOUR"]
    if item.get("title"):
        lines.append(f"TI  - {' '.join(str(item['title']).split())}")
    for author in item.get("authors") or []:
        if isinstance(author, dict):
            fam = str(author.get("family", "")).strip()
            giv = str(author.get("given", "")).strip()
            name = f"{fam}, {giv}" if fam and giv else fam
        else:
            name = str(author).strip()
        if name:
            lines.append(f"AU  - {name}")
    if item.get("journal"):
        lines.append(f"JO  - {str(item['journal']).strip()}")
    if item.get("year"):
        lines.append(f"PY  - {str(item['year']).strip()}")
    if item.get("volume"):
        lines.append(f"VL  - {str(item['volume']).strip()}")
    if item.get("issue"):
        lines.append(f"IS  - {str(item['issue']).strip()}")
    pages = str(item.get("pages", "")).strip()
    if pages:
        if "-" in pages:
            sp, ep = pages.split("-", 1)
            lines.append(f"SP  - {sp.strip()}")
            lines.append(f"EP  - {ep.strip()}")
        else:
            lines.append(f"SP  - {pages}")
    if item.get("doi"):
        lines.append(f"DO  - {str(item['doi']).strip()}")
    if item.get("abstract"):
        clean_abs = _truncate_abstract(item["abstract"])
        if clean_abs:
            lines.append(f"N2  - {clean_abs}")
    lines.append("ER  - ")
    return "\r\n".join(lines) + "\r\n"


def _send(hwnd, msg, wparam=0, lparam=0):
    user = ctypes.WinDLL("user32", use_last_error=True)
    fn = user.SendMessageTimeoutW
    fn.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
                   wintypes.UINT, wintypes.UINT, ctypes.POINTER(ctypes.c_size_t))
    fn.restype = wintypes.LPARAM
    result = ctypes.c_size_t()
    if not fn(hwnd, msg, wparam, lparam, 2, 3000, ctypes.byref(result)):
        raise WindowInterference("EndNote did not respond to a targeted control message")
    return ctypes.c_ssize_t(result.value).value


def _read_text(hwnd):
    n = _send(hwnd, 0x000E)  # WM_GETTEXTLENGTH
    if not 0 <= n <= 32767:
        raise WindowInterference("Unexpected native text control")
    buf = ctypes.create_unicode_buffer(n + 1)
    _send(hwnd, 0x000D, n + 1, ctypes.addressof(buf))
    return buf.value


def _children(parent):
    import win32gui
    controls = []

    def _collect(hwnd, _):
        controls.append(hwnd)
        return True

    try:
        win32gui.EnumChildWindows(parent, _collect, None)
    except win32gui.error:
        pass
    return controls


def _control(dialog, cid, cls):
    import win32gui
    found = []
    for h in _children(dialog):
        try:
            if (
                win32gui.GetDlgCtrlID(h) == cid
                and win32gui.GetClassName(h).casefold() == cls.casefold()
                and win32gui.IsWindowVisible(h)
            ):
                found.append(h)
        except win32gui.error:
            continue
    if len(found) != 1:
        raise WindowInterference(f"Cannot identify expected {cls} control {cid} in the target-owned dialog")
    return found[0]


def _filename_edit(picker):
    import win32gui
    controls = _children(picker)
    edits = [h for h in controls if win32gui.GetClassName(h) == "Edit" and win32gui.IsWindowVisible(h)]
    holders = [h for h in controls if win32gui.GetDlgCtrlID(h) == 1148]
    matches = [h for h in edits if win32gui.GetDlgCtrlID(h) in (1148, 1152)
               or any(win32gui.IsChild(owner, h) for owner in holders)]
    if len(matches) != 1:
        raise WindowInterference("Cannot identify the native file-name edit box")
    return matches[0]


def _select_import_option(target, dialog):
    combo = _control(dialog, 1124, "ComboBox")
    name = ctypes.create_unicode_buffer("EndNote Import")
    index = _send(combo, CB_FINDSTRINGEXACT, -1, ctypes.addressof(name))
    if index < 0:
        raise WindowInterference("EndNote Import filter is unavailable in the target dialog")
    target.require_exclusive((dialog,))
    if _send(combo, CB_SETCURSEL, index) != index:
        raise WindowInterference("Could not select EndNote Import")
    for notice in (CBN_SELCHANGE, CBN_SELENDOK):
        target.require_exclusive((dialog,))
        _send(dialog, WM_COMMAND, 1124 | (notice << 16), combo)
    if _send(combo, CB_GETCURSEL) != index:
        raise WindowInterference("The EndNote Import selection changed")
    length = _send(combo, CB_GETLBTEXTLEN, index)
    if not 0 < length < 256:
        raise WindowInterference("Cannot read the selected import filter")
    label = ctypes.create_unicode_buffer(length + 1)
    _send(combo, CB_GETLBTEXT, index, ctypes.addressof(label))
    if label.value.casefold() != "endnote import":
        raise WindowInterference("The selected filter is not EndNote Import")


def _menu_has_import(target):
    import win32api
    import win32gui
    def contains(menu):
        for i in range(win32gui.GetMenuItemCount(menu)):
            sub = win32gui.GetSubMenu(menu, i)
            if sub and contains(sub):
                return True
            if not sub and win32gui.GetMenuItemID(menu, i) == 32777:
                return not (win32gui.GetMenuState(menu, i, 0x400) & 3)
        return False
    menu = win32gui.GetMenu(target.hwnd)
    if menu:
        return contains(menu)
    # EndNote's custom MFC menu bar can draw the same commands without a system HMENU.
    dll = win32api.LoadLibraryEx(str(target.proc.args[0]), 0, 2)  # datafile only
    try:
        resource_menu = win32gui.LoadMenu(dll, 1)
        try:
            return contains(resource_menu)
        finally:
            win32gui.DestroyMenu(resource_menu)
    finally:
        win32api.FreeLibrary(dll)


def _wait_selected_import(target, dialog, enw, timeout):
    """The parent import form can show its selection after the picker disappears."""
    expected = {str(enw).casefold(), enw.name.casefold()}
    deadline = time.monotonic() + min(timeout, 5)
    while True:
        target.require_exclusive((dialog,))
        selected = _read_text(_control(dialog, 1100, "RichEdit20W")).strip()
        if selected.casefold() in expected:
            return
        if selected:
            raise WindowInterference("The target dialog displays a different import file")
        if time.monotonic() >= deadline:
            raise WindowInterference("The target import form never displayed the selected file")
        time.sleep(.1)


def _targeted_import(target, enw, timeout):
    import win32gui
    target.require_exclusive()
    if not _menu_has_import(target):
        raise WindowInterference("The bound window does not offer EndNote 21 File > Import > File")
    # Native EndNote 21 command 32777 opens resource 135, the Import File dialog.
    win32gui.PostMessage(target.hwnd, WM_COMMAND, 32777, 0)
    dialog = target.wait_dialog((), lambda _, title: title.casefold() == "import file", timeout)
    choose = _control(dialog, 1097, "Button")
    commit = _control(dialog, 1, "Button")
    if _read_text(commit).replace("&", "").casefold() != "import":
        raise WindowInterference("Unexpected confirmation button in the Import File dialog")
    target.require_exclusive((dialog,))
    win32gui.PostMessage(choose, BM_CLICK, 0, 0)
    picker = target.wait_dialog((dialog,), lambda _, title: bool(title), timeout)
    edit = _filename_edit(picker)
    confirm = _control(picker, 1, "Button")
    path = ctypes.create_unicode_buffer(str(enw))
    target.require_exclusive((dialog, picker))
    _send(edit, 0x000C, 0, ctypes.addressof(path))  # WM_SETTEXT
    if _read_text(edit) != str(enw):
        raise WindowInterference("The native file picker did not retain the exact import path")
    target.require_exclusive((dialog, picker))
    win32gui.PostMessage(confirm, BM_CLICK, 0, 0)
    target.wait_dismissed((dialog,), picker, timeout)
    _wait_selected_import(target, dialog, enw, timeout)
    _select_import_option(target, dialog)
    target.require_exclusive((dialog,))
    if not win32gui.IsWindowEnabled(commit):
        raise WindowInterference("Native Import button is disabled")
    win32gui.PostMessage(commit, BM_CLICK, 0, 0)
    target.wait_dismissed((), dialog, timeout, progress=_is_import_progress)


def _is_import_progress(hwnd, title):
    """Identify the read-only EndNote 21 progress stage, after the owned Import click."""
    import win32gui
    if title != "Importing...":
        return False
    visible = [(win32gui.GetClassName(c), win32gui.GetDlgCtrlID(c), win32gui.GetWindowText(c))
               for c in _children(hwnd) if win32gui.IsWindowVisible(c)]
    return sorted(visible) == [("Button", 2, "Cancel"), ("msctls_progress32", 3000, "")]


def run_endnote_import(enl_path: Path, enw_path: Path, timeout: float = 40.0) -> dict[str, object]:
    """Import through the candidate's own library window and owned file picker."""
    enl_path = Path(enl_path).resolve(strict=True)
    enw_path = Path(enw_path).resolve(strict=True)
    if enw_path.suffix.casefold() != ".enw" or not enw_path.is_file():
        raise ValueError("Expected an existing EndNote tagged import file")
    require_closed_endnote()
    proc = subprocess.Popen([str(resolve_endnote_exe()), str(enl_path)])
    target = TargetWindow.bind(proc, enl_path, timeout)
    try:
        _targeted_import(target, enw_path, timeout)
    except WindowInterference:
        raise  # Preserve the UI when the target or a modal is no longer certain.
    except Exception:
        try:
            target.close()
        except (OSError, RuntimeError):
            pass  # An unknown window or modal keeps the candidate open for local inspection.
        raise
    closed = target.close()
    return {"main_pid": proc.pid, "target_hwnd": target.hwnd, "method": "target_library_import_dialog",
            "cleanly_closed": closed}


def plan_import(items, records, preferred=()):
    catalog_ids = [item["pdf_id"] for item in items]
    if len(set(catalog_ids)) != len(catalog_ids):
        raise ValueError("Repeated selection identifier")
    record_index = RecordIndex(records)
    resolved, pending, aliases, members = {}, {}, {}, {}
    def add_member(group, item):
        identity = Identity.from_record(item)
        group_members = members.setdefault(group, [])
        if any(compare(identity, member) not in {"doi", "pmid", "title_author_year", "title_variant"}
               for member in group_members):
            raise IdentityReview(f"{item['pdf_id']}: not equivalent to every member of proposed group {group[1]}")
        group_members.append(identity)

    # Classify identified records before title-only variants. A title-only row
    # must never bridge different DOI/PMID records because it happened to arrive first.
    def identity_priority(item):
        identity = Identity.from_record(item)
        return (-(bool(identity.doi) + bool(identity.pmid)), not bool(item.get("pdf_path")))

    for item in sorted(items, key=identity_priority):
        pdf_id = item["pdf_id"]
        if item.get("identity_review") or item.get("is_retracted"):
            raise IdentityReview(f"{pdf_id}: {item.get('identity_review') or 'retraction flagged'}")
        rid = record_index.match(item, preferred)
        if rid is not None:
            add_member(("existing", rid), item)
            resolved[pdf_id] = rid
            continue
        if not (
            item.get("title") and item.get("authors")
            and item.get("first_author_family")
            and re.fullmatch(r"(?:18|19|20)\d{2}", str(item.get("year", "")))
        ):
            raise IdentityReview(f"{pdf_id}: missing title, author or 4-digit year")
        alias = RecordIndex(pending).match(item)
        if alias is None:
            pending[pdf_id] = dict(item)
            add_member(("new", pdf_id), item)
            aliases[pdf_id] = pdf_id
        else:
            add_member(("new", alias), item)
            if item.get("pdf_path") and not pending[alias].get("pdf_path"):
                pending[alias].update(pdf_path=item["pdf_path"], pdf_sha256=item["pdf_sha256"])
            aliases[pdf_id] = alias
    return resolved, pending, aliases


def sync_candidate(source, candidate_dir, items, work_dir, preferred=(), expected_signature=None):
    require_closed_endnote()
    source = source.resolve(strict=True)
    before = copy_sdb_snapshot(source, work_dir / "before.eni")
    old_pdf_hashes = tree_hashes(source.with_suffix(".Data") / "PDF")
    library_id = read_library_id(before)
    records = read_endnote_active_records(before)
    resolved, pending, aliases = plan_import(items, records, preferred)
    signature = (sha256_file(source), sha256_file(before))
    if expected_signature is not None and expected_signature != signature:
        raise RuntimeError("EndNote source changed during citation planning")
    candidate, copied_signature = create_library_candidate(source, candidate_dir)
    if signature != copied_signature:
        raise RuntimeError("EndNote source changed between snapshot and copy")
    import_info = None
    new_ids = set()
    after_records = records
    imported = {}
    if pending:
        enw = work_dir / "delta_import.enw"
        with enw.open("xb") as handle:
            handle.write("".join(format_enw_entry(item) for item in pending.values()).encode("utf-8"))
        import_info = run_endnote_import(candidate, enw)
        if not import_info["cleanly_closed"]:
            raise RuntimeError("EndNote import process did not close cleanly; candidate not accepted")
        require_closed_endnote()
        after = copy_sdb_snapshot(candidate, work_dir / "after.eni")
        after_records = read_endnote_active_records(after)
        if read_library_id(after) != library_id or any(after_records.get(rid) != record for rid, record in records.items()):
            raise RuntimeError("Import altered the candidate library identity or pre-existing records")
        new_ids = set(after_records) - set(records)
        if len(new_ids) != len(pending):
            raise RuntimeError(f"Expected {len(pending)} imported works, observed {len(new_ids)}")
        added_index = RecordIndex({rid: after_records[rid] for rid in new_ids})
        canonical = {key: added_index.match(item) for key, item in pending.items()}
        if set(canonical.values()) != new_ids or len(set(canonical.values())) != len(pending):
            raise RuntimeError("Imported IDs do not uniquely identify the proposed new works")
        resolved.update({key: canonical[alias] for key, alias in aliases.items()})
        imported = {canonical[key]: item for key, item in pending.items()}
    verify_library_contents(before, after if pending else before, candidate, imported, old_pdf_hashes)

    rows = {}
    for key, rid in resolved.items():
        record = after_records[rid]
        family, year = record["first_author_family"], record["year"]
        if not family or not re.fullmatch(r"(?:18|19|20)\d{2}", year) or re.search(r"[{};#\r\n]", family):
            raise IdentityReview(f"{key}: EndNote record #{rid} has unsafe temporary-citation metadata")
        rows[key] = {
            "pdf_id": key, "record_id": rid, "citation_token": f"{family}, {year} #{rid}",
            "first_author_family": family, "year": year, "title": record["title"],
            "doi": record["doi"], "sync_status": "newly_imported" if rid in new_ids else "reused_existing",
        }
    if signature != (sha256_file(source), sha256_file(sdb_path_for_enl(source))):
        raise RuntimeError("Source library changed during candidate import")
    return {
        "status": "PASS",
        "enl_path": str(candidate),
        "library_id": library_id,
        "source_signature": list(signature),
        "before_active_count": len(records),
        "after_active_count": len(after_records),
        "new_record_ids": sorted(new_ids),
        "imported_pdf_ids": {rid: item["pdf_id"] for rid, item in imported.items() if item.get("pdf_path")},
        "import_info": import_info,
        "by_pdf_id": rows,
        "citations": [rows[item["pdf_id"]] for item in items],
    }


def selected_pdfs(catalog_path, selected_ids, work):
    items = json.loads(catalog_path.read_text(encoding="utf-8"))["items"]
    by_id = {item["pdf_id"]: item for item in items}
    if len(by_id) != len(items):
        raise ValueError("Repeated PDF ID in catalog")
    selected = []
    for n, key in enumerate(dict.fromkeys(selected_ids), 1):
        item = by_id[key]
        path = Path(item["pdf_path"]).resolve(strict=True)
        if path.suffix.casefold() != ".pdf" or not path.is_file():
            raise ValueError("Selected catalog entries must refer to PDF files")
        signature = path.stat()
        folder = check_work_path(work) / "pdfs" / str(n)
        folder.mkdir(parents=True, exist_ok=False)
        frozen = folder / path.name
        with path.open("rb") as original, frozen.open("xb") as copy:
            shutil.copyfileobj(original, copy)
        if item.get("pdf_sha256") != sha256_file(frozen):
            raise RuntimeError(f"PDF changed or catalog lacks a content hash: {path}")
        if path.stat().st_mtime_ns != signature.st_mtime_ns:
            raise RuntimeError(f"PDF changed while validating: {path}")
        selected.append({**item, "pdf_path": str(frozen)})
    return selected


def main():
    parser = argparse.ArgumentParser(description="Import selected works only into a new EndNote library copy")
    parser.add_argument("--enl", required=True, type=Path, help="Read-only source library")
    parser.add_argument("--candidate-dir", required=True, type=Path, help="New directory for the full library copy")
    parser.add_argument("--catalog", required=True, type=Path)
    parser.add_argument("--select-pdf-ids", required=True)
    parser.add_argument("--work-dir", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    args.work_dir = check_work_path(args.work_dir)
    args.candidate_dir = check_work_path(args.candidate_dir)
    args.output = check_work_path(args.output, ".json")
    ids = [item.strip() for item in args.select_pdf_ids.split(",") if item.strip()]
    if not ids or args.output.exists():
        parser.error("Select at least one PDF ID and use a new output path")
    if args.work_dir.resolve() == args.candidate_dir.resolve():
        parser.error("Use separate work and candidate directories")
    args.work_dir.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="sync-", dir=args.work_dir))
    result = sync_candidate(args.enl, args.candidate_dir, selected_pdfs(args.catalog, ids, work), work)
    tmp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=args.output.parent, prefix=f".{args.output.name}.", suffix=".tmp", delete=False
        ) as handle:
            tmp_path = handle.name
            json.dump(result, handle, ensure_ascii=False, indent=2)
        if os.name == "nt":
            os.rename(tmp_path, args.output)
        else:
            os.link(tmp_path, args.output)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
