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
import tempfile
import time
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
from reference_identity import IdentityReview, RecordIndex


WM_CLOSE = 0x0010
BM_CLICK = 0x00F5
CB_GETCOUNT = 0x0146
CB_GETLBTEXT = 0x0148
CB_GETLBTEXTLEN = 0x0149
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
        clean_abs = " ".join(str(item["abstract"]).split())[:3000]
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
        clean_abs = " ".join(str(item["abstract"]).split())[:3000]
        if clean_abs:
            lines.append(f"N2  - {clean_abs}")
    lines.append("ER  - ")
    return "\r\n".join(lines) + "\r\n"


def _win32():
    import win32con
    import win32gui
    import win32process

    return win32gui, win32process, win32con


def _visible_windows_for_pids(pids: set[int]) -> list[dict[str, object]]:
    win32gui, win32process, win32con = _win32()
    windows: list[int] = []
    win32gui.EnumWindows(lambda hwnd, _: windows.append(hwnd), None)
    found: list[dict[str, object]] = []
    for hwnd in windows:
        if not win32gui.IsWindowVisible(hwnd):
            continue
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid not in pids:
            continue
        found.append(
            {
                "hwnd": int(hwnd),
                "pid": int(pid),
                "class": win32gui.GetClassName(hwnd) or "",
                "title": (win32gui.GetWindowText(hwnd) or "").strip(),
            }
        )
    return found


def _inspect_dialog_tree(hwnd: int) -> list[dict[str, object]]:
    win32gui, _, win32con = _win32()
    controls: list[dict[str, object]] = []

    def walk(parent: int) -> None:
        child = win32gui.GetWindow(parent, win32con.GW_CHILD)
        while child:
            controls.append(
                {
                    "hwnd": int(child),
                    "ctrl_id": int(win32gui.GetDlgCtrlID(child)),
                    "class": win32gui.GetClassName(child) or "",
                    "text": (win32gui.GetWindowText(child) or "").strip(),
                }
            )
            walk(child)
            child = win32gui.GetWindow(child, win32con.GW_HWNDNEXT)

    walk(hwnd)
    return controls


def _try_select_endnote_import_filter(hwnd: int) -> dict[str, object] | None:
    """If an EndNote Import Filter modal appears, select 'EndNote Import' and click OK."""
    win32gui, _, _ = _win32()
    controls = _inspect_dialog_tree(hwnd)
    combo = next((c for c in controls if c["class"].casefold() == "combobox"), None)
    ok_btn = next(
        (
            c
            for c in controls
            if c["class"].casefold() == "button"
            and (c["ctrl_id"] == 1 or c["text"] in {"OK", "确定", "&OK", "Import", "&Import"})
        ),
        None,
    )
    if combo is None or ok_btn is None:
        return None
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.SendMessageW.restype = ctypes.c_ssize_t
    combo_hwnd = int(combo["hwnd"])
    count = int(user32.SendMessageW(combo_hwnd, CB_GETCOUNT, 0, 0))
    items: list[str] = []
    for idx in range(max(count, 0)):
        length = int(user32.SendMessageW(combo_hwnd, CB_GETLBTEXTLEN, idx, 0))
        buf = ctypes.create_unicode_buffer(max(length + 1, 1))
        user32.SendMessageW(combo_hwnd, CB_GETLBTEXT, idx, ctypes.addressof(buf))
        items.append(buf.value)
    match_idx = next((i for i, val in enumerate(items) if val.casefold() == "endnote import"), None)
    if match_idx is None:
        return None
    user32.SendMessageW(combo_hwnd, CB_SETCURSEL, match_idx, 0)
    parent = win32gui.GetParent(combo_hwnd)
    ctrl_id = int(combo["ctrl_id"])
    for note in (CBN_SELCHANGE, CBN_SELENDOK):
        win32gui.SendMessage(parent, WM_COMMAND, ctrl_id | (note << 16), combo_hwnd)
    win32gui.SendMessage(int(ok_btn["hwnd"]), BM_CLICK, 0, 0)
    return {"handled": True, "selected_filter": items[match_idx]}


def run_endnote_import(enl_path: Path, enw_path: Path, timeout: float = 40.0) -> dict[str, object]:
    """Open EndNote on target .enl, import .enw, and close the spawned instance cleanly."""
    win32gui, _, _ = _win32()
    endnote_exe = resolve_endnote_exe()
    sdb_file = sdb_path_for_enl(enl_path)
    wal_file = sdb_file.with_name("sdb.eni-wal")

    require_closed_endnote()
    proc = subprocess.Popen([str(endnote_exe), str(enl_path)])
    open_deadline = time.monotonic() + timeout
    main_titles: list[str] = []
    while time.monotonic() < open_deadline:
        if proc.poll() is not None:
            raise RuntimeError(f"EndNote.exe exited prematurely (rc={proc.returncode}) while opening {enl_path.name}")
        wins = _visible_windows_for_pids({proc.pid})
        main_titles = [str(w["title"]) for w in wins if w["title"]]
        if any(
            t.casefold() in {enl_path.name.casefold(), f"endnote 21 - {enl_path.name}".casefold()}
            for t in main_titles
        ):
            break
        time.sleep(0.25)
    else:
        for win in wins:
            if win["title"]:
                win32gui.PostMessage(int(win["hwnd"]), WM_CLOSE, 0, 0)
        raise RuntimeError(f"Timed out waiting for EndNote window for {enl_path.name}; titles={main_titles}")

    time.sleep(2.5)
    before_sig = (
        sdb_file.stat().st_size,
        sdb_file.stat().st_mtime_ns,
        wal_file.stat().st_size if wal_file.is_file() else 0,
        wal_file.stat().st_mtime_ns if wal_file.is_file() else 0,
    )

    expected_titles = {enl_path.name.casefold(), f"endnote 21 - {enl_path.name}".casefold()}
    active_titles = [w["title"].casefold() for w in _visible_windows_for_pids({proc.pid}) if w["title"]]
    if set(list_process_pids("EndNote.exe")) != {proc.pid} or not (set(active_titles) & expected_titles):
        for win in _visible_windows_for_pids({proc.pid}):
            if win["title"].casefold() in expected_titles:
                win32gui.PostMessage(int(win["hwnd"]), WM_CLOSE, 0, 0)
        raise RuntimeError("EndNote candidate is no longer the sole verified import target")
    imp = subprocess.Popen([str(endnote_exe), str(enw_path)])
    import_deadline = time.monotonic() + 30.0
    activity_seen = False
    modal_diagnostics: list[dict[str, object]] = []
    handled_dialogs: list[dict[str, object]] = []
    last_activity_at = time.monotonic()

    while time.monotonic() < import_deadline:
        all_pids = {proc.pid, imp.pid}
        wins = _visible_windows_for_pids(all_pids)
        dialogs = [w for w in wins if w["class"] == "#32770"]
        for d in dialogs:
            title_low = str(d["title"]).casefold()
            if "import" in title_low or "filter" in title_low:
                handled = _try_select_endnote_import_filter(int(d["hwnd"]))
                if handled:
                    handled_dialogs.append({"window": d, **handled})
                    continue
            controls = _inspect_dialog_tree(int(d["hwnd"]))
            modal_diagnostics.append({"window": d, "controls": controls})

        cur_sig = (
            sdb_file.stat().st_size,
            sdb_file.stat().st_mtime_ns,
            wal_file.stat().st_size if wal_file.is_file() else 0,
            wal_file.stat().st_mtime_ns if wal_file.is_file() else 0,
        )
        if cur_sig != before_sig:
            if not activity_seen:
                activity_seen = True
            last_activity_at = time.monotonic()
            before_sig = cur_sig

        imp_done = imp.poll() is not None
        elapsed_quiet = time.monotonic() - last_activity_at
        if imp_done and not dialogs and ((activity_seen and elapsed_quiet >= 2.5) or elapsed_quiet >= 4.5):
            break
        time.sleep(0.25)

    if imp.poll() is None:
        try:
            imp.wait(timeout=5.0)
        except subprocess.TimeoutExpired:
            pass

    # Cleanly close only our spawned EndNote instance
    close_wins = _visible_windows_for_pids({proc.pid})
    for w in close_wins:
        if w["class"] != "#32770" and w["title"]:
            win32gui.PostMessage(int(w["hwnd"]), WM_CLOSE, 0, 0)

    close_deadline = time.monotonic() + 20.0
    cleanly_closed = False
    while time.monotonic() < close_deadline:
        if proc.poll() is not None:
            cleanly_closed = True
            break
        time.sleep(0.25)

    return {
        "endnote_exe": str(endnote_exe),
        "main_pid": proc.pid,
        "import_pid": imp.pid,
        "import_returncode": imp.poll(),
        "main_titles": main_titles,
        "db_activity_seen": activity_seen,
        "handled_dialogs": handled_dialogs,
        "unexpected_dialogs": modal_diagnostics,
        "cleanly_closed": cleanly_closed,
    }


def plan_import(items, records, preferred=()):
    catalog_ids = [item["pdf_id"] for item in items]
    if len(set(catalog_ids)) != len(catalog_ids):
        raise ValueError("Repeated selection identifier")
    record_index = RecordIndex(records)
    resolved, pending, aliases = {}, {}, {}
    for item in items:
        pdf_id = item["pdf_id"]
        if item.get("identity_review") or item.get("is_retracted"):
            raise IdentityReview(f"{pdf_id}: {item.get('identity_review') or 'retraction flagged'}")
        rid = record_index.match(item, preferred)
        if rid is not None:
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
            pending[pdf_id] = item
            aliases[pdf_id] = pdf_id
        else:
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


def selected_pdfs(catalog_path, selected_ids):
    items = json.loads(catalog_path.read_text(encoding="utf-8"))["items"]
    by_id = {item["pdf_id"]: item for item in items}
    if len(by_id) != len(items):
        raise ValueError("Repeated PDF ID in catalog")
    selected = []
    for key in dict.fromkeys(selected_ids):
        item = by_id[key]
        path = Path(item["pdf_path"])
        signature = path.stat()
        if item.get("pdf_sha256") != sha256_file(path):
            raise RuntimeError(f"PDF changed or catalog lacks a content hash: {path}")
        if path.stat().st_mtime_ns != signature.st_mtime_ns:
            raise RuntimeError(f"PDF changed while validating: {path}")
        selected.append(item)
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
    result = sync_candidate(args.enl, args.candidate_dir, selected_pdfs(args.catalog, ids), work)
    with args.output.open("x", encoding="utf-8") as handle:
        json.dump(result, handle, ensure_ascii=False, indent=2)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
