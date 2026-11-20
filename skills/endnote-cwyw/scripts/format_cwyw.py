# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
# ]
# ///

from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import hashlib
import json
import os
import re
import subprocess
import sys
import threading
import time
import winreg
from pathlib import Path

from audit_cwyw_docx import (
    BIBLIOGRAPHY_HEADINGS,
    TEMPORARY_CITATION_RE,
    audit,
    parse_expected_ids,
)


CB_GETCOUNT = 0x0146
CB_GETCURSEL = 0x0147
CB_GETLBTEXT = 0x0148
CB_GETLBTEXTLEN = 0x0149
CB_SETCURSEL = 0x014E
CBN_SELCHANGE = 1
CBN_SELENDOK = 9


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def temporary_state(path: Path) -> dict[str, object]:
    import zipfile
    from xml.etree import ElementTree as ET

    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    texts: list[str] = []
    with zipfile.ZipFile(path, "r") as archive:
        names = set(archive.namelist())
        for member in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
            if member not in names:
                continue
            root = ET.fromstring(archive.read(member))
            texts.append("".join(node.text or "" for node in root.iter(f"{{{w_ns}}}t")))
    text = "\n".join(texts)
    citations = TEMPORARY_CITATION_RE.findall(text)
    ids = sorted({int(value) for citation in citations for value in re.findall(r"#(\d+)", citation)})
    return {"citation_groups": len(citations), "record_ids": ids, "citations": citations}


def resolve_endnote_exe() -> Path:
    reg_queries = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe"),
        (winreg.HKEY_CLASSES_ROOT, r"EndNotePlus.ENW\shell\open\command"),
    ]
    for hive, subkey in reg_queries:
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


def list_process_pids(image_names: set[str]) -> list[tuple[int, str]]:
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

    TH32CS_SNAPPROCESS = 0x00000002
    INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateToolhelp32Snapshot.argtypes = (wintypes.DWORD, wintypes.DWORD)
    kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    kernel32.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    kernel32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

    snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snap in (None, INVALID_HANDLE_VALUE):
        raise ctypes.WinError(ctypes.get_last_error())
    targets = {name.casefold() for name in image_names}
    found: list[tuple[int, str]] = []
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        ok = kernel32.Process32FirstW(snap, ctypes.byref(entry))
        if not ok:
            raise ctypes.WinError(ctypes.get_last_error())
        while ok:
            exe = entry.szExeFile
            if exe.casefold() in targets:
                found.append((int(entry.th32ProcessID), exe))
            ok = kernel32.Process32NextW(snap, ctypes.byref(entry))
        if ctypes.get_last_error() != 18:
            raise ctypes.WinError(ctypes.get_last_error())
    finally:
        kernel32.CloseHandle(snap)
    return found


def visible_window_titles(pids: set[int], win32gui: object, win32process: object) -> list[tuple[int, int, str]]:
    if not pids:
        return []
    windows: list[int] = []
    win32gui.EnumWindows(lambda hwnd, _: windows.append(hwnd), None)
    results: list[tuple[int, int, str]] = []
    for hwnd in windows:
        if not win32gui.IsWindowVisible(hwnd):
            continue
        _, candidate_pid = win32process.GetWindowThreadProcessId(hwnd)
        if candidate_pid in pids:
            title = win32gui.GetWindowText(hwnd) or ""
            if title.strip():
                results.append((int(hwnd), int(candidate_pid), title))
    return results


def ensure_endnote_library_open(
    enl_path: Path,
    win32gui: object,
    win32process: object,
    timeout: float = 35.0,
) -> tuple[dict[str, object], subprocess.Popen[bytes] | None]:
    resolved_enl = enl_path.resolve()
    if not resolved_enl.is_file():
        raise FileNotFoundError(resolved_enl)
    existing = list_process_pids({"EndNote.exe"})
    if existing:
        raise RuntimeError(f"Close EndNote before opening the explicit library {resolved_enl.name}: {existing}")

    endnote_exe = resolve_endnote_exe()
    proc = subprocess.Popen([str(endnote_exe), str(resolved_enl)])
    deadline = time.monotonic() + timeout
    titles: list[str] = []
    while time.monotonic() < deadline:
        if proc.poll() is not None:
            raise RuntimeError(f"EndNote.exe exited prematurely with code {proc.returncode}")
        win_info = visible_window_titles({proc.pid}, win32gui, win32process)
        titles = [title for _, _, title in win_info]
        good = [
            title
            for title in titles
            if title.casefold() in {
                resolved_enl.name.casefold(),
                f"endnote 21 - {resolved_enl.name}".casefold(),
            }
        ]
        if good:
            time.sleep(2.0)
            return (
                {
                    "launched_by_us": True,
                    "pid": proc.pid,
                    "pids": [proc.pid],
                    "titles": titles,
                    "enl": str(resolved_enl),
                    "endnote_exe": str(endnote_exe),
                },
                proc,
            )
        time.sleep(0.25)
    close_spawned_endnote({"launched_by_us": True, "pid": proc.pid}, proc, win32gui, win32process, timeout=10.0)
    raise RuntimeError(
        f"EndNote did not show target library window for {resolved_enl.name} within {timeout}s; seen titles={titles}"
    )


def close_spawned_endnote(
    info: dict[str, object] | None,
    proc: subprocess.Popen[bytes] | None,
    win32gui: object,
    win32process: object,
    timeout: float = 15.0,
) -> dict[str, object] | None:
    if not info or not info.get("launched_by_us") or proc is None or proc.pid != info.get("pid"):
        return None
    pid = int(info["pid"])
    if proc.poll() is not None:
        return {"pid": pid, "wm_close_posted": 0, "cleanly_exited": True}
    WM_CLOSE = 0x0010
    windows: list[int] = []
    win32gui.EnumWindows(lambda hwnd, _: windows.append(hwnd), None)
    posted = 0
    for hwnd in windows:
        if not win32gui.IsWindowVisible(hwnd):
            continue
        _, candidate_pid = win32process.GetWindowThreadProcessId(hwnd)
        if candidate_pid == pid and win32gui.GetWindowText(hwnd):
            win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)
            posted += 1
    cleanly_exited = False
    if proc is not None:
        try:
            proc.wait(timeout=timeout)
            cleanly_exited = True
        except subprocess.TimeoutExpired:
            cleanly_exited = False
    else:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if not any(p == pid for p, _ in list_process_pids({"EndNote.exe"})):
                cleanly_exited = True
                break
            time.sleep(0.25)
    return {"pid": pid, "wm_close_posted": posted, "cleanly_exited": cleanly_exited}


def direct_children(hwnd: int, win32gui: object, win32con: object) -> list[int]:
    children: list[int] = []
    child = win32gui.GetWindow(hwnd, win32con.GW_CHILD)
    while child:
        children.append(child)
        child = win32gui.GetWindow(child, win32con.GW_HWNDNEXT)
    return children


def find_control(hwnd: int, control_id: int, win32gui: object, win32con: object) -> int | None:
    for child in direct_children(hwnd, win32gui, win32con):
        if win32gui.GetDlgCtrlID(child) == control_id:
            return child
        nested = find_control(child, control_id, win32gui, win32con)
        if nested is not None:
            return nested
    return None


def combo_items(hwnd: int, win32gui: object) -> dict[str, object]:
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.SendMessageW.argtypes = [
        wintypes.HWND,
        wintypes.UINT,
        wintypes.WPARAM,
        wintypes.LPARAM,
    ]
    user32.SendMessageW.restype = ctypes.c_ssize_t
    count = int(user32.SendMessageW(hwnd, CB_GETCOUNT, 0, 0))
    current = int(user32.SendMessageW(hwnd, CB_GETCURSEL, 0, 0))
    items: list[str] = []
    for index in range(max(count, 0)):
        length = int(user32.SendMessageW(hwnd, CB_GETLBTEXTLEN, index, 0))
        buffer = ctypes.create_unicode_buffer(max(length + 1, 1))
        user32.SendMessageW(hwnd, CB_GETLBTEXT, index, ctypes.addressof(buffer))
        items.append(buffer.value)
    return {"count": count, "current_index": current, "items": items}


def select_combo(hwnd: int, index: int) -> int:
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.SendMessageW.argtypes = [
        wintypes.HWND,
        wintypes.UINT,
        wintypes.WPARAM,
        wintypes.LPARAM,
    ]
    user32.SendMessageW.restype = ctypes.c_ssize_t
    return int(user32.SendMessageW(hwnd, CB_SETCURSEL, index, 0))


def find_control_by_class(
    hwnd: int, class_name: str, win32gui: object, win32con: object
) -> int | None:
    target = class_name.casefold()
    for child in direct_children(hwnd, win32gui, win32con):
        if (win32gui.GetClassName(child) or "").casefold() == target:
            return child
        nested = find_control_by_class(child, class_name, win32gui, win32con)
        if nested is not None:
            return nested
    return None


def find_ok_button_fallback(
    hwnd: int, win32gui: object, win32con: object
) -> int | None:
    by_id = find_control(hwnd, 1, win32gui, win32con)
    if by_id is not None:
        return by_id
    for child in direct_children(hwnd, win32gui, win32con):
        if (win32gui.GetClassName(child) or "").casefold() == "button":
            label = (win32gui.GetWindowText(child) or "").strip()
            if label in {"OK", "确定", "&OK"}:
                return child
        nested = find_ok_button_fallback(child, win32gui, win32con)
        if nested is not None:
            return nested
    return None


def collect_static_texts(hwnd: int, win32gui: object, win32con: object) -> list[str]:
    texts: list[str] = []
    for child in direct_children(hwnd, win32gui, win32con):
        if (win32gui.GetClassName(child) or "").casefold() == "static":
            label = (win32gui.GetWindowText(child) or "").strip()
            if label:
                texts.append(label)
        texts.extend(collect_static_texts(child, win32gui, win32con))
    return texts


def inspect_modal_dialogs(
    pids: set[int],
    win32gui: object,
    win32process: object,
    win32con: object,
    *,
    exclude_hwnd: int | None = None,
) -> list[dict[str, object]]:
    if not pids:
        return []
    windows: list[int] = []
    win32gui.EnumWindows(lambda hwnd, _: windows.append(hwnd), None)
    found: list[dict[str, object]] = []
    for hwnd in windows:
        if exclude_hwnd is not None and hwnd == exclude_hwnd:
            continue
        if not win32gui.IsWindowVisible(hwnd):
            continue
        _, candidate_pid = win32process.GetWindowThreadProcessId(hwnd)
        if candidate_pid not in pids:
            continue
        if win32gui.GetClassName(hwnd) != "#32770":
            continue
        found.append(
            {
                "hwnd": int(hwnd),
                "pid": int(candidate_pid),
                "title": (win32gui.GetWindowText(hwnd) or "").strip(),
                "static_texts": collect_static_texts(hwnd, win32gui, win32con),
            }
        )
    return found


def find_style_dialog(
    process_id: int,
    timeout: float,
    win32gui: object,
    win32process: object,
    win32con: object,
) -> tuple[int, int, int, int] | None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        windows: list[int] = []
        win32gui.EnumWindows(lambda hwnd, _: windows.append(hwnd), None)
        for hwnd in windows:
            _, candidate_pid = win32process.GetWindowThreadProcessId(hwnd)
            if candidate_pid != process_id or not win32gui.IsWindowVisible(hwnd):
                continue
            if win32gui.GetClassName(hwnd) != "#32770":
                continue
            combo = find_control(hwnd, 5049, win32gui, win32con)
            ok_button = find_control(hwnd, 1, win32gui, win32con)
            if combo is None:
                combo = find_control_by_class(hwnd, "ComboBox", win32gui, win32con)
            if ok_button is None:
                ok_button = find_ok_button_fallback(hwnd, win32gui, win32con)
            if combo is not None and ok_button is not None:
                ctrl_id = int(win32gui.GetDlgCtrlID(combo))
                return hwnd, combo, ok_button, ctrl_id
        time.sleep(0.1)
    return None


def document_style(document: object) -> str | None:
    try:
        value = document.Variables("EN.Layout").Value
    except Exception:
        return None
    match = re.search(r"<Style>(.*?)</Style>", str(value), flags=re.DOTALL)
    return match.group(1) if match else None


def collect_com_report(document: object) -> dict[str, object]:
    fields = []
    for field in document.Fields:
        code = str(field.Code.Text or "").strip()
        result = str(field.Result.Text or "")
        main_citation = code.startswith("ADDIN EN.CITE") and not code.startswith("ADDIN EN.CITE.DATA")
        fields.append(
            {
                "code": code,
                "result": result,
                "superscript": int(field.Result.Font.Superscript or 0) if main_citation else 0,
            }
        )
    main = [
        field
        for field in fields
        if field["code"].startswith("ADDIN EN.CITE")
        and field["code"] != "ADDIN EN.CITE.DATA"
    ]
    data_fields = [field for field in fields if field["code"] == "ADDIN EN.CITE.DATA"]
    reference_lists = [field for field in fields if field["code"] == "ADDIN EN.REFLIST"]
    bibliography = reference_lists[0]["result"] if reference_lists else ""
    entries = [line.strip() for line in bibliography.split("\r") if line.strip()]
    heading_detected = False
    if entries and entries[0].casefold() in BIBLIOGRAPHY_HEADINGS:
        heading_detected = True
        entries = entries[1:]
    return {
        "main_citation_fields": len(main),
        "citation_data_fields": len(data_fields),
        "reference_list_fields": len(reference_lists),
        "bibliography_heading_detected": heading_detected,
        "bibliography_entry_count": len(entries),
        "temporary_citations_remaining": len(
            TEMPORARY_CITATION_RE.findall(document.Content.Text)
        ),
        "all_main_citations_superscript": bool(main)
        and all(field["superscript"] == -1 for field in main),
        "citation_results": [field["result"] for field in main],
        "style_name": document_style(document),
    }


def report_complete(report: dict[str, object], expected: dict[str, object]) -> bool:
    if report["main_citation_fields"] != expected["main"]:
        return False
    if report["reference_list_fields"] != 1:
        return False
    if report["temporary_citations_remaining"] != 0:
        return False
    if expected.get("references") is not None and report["bibliography_entry_count"] != expected["references"]:
        return False
    if expected.get("style") is not None and report["style_name"] != expected["style"]:
        return False
    if expected.get("superscript") and not report["all_main_citations_superscript"]:
        return False
    return True


def wait_for_complete(document: object, expected: dict[str, object], timeout: float, pythoncom: object) -> dict[str, object]:
    deadline = time.monotonic() + timeout
    stable = 0
    last_signature: tuple[object, ...] | None = None
    last_report: dict[str, object] = {}
    last_error = ""
    while time.monotonic() < deadline:
        pythoncom.PumpWaitingMessages()
        try:
            last_report = collect_com_report(document)
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            time.sleep(0.5)
            continue
        signature = (
            last_report["main_citation_fields"],
            last_report["citation_data_fields"],
            last_report["reference_list_fields"],
            last_report["bibliography_entry_count"],
            last_report["temporary_citations_remaining"],
            last_report["style_name"],
        )
        if report_complete(last_report, expected):
            stable = stable + 1 if signature == last_signature else 1
            if stable >= 3:
                return last_report
        else:
            stable = 0
        last_signature = signature
        time.sleep(0.5)
    raise RuntimeError(f"CWYW output did not stabilize before timeout: {last_report}; last COM error: {last_error}")


RPC_E_CALL_REJECTED = -2147418111


def com_call_with_retry(
    fn,
    pythoncom: object,
    *,
    retries: int = 6,
    delay: float = 0.8,
):
    """Retry COM calls that return RPC_E_CALL_REJECTED (0x80010001) while Word/EndNote is busy."""
    for attempt in range(retries):
        try:
            return fn()
        except Exception as exc:
            if getattr(exc, "hresult", None) != RPC_E_CALL_REJECTED or attempt + 1 >= retries:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(delay)


def save_if_complete(
    document: object,
    output: Path,
    expected: dict[str, object],
    timeout: float,
    pythoncom: object,
) -> dict[str, object]:
    report = wait_for_complete(document, expected, timeout, pythoncom)
    com_call_with_retry(document.Repaginate, pythoncom)
    com_call_with_retry(lambda: document.SaveAs2(str(output), FileFormat=16), pythoncom)
    if not output.is_file():
        raise RuntimeError("Word SaveAs2 returned without creating the output")
    return report


def recover_running_document(
    source: Path,
    output: Path,
    expected: dict[str, object],
    timeout: float,
    word_pid: int,
) -> dict[str, object]:
    import pythoncom
    import win32com.client
    import win32process

    pythoncom.CoInitialize()
    document = None
    application = None
    owned = False
    try:
        candidate = win32com.client.GetActiveObject("Word.Application")
        document = next(
            (item for item in candidate.Documents if Path(item.FullName).resolve() == source),
            None,
        )
        if document is None:
            raise RuntimeError("No exact candidate document is present in the running Word application")
        owned = win32process.GetWindowThreadProcessId(document.ActiveWindow.Hwnd)[1] == word_pid
        if not owned:
            raise RuntimeError("Recovery document does not belong to this attempt's Word process")
        application = document.Application
        report = save_if_complete(document, output, expected, timeout, pythoncom)
        return {"status": "PASS", "report": report}
    finally:
        if owned and document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if owned and application is not None:
            try:
                if application.Documents.Count == 0:
                    application.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert EndNote temporary citations into genuine CWYW fields through Word."
    )
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--enl", type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--style")
    parser.add_argument("--expect-main", type=int)
    parser.add_argument("--expect-references", type=int)
    parser.add_argument("--expect-record-ids")
    parser.add_argument("--expected-source-sha256")
    parser.add_argument("--require-superscript", action="store_true")
    parser.add_argument("--timeout", type=float, default=180.0)
    parser.add_argument("--preflight-only", action="store_true")
    return parser


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    args = build_parser().parse_args()
    source = args.input.resolve()
    output = args.output.resolve()
    report_path = args.report.resolve() if args.report else None
    if sys.platform != "win32":
        raise RuntimeError("EndNote CWYW automation requires native Windows")
    if not source.is_file():
        raise FileNotFoundError(source)
    if source == output:
        raise RuntimeError("Input and output must be different paths")
    if output.exists():
        raise FileExistsError(f"Refusing to overwrite output: {output}")
    if report_path and report_path.exists():
        raise FileExistsError(f"Refusing to overwrite report: {report_path}")
    source_hash = sha256_file(source)
    if args.expected_source_sha256 and source_hash.casefold() != args.expected_source_sha256.casefold():
        raise RuntimeError(
            f"Frozen source hash changed: expected {args.expected_source_sha256}, found {source_hash}"
        )

    before = audit(source)
    temporary = temporary_state(source)
    existing_main = int(before.get("fields", {}).get("main_citation_fields", 0))
    inferred_main = existing_main + int(temporary["citation_groups"])
    expected_ids = parse_expected_ids(args.expect_record_ids)
    inferred_ids = sorted(
        set(before.get("fields", {}).get("record_ids_from_fields", []))
        | set(temporary["record_ids"])
    )
    expected = {
        "main": args.expect_main if args.expect_main is not None else inferred_main,
        "references": args.expect_references,
        "style": args.style,
        "superscript": args.require_superscript,
        "record_ids": expected_ids if expected_ids is not None else inferred_ids,
    }
    preflight_checks = {
        "source_hash_matches": True,
        "new_output_path": not output.exists(),
        "has_existing_or_temporary_citations": existing_main > 0
        or temporary["citation_groups"] > 0,
        "expected_main_not_less_than_existing": int(expected["main"]) >= existing_main,
        "explicit_record_ids_match_inferred": expected_ids is None
        or expected_ids == inferred_ids,
    }
    preflight = {
        "source": str(source),
        "source_sha256": source_hash,
        "output": str(output),
        "before_audit_status": before["status"],
        "existing_main_citations": existing_main,
        "temporary": temporary,
        "expected": expected,
        "checks": preflight_checks,
        "status": "PASS" if all(preflight_checks.values()) else "FAIL",
    }
    if args.preflight_only:
        print(json.dumps(preflight, ensure_ascii=False, indent=2))
        return 0 if preflight["status"] == "PASS" else 1
    if preflight["status"] != "PASS":
        print(json.dumps(preflight, ensure_ascii=False, indent=2))
        return 1

    try:
        import pythoncom
        import win32com.client
        import win32con
        import win32gui
        import win32process
    except ImportError as exc:
        raise RuntimeError("pywin32 is required in the active Windows environment") from exc

    shared: dict[str, object] = {}
    endnote_info: dict[str, object] | None = None
    endnote_proc = None
    endnote_close_info: dict[str, object] | None = None
    style_selection: dict[str, object] | None = None
    handled_style_hwnd: int | None = None
    unexpected_seen_at: dict[int, float] = {}
    existing_word_pids = {pid for pid, _ in list_process_pids({"WINWORD.EXE"})}

    def record_unexpected_dialogs() -> None:
        pids: set[int] = set()
        if shared.get("word_pid"):
            pids.add(int(shared["word_pid"]))
        if endnote_info and endnote_info.get("pids"):
            pids.update(int(p) for p in endnote_info["pids"])
        dialogs = inspect_modal_dialogs(
            pids,
            win32gui,
            win32process,
            win32con,
            exclude_hwnd=handled_style_hwnd,
        )
        now = time.monotonic()
        active_hwnds = {int(d["hwnd"]) for d in dialogs}
        for stale in [h for h in unexpected_seen_at if h not in active_hwnds]:
            unexpected_seen_at.pop(stale, None)
        for d in dialogs:
            hwnd = int(d["hwnd"])
            # Do not flag Configure Bibliography while waiting to handle it
            if shared.get("dialog_stage") and style_selection is None and not shared.get("style_selection_error"):
                if find_control(hwnd, 5049, win32gui, win32con) is not None or find_control_by_class(hwnd, "ComboBox", win32gui, win32con) is not None:
                    continue
            first = unexpected_seen_at.setdefault(hwnd, now)
            if now - first > 3.0:
                shared["unexpected_modal_dialog"] = {
                    **d,
                    "visible_seconds": round(now - first, 2),
                }

    def worker() -> None:
        pythoncom.CoInitialize()
        application = None
        document = None
        scratch = None
        owned = False
        previous_alerts = None
        try:
            application = win32com.client.DispatchEx("Word.Application")
            scratch = application.Documents.Add(Visible=False)
            pid = win32process.GetWindowThreadProcessId(scratch.Windows.Item(1).Hwnd)[1]
            if not pid or pid in existing_word_pids:
                raise RuntimeError("Word COM did not open the candidate in a dedicated process")
            owned = True
            shared["word_pid"] = pid
            application.Visible = True
            previous_alerts = application.DisplayAlerts
            application.DisplayAlerts = 0
            try:
                for addin in application.COMAddIns:
                    prog_id = str(getattr(addin, "ProgId", "") or "")
                    desc = str(getattr(addin, "Description", "") or "")
                    if (
                        "endnote.wordaddin.connect" in prog_id.casefold()
                        or "cite while you write" in desc.casefold()
                    ):
                        if not bool(addin.Connect):
                            addin.Connect = True
                        shared["cwyw_addin_connected"] = bool(addin.Connect)
                        break
            except Exception as addin_exc:
                shared["cwyw_addin_check_error"] = (
                    f"{type(addin_exc).__name__}: {addin_exc}"
                )
            scratch.Close(False)
            scratch = None
            document = application.Documents.Open(
                str(source), ReadOnly=False, AddToRecentFiles=False, Visible=True
            )
            if win32process.GetWindowThreadProcessId(document.ActiveWindow.Hwnd)[1] != pid:
                raise RuntimeError("Candidate document opened in an unverified Word window")
            shared["opened"] = True
            shared["format_started"] = True
            document.Activate()
            com_call_with_retry(
                lambda: application.Run("ENFormatBibliography"), pythoncom
            )
            shared["initial_format_returned"] = True
            active = application.ActiveDocument
            if Path(active.FullName).resolve() != source:
                raise RuntimeError("The original candidate is no longer the active Word document")
            document = active
            if args.style:
                current_style = document_style(document)
                shared["style_after_initial_format"] = current_style
                if current_style == args.style:
                    shared["style_already_active"] = True
                else:
                    shared["dialog_stage"] = True
                    com_call_with_retry(
                        lambda: application.Run("ENFormatBibliographyDlg"), pythoncom
                    )
                    shared["style_dialog_returned"] = True
                    com_call_with_retry(
                        lambda: application.Run("ENFormatBibliography"), pythoncom
                    )
                    shared["post_style_format_returned"] = True
            shared["com_report"] = save_if_complete(
                document, output, expected, args.timeout, pythoncom
            )
            shared["saved"] = True
        except Exception as exc:
            shared["error"] = f"{type(exc).__name__}: {exc}"
            shared["needs_recovery"] = bool(shared.get("format_started"))
        finally:
            if scratch is not None:
                try:
                    scratch.Close(False)
                except Exception:
                    pass
            if not shared.get("needs_recovery"):
                if document is not None:
                    try:
                        document.Close(False)
                    except Exception:
                        pass
                if owned and application is not None:
                    try:
                        if application.Documents.Count == 0:
                            application.Quit()
                        elif previous_alerts is not None:
                            application.DisplayAlerts = previous_alerts
                    except Exception:
                        pass
            pythoncom.CoUninitialize()

    thread = threading.Thread(target=worker, daemon=True)
    try:
        if args.enl is not None:
            endnote_info, endnote_proc = ensure_endnote_library_open(
                args.enl.resolve(), win32gui, win32process
            )
        thread.start()
        if args.style:
            deadline = time.monotonic() + args.timeout
            while time.monotonic() < deadline and not shared.get("dialog_stage"):
                if (
                    shared.get("error")
                    or shared.get("style_already_active")
                    or shared.get("saved")
                    or not thread.is_alive()
                ):
                    break
                record_unexpected_dialogs()
                time.sleep(0.1)
            if shared.get("style_already_active"):
                style_selection = {"already_active": True, "requested": args.style}
            elif shared.get("dialog_stage"):
                found = find_style_dialog(
                    int(shared["word_pid"]),
                    min(args.timeout, 60.0),
                    win32gui,
                    win32process,
                    win32con,
                )
                if found is None:
                    shared["style_selection_error"] = "Configure Bibliography dialog not found"
                else:
                    dialog, combo, ok_button, combo_ctrl_id = found
                    handled_style_hwnd = dialog
                    before_combo = combo_items(combo, win32gui)
                    items = list(before_combo["items"])
                    if args.style not in items:
                        win32gui.PostMessage(dialog, win32con.WM_CLOSE, 0, 0)
                        shared["style_selection_error"] = (
                            f"Requested style {args.style!r} is not in the CWYW selector"
                        )
                    else:
                        index = items.index(args.style)
                        selected = select_combo(combo, index)
                        parent = win32gui.GetParent(combo)
                        for notification in (CBN_SELCHANGE, CBN_SELENDOK):
                            win32gui.SendMessage(
                                parent,
                                win32con.WM_COMMAND,
                                combo_ctrl_id | (notification << 16),
                                combo,
                            )
                        after_combo = combo_items(combo, win32gui)
                        win32gui.SendMessage(ok_button, win32con.BM_CLICK, 0, 0)
                        style_selection = {
                            "before": before_combo,
                            "requested": args.style,
                            "selected_index": index,
                            "combo_ctrl_id": combo_ctrl_id,
                            "select_return": selected,
                            "after": after_combo,
                        }

        join_deadline = time.monotonic() + args.timeout + 90.0
        while thread.is_alive() and time.monotonic() < join_deadline:
            record_unexpected_dialogs()
            thread.join(timeout=0.25)
    finally:
        endnote_close_info = close_spawned_endnote(
            endnote_info, endnote_proc, win32gui, win32process
        )

    recovery: dict[str, object] | None = None
    # Recover only after the original COM worker has returned. Attaching while
    # it is still blocked in a modal dialog can corrupt the same live document.
    if (
        not output.exists()
        and shared.get("needs_recovery")
        and not thread.is_alive()
    ):
        try:
            recovery = recover_running_document(source, output, expected, min(args.timeout, 90.0), int(shared["word_pid"]))
            shared["recovered"] = True
        except Exception as exc:
            recovery = {"status": "FAIL", "error": f"{type(exc).__name__}: {exc}"}

    after: dict[str, object] | None = None
    if output.exists():
        after = audit(
            output,
            expect_main=int(expected["main"]),
            expect_references=(
                int(expected["references"])
                if expected["references"] is not None
                else None
            ),
            expect_record_ids=list(expected["record_ids"]),
            expect_style=str(expected["style"]) if expected["style"] else None,
            require_superscript=bool(expected["superscript"]),
        )

    success = bool(
        output.exists()
        and after is not None
        and after["status"] == "PASS"
        and not thread.is_alive()
        and not shared.get("style_selection_error")
        and (not args.enl or (endnote_close_info and endnote_close_info["cleanly_exited"]))
    )
    result = {
        "status": "PASS" if success else "FAIL",
        "preflight": preflight,
        "endnote_library": endnote_info,
        "endnote_close": endnote_close_info,
        "style_selection": style_selection,
        "worker": shared,
        "thread_alive": thread.is_alive(),
        "recovery": recovery,
        "output_audit": after,
        "output_sha256": sha256_file(output) if output.exists() else None,
    }
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if report_path:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    return 0 if success else 1


if __name__ == "__main__":
    raise SystemExit(main())
