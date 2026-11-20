# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
#     "PyMuPDF>=1.24.0",
#     "python-docx>=1.1.0",
#     "lxml>=5.0.0",
#     "httpx>=0.27.0",
# ]
# ///

from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import importlib
import json
import os
import re
import sys
import winreg
from pathlib import Path


REQUIRED_IMPORTS = [
    "win32com.client",
    "win32gui",
    "win32process",
    "fitz",
    "docx",
    "lxml",
    "httpx",
]


def query_registry_default(hive: int, subkey: str) -> str | None:
    try:
        with winreg.OpenKey(hive, subkey) as key:
            val, _ = winreg.QueryValueEx(key, "")
            return str(val).strip()
    except OSError:
        return None


def parse_exe_command(raw: str) -> Path | None:
    cleaned = os.path.expandvars(raw.strip())
    if not cleaned:
        return None
    if cleaned.startswith('"'):
        parts = cleaned.split('"', 2)
        if len(parts) >= 2:
            cand = Path(parts[1])
            return cand if cand.is_file() else None
    direct = Path(cleaned)
    if direct.is_file():
        return direct
    match = re.match(r"^(.*?\.exe)(?:\s|$)", cleaned, flags=re.IGNORECASE)
    if match:
        cand = Path(match.group(1))
        if cand.is_file():
            return cand
    return None


def resolve_exe_from_reg(queries: list[tuple[int, str]]) -> Path | None:
    for hive, subkey in queries:
        raw = query_registry_default(hive, subkey)
        if not raw:
            continue
        candidate = parse_exe_command(raw)
        if candidate is not None:
            return candidate
    return None


def check_cwyw_addin() -> list[dict[str, object]]:
    locations = [
        ("HKCU", winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\Word\Addins\EndNote.WordAddin.Connect"),
        ("HKLM", winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\Word\Addins\EndNote.WordAddin.Connect"),
        ("HKLM_WOW64", winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Office\Word\Addins\EndNote.WordAddin.Connect"),
    ]
    entries: list[dict[str, object]] = []
    for root_name, hive, subkey in locations:
        try:
            with winreg.OpenKey(hive, subkey) as key:
                load_behavior = None
                friendly_name = None
                try:
                    load_behavior, _ = winreg.QueryValueEx(key, "LoadBehavior")
                except OSError:
                    pass
                try:
                    friendly_name, _ = winreg.QueryValueEx(key, "FriendlyName")
                except OSError:
                    pass
                entries.append(
                    {
                        "hive": root_name,
                        "subkey": subkey,
                        "friendly_name": friendly_name,
                        "load_behavior": load_behavior,
                    }
                )
        except OSError:
            continue
    return entries


def list_processes_snapshot(image_names: set[str]) -> list[dict[str, object]]:
    """Fallback process enumeration using Win32 Toolhelp snapshot."""
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
    snap = kernel32.CreateToolhelp32Snapshot(0x00000002, 0)
    if snap == ctypes.c_void_p(-1).value:
        return []
    targets = {n.casefold() for n in image_names}
    rows: list[dict[str, object]] = []
    try:
        entry = PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        ok = kernel32.Process32FirstW(snap, ctypes.byref(entry))
        while ok:
            if entry.szExeFile.casefold() in targets:
                rows.append({"pid": int(entry.th32ProcessID), "name": entry.szExeFile, "command_line": ""})
            ok = kernel32.Process32NextW(snap, ctypes.byref(entry))
    finally:
        kernel32.CloseHandle(snap)
    return rows


def inspect_running_processes() -> list[dict[str, object]]:
    """Enumerate running WINWORD.EXE and EndNote.exe with command line and visible titles."""
    procs: list[dict[str, object]] = []
    try:
        import win32com.client

        wmi = win32com.client.GetObject("winmgmts:")
        items = wmi.ExecQuery(
            "SELECT ProcessId, Name, CommandLine FROM Win32_Process WHERE Name='WINWORD.EXE' OR Name='EndNote.exe'"
        )
        for item in items:
            procs.append(
                {
                    "pid": int(item.ProcessId),
                    "name": str(item.Name),
                    "command_line": str(item.CommandLine or ""),
                }
            )
    except Exception:
        procs = list_processes_snapshot({"WINWORD.EXE", "EndNote.exe"})

    try:
        import win32gui
        import win32process

        windows: list[int] = []
        win32gui.EnumWindows(lambda hwnd, _: windows.append(hwnd), None)
        titles_by_pid: dict[int, list[str]] = {}
        for hwnd in windows:
            if not win32gui.IsWindowVisible(hwnd):
                continue
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            title = (win32gui.GetWindowText(hwnd) or "").strip()
            if title:
                titles_by_pid.setdefault(int(pid), []).append(title)
    except Exception:
        titles_by_pid = {}

    for proc in procs:
        titles = titles_by_pid.get(int(proc["pid"]), [])
        proc["visible_window_titles"] = titles
        cmd = str(proc.get("command_line", "")).casefold()
        proc["is_hidden_automation"] = (
            str(proc["name"]).casefold() == "winword.exe"
            and ("-embedding" in cmd or "/automation" in cmd)
            and not titles
        )
    return procs


def collect_styles(endnote_exe: Path | None) -> dict[str, object]:
    search_dirs: list[Path] = []
    if endnote_exe is not None:
        search_dirs.append(endnote_exe.parent / "Styles")
    user_profile = Path(os.environ.get("USERPROFILE", str(Path.home())))
    search_dirs.append(user_profile / "Documents" / "EndNote" / "Styles")
    app_data = os.environ.get("APPDATA")
    if app_data:
        search_dirs.append(Path(app_data) / "EndNote" / "Styles")

    style_names: set[str] = set()
    checked_dirs: list[str] = []
    for folder in search_dirs:
        if folder.is_dir():
            checked_dirs.append(str(folder))
            for ens in folder.glob("*.ens"):
                style_names.add(ens.name)
    return {
        "style_dirs": checked_dirs,
        "style_count": len(style_names),
        "has_nature_ens": any(name.casefold() == "nature.ens" for name in style_names),
        "sample_nature_styles": sorted(
            name for name in style_names if "nature" in name.casefold()
        )[:15],
    }


def run_doctor(enl: Path | None, pdf_dir: Path | None) -> dict[str, object]:
    import_results: dict[str, bool] = {}
    import_errors: dict[str, str] = {}
    for mod in REQUIRED_IMPORTS:
        try:
            importlib.import_module(mod)
            import_results[mod] = True
        except Exception as exc:
            import_results[mod] = False
            import_errors[mod] = f"{type(exc).__name__}: {exc}"

    word_exe = resolve_exe_from_reg(
        [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"),
        ]
    )
    endnote_exe = resolve_exe_from_reg(
        [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe"),
            (winreg.HKEY_CLASSES_ROOT, r"EndNotePlus.ENW\shell\open\command"),
        ]
    )
    addin_entries = check_cwyw_addin()
    cwyw_load_ok = any(entry.get("load_behavior") == 3 for entry in addin_entries)
    running_procs = inspect_running_processes() if sys.platform == "win32" else []
    styles_info = collect_styles(endnote_exe)

    enl_info: dict[str, object] | None = None
    enl_ok = True
    if enl is not None:
        resolved_enl = enl.resolve()
        sdb_path = resolved_enl.parent / f"{resolved_enl.stem}.Data" / "sdb" / "sdb.eni"
        enl_ok = resolved_enl.is_file() and sdb_path.is_file()
        enl_info = {
            "enl_path": str(resolved_enl),
            "enl_exists": resolved_enl.is_file(),
            "sdb_path": str(sdb_path),
            "sdb_exists": sdb_path.is_file(),
            "sdb_size": sdb_path.stat().st_size if sdb_path.is_file() else 0,
        }

    pdf_info: dict[str, object] | None = None
    pdf_ok = True
    if pdf_dir is not None:
        resolved_dir = pdf_dir.resolve()
        pdf_files = sorted(resolved_dir.rglob("*.pdf")) if resolved_dir.is_dir() else []
        pdf_ok = resolved_dir.is_dir() and len(pdf_files) > 0
        pdf_info = {
            "pdf_dir": str(resolved_dir),
            "exists": resolved_dir.is_dir(),
            "pdf_count": len(pdf_files),
        }

    hard_checks = {
        "is_windows": sys.platform == "win32",
        "required_imports_ok": all(import_results.values()),
        "word_installed": word_exe is not None,
        "endnote_installed": endnote_exe is not None,
        "cwyw_addin_load_behavior_3": cwyw_load_ok,
        "enl_valid_if_supplied": enl_ok,
        "pdf_dir_valid_if_supplied": pdf_ok,
    }
    warnings: list[str] = []
    if any(p.get("is_hidden_automation") for p in running_procs):
        warnings.append("Hidden /Automation -Embedding WINWORD.EXE process detected; verify before formatting.")
    if any(str(p.get("name", "")).casefold() == "endnote.exe" for p in running_procs):
        warnings.append("EndNote.exe is currently running; close EndNote before taking a clean sdb.eni snapshot or syncing.")
    if not styles_info.get("has_nature_ens"):
        warnings.append("Nature.ens was not found in scanned EndNote style folders.")

    if not all(hard_checks.values()):
        status = "FAIL"
    elif warnings:
        status = "WARN"
    else:
        status = "PASS"

    return {
        "status": status,
        "python": {
            "executable": sys.executable,
            "version": sys.version.split()[0],
            "imports": import_results,
            "import_errors": import_errors,
        },
        "word_exe": str(word_exe) if word_exe else None,
        "endnote_exe": str(endnote_exe) if endnote_exe else None,
        "cwyw_addin": addin_entries,
        "running_processes": running_procs,
        "styles": styles_info,
        "enl": enl_info,
        "pdf_dir": pdf_info,
        "checks": hard_checks,
        "warnings": warnings,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Preflight environment diagnostics for Windows Word + EndNote CWYW automation."
    )
    parser.add_argument("--enl", type=Path)
    parser.add_argument("--pdf-dir", type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--overwrite", action="store_true")
    return parser


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    args = build_parser().parse_args()
    result = run_doctor(args.enl, args.pdf_dir)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        report = args.report.resolve()
        if report.exists() and not args.overwrite:
            raise FileExistsError(f"Refusing to overwrite report: {report}")
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    return 1 if result["status"] == "FAIL" else 0


if __name__ == "__main__":
    raise SystemExit(main())
