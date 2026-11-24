"""Identify one EndNote library window before interacting with native Windows UI."""

import ctypes
from ctypes import wintypes
import os
from pathlib import Path
import re
import time


class WindowInterference(RuntimeError):
    """EndNote no longer presents the isolated library or expected modal chain."""


def _process_start(proc):
    times = [wintypes.FILETIME() for _ in range(4)]
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.GetProcessTimes.argtypes = [
        wintypes.HANDLE,
        ctypes.POINTER(wintypes.FILETIME),
        ctypes.POINTER(wintypes.FILETIME),
        ctypes.POINTER(wintypes.FILETIME),
        ctypes.POINTER(wintypes.FILETIME),
    ]
    kernel.GetProcessTimes.restype = wintypes.BOOL
    if not kernel.GetProcessTimes(wintypes.HANDLE(int(proc._handle)), *(ctypes.byref(t) for t in times)):
        raise ctypes.WinError(ctypes.get_last_error())
    return times[0].dwHighDateTime << 32 | times[0].dwLowDateTime


def _session_endnote_pids(pid):
    """Count other EndNote instances in the same Windows session, including hidden ones."""
    class ProcessInfo(ctypes.Structure):
        _fields_ = [("session", wintypes.DWORD), ("pid", wintypes.DWORD),
                    ("name", wintypes.LPWSTR), ("user_sid", ctypes.c_void_p)]

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.ProcessIdToSessionId.argtypes = [wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
    kernel.ProcessIdToSessionId.restype = wintypes.BOOL
    session = wintypes.DWORD()
    if not kernel.ProcessIdToSessionId(wintypes.DWORD(pid), ctypes.byref(session)):
        raise ctypes.WinError(ctypes.get_last_error())
    wts = ctypes.WinDLL("wtsapi32", use_last_error=True)
    rows = ctypes.POINTER(ProcessInfo)()
    count = wintypes.DWORD()
    wts.WTSEnumerateProcessesW.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.DWORD,
        ctypes.POINTER(ctypes.POINTER(ProcessInfo)), ctypes.POINTER(wintypes.DWORD)
    ]
    wts.WTSEnumerateProcessesW.restype = wintypes.BOOL
    wts.WTSFreeMemory.argtypes = [ctypes.c_void_p]
    wts.WTSFreeMemory.restype = None
    if not wts.WTSEnumerateProcessesW(None, 0, 1, ctypes.byref(rows), ctypes.byref(count)):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return {row.pid for row in rows[:count.value]
                if row.session == session.value and (row.name or "").casefold() == "endnote.exe"}
    finally:
        wts.WTSFreeMemory(rows)


def _holders(resource):
    """Query Windows Restart Manager for PID + creation FILETIME using this exact file."""
    class Process(ctypes.Structure):
        _fields_ = [("pid", wintypes.DWORD), ("start", wintypes.FILETIME)]

    class Info(ctypes.Structure):
        _fields_ = [("process", Process), ("app", wintypes.WCHAR * 256),
                    ("service", wintypes.WCHAR * 64), ("kind", ctypes.c_int),
                    ("status", wintypes.ULONG), ("session", wintypes.DWORD), ("restartable", wintypes.BOOL)]

    rm = ctypes.WinDLL("Rstrtmgr")
    rm.RmStartSession.argtypes = [ctypes.POINTER(wintypes.DWORD), wintypes.DWORD, wintypes.LPWSTR]
    rm.RmStartSession.restype = wintypes.DWORD
    rm.RmRegisterResources.argtypes = [
        wintypes.DWORD, wintypes.UINT, ctypes.POINTER(wintypes.LPCWSTR),
        wintypes.UINT, ctypes.c_void_p, wintypes.UINT, ctypes.c_void_p,
    ]
    rm.RmRegisterResources.restype = wintypes.DWORD
    rm.RmGetList.argtypes = [
        wintypes.DWORD, ctypes.POINTER(wintypes.UINT), ctypes.POINTER(wintypes.UINT),
        ctypes.POINTER(Info), ctypes.POINTER(wintypes.DWORD),
    ]
    rm.RmGetList.restype = wintypes.DWORD
    rm.RmEndSession.argtypes = [wintypes.DWORD]
    rm.RmEndSession.restype = wintypes.DWORD
    session = wintypes.DWORD()
    key = ctypes.create_unicode_buffer(33)
    rc = rm.RmStartSession(ctypes.byref(session), 0, key)
    if rc:
        raise ctypes.WinError(rc)
    try:
        names = (wintypes.LPCWSTR * 1)(str(resource))
        rc = rm.RmRegisterResources(session, 1, names, 0, None, 0, None)
        if rc:
            raise ctypes.WinError(rc)
        needed, capacity, reasons = wintypes.UINT(), wintypes.UINT(0), wintypes.DWORD()
        entries = None
        for _ in range(4):
            rc = rm.RmGetList(session, ctypes.byref(needed), ctypes.byref(capacity), entries, ctypes.byref(reasons))
            if rc == 0:
                return {(p.process.pid, p.process.start.dwHighDateTime << 32 | p.process.start.dwLowDateTime)
                        for p in (entries or [])[:capacity.value]}
            if rc != 234:  # ERROR_MORE_DATA
                raise ctypes.WinError(rc)
            capacity.value = needed.value
            entries = (Info * capacity.value)()
        raise WindowInterference("The set of processes using the target library kept changing")
    finally:
        rm.RmEndSession(session)
        # Clear thread LastError=234 left behind by the RmGetList buffer probe
        ctypes.WinDLL("kernel32").SetLastError(0)


def _windows(pid):
    import win32gui
    import win32process
    found = {}
    total_seen = [0]

    def collect(hwnd, _):
        total_seen[0] += 1
        try:
            if win32gui.IsWindowVisible(hwnd) and win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
                title = win32gui.GetWindowText(hwnd).strip()
                cls = win32gui.GetClassName(hwnd)
                if title or cls == "#32770":
                    found[hwnd] = (title, cls, win32gui.GetWindow(hwnd, 4))  # GW_OWNER
        except win32gui.error:
            pass
        return True

    ctypes.WinDLL("kernel32").SetLastError(0)
    try:
        win32gui.EnumWindows(collect, None)
    except Exception as exc:
        raise WindowInterference(
            f"Cannot enumerate interactive desktop windows in this session ({exc}); "
            "run native EndNote import/formatting from an interactive Windows desktop."
        ) from exc
    if total_seen[0] == 0:
        raise WindowInterference(
            "No interactive desktop windows enumerated in this session; "
            "run native EndNote import/formatting from an interactive Windows desktop."
        )
    return found


PROMOTION_TITLE = "Holiday Sale: Save 25% when you upgrade to the new EndNote!"


def _promotion_cancel(dialog):
    """Recognize the tested EndNote 21 promotion by its visible native buttons."""
    import win32gui
    buttons = {}
    embedded = []

    def collect(hwnd, _):
        try:
            if win32gui.IsWindowVisible(hwnd):
                cls = win32gui.GetClassName(hwnd)
                if cls == "Internet Explorer_Server":
                    embedded.append(hwnd)
                if cls == "Button" and win32gui.GetParent(hwnd) == dialog:
                    buttons[win32gui.GetDlgCtrlID(hwnd)] = (win32gui.GetWindowText(hwnd), hwnd)
        except win32gui.error:
            pass
        return True

    try:
        win32gui.EnumChildWindows(dialog, collect, None)
    except win32gui.error:
        return None
    signature = {key: text for key, (text, _) in buttons.items()}
    if signature != {3617: "Learn More...", 2: "Cancel", 3619: "Ignore This"} or len(embedded) != 1:
        return None
    cancel = buttons[2][1]
    return cancel if win32gui.IsWindowEnabled(cancel) else None


def _is_hidden_promotion(dialog, cancel, frame, pid):
    """EndNote can retain a dismissed announcement as the owner of its import dialog."""
    import win32gui
    import win32process
    return (win32gui.IsWindow(dialog) and not win32gui.IsWindowVisible(dialog)
            and win32process.GetWindowThreadProcessId(dialog)[1] == pid
            and win32gui.GetClassName(dialog) == "#32770"
            and win32gui.GetWindowText(dialog) == PROMOTION_TITLE
            and win32gui.GetWindow(dialog, 4) == frame
            and win32gui.IsWindow(cancel) and win32gui.GetParent(cancel) == dialog
            and win32gui.GetDlgCtrlID(cancel) == 2 and win32gui.GetClassName(cancel) == "Button"
            and win32gui.GetWindowText(cancel) == "Cancel")


def _matches_library_title(title: str, path_name: str) -> bool:
    cleaned = title.strip()
    if not cleaned:
        return False
    if cleaned.casefold() == path_name.casefold():
        return True
    stem = re.escape(Path(path_name).stem)
    return bool(re.fullmatch(rf"endnote(?:\s+(?:x\d+|\d+(?:\.\d+)*))?\s+-\s+{stem}(?:\.enl)?", cleaned, re.IGNORECASE))


class TargetWindow:
    def __init__(self, proc, path, hwnd, cls, started):
        self.proc, self.path, self.hwnd, self.cls = proc, path, hwnd, cls
        self.started = started
        self.cancelled_promotion = None

    @classmethod
    def bind(cls, proc, path, timeout=35):
        path = Path(path).resolve(strict=True)
        db = path.with_suffix(".Data") / "sdb/sdb.eni"
        started = _process_start(proc)
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if proc.poll() is not None:
                raise WindowInterference("The launched EndNote process exited before library binding")
            windows = _windows(proc.pid)
            candidates = [h for h, (t, c, _) in windows.items()
                          if _matches_library_title(t, path.name) and c.casefold().startswith("afx:")]
            if candidates:
                if len(candidates) != 1:
                    raise WindowInterference("Multiple EndNote library windows matched the candidate. Close other libraries and retry.")
                if (proc.pid, started) in _holders(db):
                    target = cls(proc, path, candidates[0], windows[candidates[0]][1], started)
                    target.wait_for_startup(deadline)
                    return target
            time.sleep(.15)
        raise WindowInterference("Could not bind the new EndNote window to the exact candidate database")

    def _raw_view(self):
        if self.proc.poll() is not None or _process_start(self.proc) != self.started:
            raise WindowInterference("The bound EndNote process no longer exists")
        if _session_endnote_pids(self.proc.pid) != {self.proc.pid}:
            raise WindowInterference("Another EndNote process is active in this Windows session")
        windows = _windows(self.proc.pid)
        title, cls, _ = windows.get(self.hwnd, ("", "", None))
        if not _matches_library_title(title, self.path.name) or cls != self.cls:
            raise WindowInterference("The bound library HWND or title changed")
        if self.cancelled_promotion:
            dialog, cancel = self.cancelled_promotion
            for hwnd, (label, kind, owner) in windows.items():
                if (owner == dialog and kind == "#32770" and label in {"Import File", "Importing..."}
                        and _is_hidden_promotion(dialog, cancel, self.hwnd, self.proc.pid)):
                    windows[hwnd] = (label, kind, self.hwnd)
        return windows

    def _view(self):
        windows = self._raw_view()
        extra = set(windows) - {self.hwnd}
        if len(extra) != 1:
            return windows
        dialog = next(iter(extra))
        if windows[dialog] != (PROMOTION_TITLE, "#32770", self.hwnd):
            return windows
        self._verify(windows, (dialog,))
        cancel = _promotion_cancel(dialog)
        # The vendor creates the modal frame before painting its embedded page and buttons.
        for _ in range(15):
            if cancel:
                break
            time.sleep(.1)
            current = self._raw_view()
            if current != windows:
                return current
            self._verify(current, (dialog,))
            cancel = _promotion_cancel(dialog)
        if not cancel:
            return windows
        current = self._raw_view()
        self._verify(current, (dialog,))
        if current != windows or _promotion_cancel(dialog) != cancel:
            raise WindowInterference("The EndNote announcement changed. Review its window and retry.")
        import win32gui
        win32gui.SendMessageTimeout(cancel, 0x00F5, 0, 0, 0x0002, 2000)  # BM_CLICK, Cancel
        self.cancelled_promotion = (dialog, cancel)
        deadline = time.monotonic() + 3
        while time.monotonic() < deadline:
            current = self._raw_view()
            if dialog not in current:
                return current
            self._verify(current, (dialog,))
            if current[dialog] != windows[dialog]:
                break
            time.sleep(.08)
        raise WindowInterference("Close the announcement in the private EndNote library, then retry.")

    def _verify(self, windows, dialogs):
        if set(windows) != {self.hwnd, *dialogs}:
            raise WindowInterference("Review the open EndNote windows. Dismiss announcements normally, close other libraries when ready, and retry on a fresh copy.")
        owner = self.hwnd
        for hwnd in dialogs:
            if windows[hwnd][1:] != ("#32770", owner):
                raise WindowInterference("An EndNote dialog lost its expected owner")
            owner = hwnd
        if (self.proc.pid, self.started) not in _holders(self.path.with_suffix(".Data") / "sdb/sdb.eni"):
            raise WindowInterference("The process no longer holds the exact target database")

    def require_exclusive(self, dialogs=()):
        """Permit precisely this library and an explicitly observed, nested modal chain."""
        self._verify(self._view(), dialogs)

    def wait_for_startup(self, deadline):
        """Check delayed startup windows before opening an import or formatting dialog."""
        settled_at = time.monotonic() + 2.5
        while True:
            self.require_exclusive()
            now = time.monotonic()
            if now >= deadline:
                raise WindowInterference("EndNote startup did not settle before the timeout. Inspect its windows, then retry.")
            if now >= settled_at:
                return
            time.sleep(min(.15, settled_at - now, deadline - now))

    def wait_dialog(self, parent_chain, predicate, timeout):
        """Discover a new modal by *both* its direct owner and its expected controls."""
        deadline = time.monotonic() + timeout
        owner = parent_chain[-1] if parent_chain else self.hwnd
        while time.monotonic() < deadline:
            windows = self._view()
            others = set(windows) - {self.hwnd, *parent_chain}
            if not others:
                self._verify(windows, parent_chain)
            elif len(others) == 1:
                hwnd = others.pop()
                if windows[hwnd][1:] != ("#32770", owner):
                    raise WindowInterference("An unrelated window appeared while waiting for a dialog")
                self._verify(windows, (*parent_chain, hwnd))
                if predicate(hwnd, windows[hwnd][0]):
                    return hwnd
            else:
                raise WindowInterference("Unexpected EndNote windows during native import")
            time.sleep(.1)
        raise WindowInterference("The expected target-owned EndNote dialog did not appear")

    def wait_dismissed(self, parent_chain, dialog, timeout, *, progress=None):
        """Wait without clicking any progress UI that follows the confirmed operation."""
        deadline = time.monotonic() + timeout
        quiet_since = None
        saw_progress = False
        while time.monotonic() < deadline:
            windows = self._view()
            if dialog in windows:
                self._verify(windows, (*parent_chain, dialog))
                quiet_since = None
            elif set(windows) == {self.hwnd, *parent_chain}:
                self._verify(windows, parent_chain)
                if progress is None or saw_progress:
                    return
                # The import dialog can vanish shortly before EndNote creates its progress window.
                now = time.monotonic()
                if quiet_since is None:
                    quiet_since = now
                elif now - quiet_since >= .75:
                    return
            else:
                extra = set(windows) - {self.hwnd, *parent_chain}
                if progress is None or len(extra) != 1:
                    raise WindowInterference("An unexpected window appeared after confirming the dialog")
                successor = extra.pop()
                owner = parent_chain[-1] if parent_chain else self.hwnd
                if windows[successor][1:] != ("#32770", owner):
                    raise WindowInterference("The progress dialog has an unexpected owner or class")
                # Capture short-lived controls before the slower Restart Manager query.
                if not progress(successor, windows[successor][0]):
                    raise WindowInterference("Unexpected target-owned progress window")
                self._verify(windows, (*parent_chain, successor))
                saw_progress = True
                quiet_since = None
            time.sleep(.1)
        raise WindowInterference("The target-owned dialog did not complete")

    def close(self, timeout=15):
        """Close the bound frame once and wait for its original process to exit naturally."""
        import win32gui
        self.require_exclusive()
        win32gui.SendMessageTimeout(self.hwnd, 0x0010, 0, 0, 0x0002, 3000)
        db = self.path.with_suffix(".Data") / "sdb/sdb.eni"
        identity = (self.proc.pid, self.started)
        released = False
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            # The launcher remains in this session after the EndNote child exits.
            peers = _session_endnote_pids(os.getpid())
            if peers - {self.proc.pid}:
                raise WindowInterference("Another EndNote process appeared during library shutdown")
            exit_code = self.proc.poll()
            if exit_code is not None:
                if exit_code != 0:
                    return False
                if not peers and identity not in _holders(db):
                    return True
            else:
                if _process_start(self.proc) != self.started:
                    raise WindowInterference("The original EndNote process identity changed during shutdown")
                windows = _windows(self.proc.pid)
                if windows:
                    title, cls, owner = windows.get(self.hwnd, ("", "", None))
                    if (released or len(windows) != 1 or cls != self.cls or owner
                            or not _matches_library_title(title, self.path.name)):
                        raise WindowInterference("An unexpected window appeared during library shutdown")
                elif identity not in _holders(db):
                    released = True
                elif released:
                    raise WindowInterference("The closing process reopened the target database")
            time.sleep(.1)
        return False
