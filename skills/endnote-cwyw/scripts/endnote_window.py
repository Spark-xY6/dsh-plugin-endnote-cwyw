"""Identify one EndNote library window before interacting with native Windows UI."""

import ctypes
from ctypes import wintypes
import os
from pathlib import Path
import time


class WindowInterference(RuntimeError):
    """EndNote no longer presents the isolated library or expected modal chain."""


def _process_start(proc):
    times = [wintypes.FILETIME() for _ in range(4)]
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    if not kernel.GetProcessTimes(wintypes.HANDLE(int(proc._handle)), *(ctypes.byref(t) for t in times)):
        raise ctypes.WinError(ctypes.get_last_error())
    return times[0].dwHighDateTime << 32 | times[0].dwLowDateTime


def _session_endnote_pids(pid):
    """Count other EndNote instances in the same Windows session, including hidden ones."""
    class ProcessInfo(ctypes.Structure):
        _fields_ = [("session", wintypes.DWORD), ("pid", wintypes.DWORD),
                    ("name", wintypes.LPWSTR), ("user_sid", ctypes.c_void_p)]

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
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
    wts.WTSFreeMemory.argtypes = [ctypes.c_void_p]
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


def _windows(pid):
    import win32gui
    import win32process
    found = {}
    def collect(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
            title = win32gui.GetWindowText(hwnd).strip()
            cls = win32gui.GetClassName(hwnd)
            if title or cls == "#32770":
                found[hwnd] = (title, cls, win32gui.GetWindow(hwnd, 4))  # GW_OWNER
    win32gui.EnumWindows(collect, None)
    return found


class TargetWindow:
    def __init__(self, proc, path, hwnd, cls, started):
        self.proc, self.path, self.hwnd, self.cls = proc, path, hwnd, cls
        self.started = started
        self.titles = {path.name.casefold(), ("EndNote 21 - " + path.name).casefold()}

    @classmethod
    def bind(cls, proc, path, timeout=35):
        path = Path(path).resolve(strict=True)
        db = path.with_suffix(".Data") / "sdb/sdb.eni"
        started = _process_start(proc)
        expected = {path.name.casefold(), ("EndNote 21 - " + path.name).casefold()}
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if proc.poll() is not None:
                raise WindowInterference("The launched EndNote process exited before library binding")
            windows = _windows(proc.pid)
            candidates = [h for h, (t, c, _) in windows.items()
                          if t.casefold() in expected and c.casefold().startswith("afx:")]
            if candidates:
                if len(windows) != 1:
                    raise WindowInterference("Another EndNote window is present during library binding")
                if (proc.pid, started) in _holders(db):
                    target = cls(proc, path, candidates[0], windows[candidates[0]][1], started)
                    target.require_exclusive()
                    return target
            time.sleep(.15)
        raise WindowInterference("Could not bind the new EndNote window to the exact candidate database")

    def _view(self):
        if self.proc.poll() is not None or _process_start(self.proc) != self.started:
            raise WindowInterference("The bound EndNote process no longer exists")
        if _session_endnote_pids(self.proc.pid) != {self.proc.pid}:
            raise WindowInterference("Another EndNote process is active in this Windows session")
        windows = _windows(self.proc.pid)
        title, cls, _ = windows.get(self.hwnd, ("", "", None))
        if title.casefold() not in self.titles or cls != self.cls:
            raise WindowInterference("The bound library HWND or title changed")
        return windows

    def _verify(self, windows, dialogs):
        if set(windows) != {self.hwnd, *dialogs}:
            raise WindowInterference("Another EndNote window or an unexpected dialog is open")
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
        while time.monotonic() < deadline:
            windows = self._view()
            if dialog in windows:
                self._verify(windows, (*parent_chain, dialog))
            elif set(windows) == {self.hwnd, *parent_chain}:
                self._verify(windows, parent_chain)
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
                            or title.casefold() not in self.titles):
                        raise WindowInterference("An unexpected window appeared during library shutdown")
                elif identity not in _holders(db):
                    released = True
                elif released:
                    raise WindowInterference("The closing process reopened the target database")
            time.sleep(.1)
        return False
