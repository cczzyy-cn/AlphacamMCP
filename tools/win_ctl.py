# -*- coding: utf-8 -*-
"""Enumerate top-level windows with their owning PID/class (ctypes, no compiler
needed) so we can locate AlphaCAM's modal error dialog.

    python tools/win_ctl.py [--dismiss PID]
"""
import ctypes
import ctypes.wintypes as wt
import sys

user32 = ctypes.WinDLL("user32", use_last_error=True)

WNDENUMPROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
user32.EnumWindows.argtypes = [WNDENUMPROC, wt.LPARAM]
user32.GetWindowTextLengthW.argtypes = [wt.HWND]
user32.GetWindowTextW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
user32.GetClassNameW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
user32.IsWindowVisible.argtypes = [wt.HWND]
user32.GetWindowThreadProcessId.argtypes = [wt.HWND, ctypes.POINTER(wt.DWORD)]

WM_COMMAND = 0x0111
WM_CLOSE = 0x0010
IDOK = 1


def text(hwnd):
    n = user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(n + 2)
    user32.GetWindowTextW(hwnd, buf, n + 2)
    return buf.value


def cls(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def enum():
    rows = []

    @WNDENUMPROC
    def cb(hwnd, _):
        pid = wt.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        rows.append((hwnd, pid.value, cls(hwnd), text(hwnd),
                     bool(user32.IsWindowVisible(hwnd))))
        return True

    user32.EnumWindows(cb, 0)
    return rows


def main():
    want_dismiss = None
    only_pid = None
    if "--dismiss" in sys.argv:
        want_dismiss = int(sys.argv[sys.argv.index("--dismiss") + 1])
    if "--pid" in sys.argv:
        only_pid = int(sys.argv[sys.argv.index("--pid") + 1])
    if "--visible" in sys.argv:
        globals()["ONLY_VISIBLE"] = True

    rows = enum()
    only_visible = "--visible" in sys.argv
    print("%-10s %-7s %-22s %-6s %s" % ("HWND", "PID", "CLASS", "VIS", "TITLE"))
    for hwnd, pid, c, t, vis in sorted(rows, key=lambda r: r[1]):
        if only_pid is not None and pid != only_pid:
            continue
        if only_pid is None and only_visible and not vis:
            continue
        print("%-10d %-7d %-22s %-6s %s" % (hwnd, pid, c[:22], vis, t[:70]))

    if want_dismiss is not None:
        print("\n-- dialogs (#32770) of pid %d --" % want_dismiss)
        n = 0
        for hwnd, pid, c, t, vis in rows:
            if pid == want_dismiss and c == "#32770":
                print("   hwnd=%d class=%s vis=%s title=%r" % (hwnd, c, vis, t))
                rc = user32.PostMessageW(hwnd, WM_COMMAND, IDOK, 0)
                print("      PostMessage(WM_COMMAND, IDOK) -> %s" % bool(rc))
                n += 1
        print("   %d dialog(s) handled" % n)

    if "--focus" in sys.argv:
        hwnd = int(sys.argv[sys.argv.index("--focus") + 1])
        user32.ShowWindow(hwnd, 9)          # SW_RESTORE
        user32.SetForegroundWindow(hwnd)
        print("\nfocused hwnd=%d -> %r (foreground=%d)"
              % (hwnd, text(hwnd), user32.GetForegroundWindow()))

    if "--close-title" in sys.argv:
        needle = sys.argv[sys.argv.index("--close-title") + 1]
        print("\n-- closing windows whose title contains %r --" % needle)
        n = 0
        for hwnd, pid, c, t, vis in rows:
            if needle.lower() in t.lower():
                rc = user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
                print("   hwnd=%-10d pid=%-6d %-14s %r -> %s" % (hwnd, pid, c, t, bool(rc)))
                n += 1
        print("   %d window(s) asked to close" % n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
