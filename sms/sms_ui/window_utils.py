import ctypes
from ctypes import wintypes
import sys


def _window_work_area(win, parent):
    fallback = (0, 0, win.winfo_screenwidth(), win.winfo_screenheight(), 0, 0)
    if sys.platform != "win32":
        return fallback
    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)

        class MonitorInfo(ctypes.Structure):
            _fields_ = [("size", wintypes.DWORD), ("monitor", wintypes.RECT),
                        ("work", wintypes.RECT), ("flags", wintypes.DWORD)]

        user32.MonitorFromRect.argtypes = [ctypes.POINTER(wintypes.RECT), wintypes.DWORD]
        user32.MonitorFromRect.restype = wintypes.HANDLE
        user32.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MonitorInfo)]
        user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
        user32.GetAncestor.restype = wintypes.HWND
        user32.GetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.GetMenu.argtypes = [wintypes.HWND]
        user32.GetMenu.restype = wintypes.HMENU
        user32.AdjustWindowRectEx.argtypes = [ctypes.POINTER(wintypes.RECT), wintypes.DWORD,
                                             wintypes.BOOL, wintypes.DWORD]
        x, y = parent.winfo_rootx(), parent.winfo_rooty()
        rect = wintypes.RECT(x, y, x + parent.winfo_width(), y + parent.winfo_height())
        monitor = user32.MonitorFromRect(ctypes.byref(rect), 2)
        info = MonitorInfo()
        info.size = ctypes.sizeof(info)
        if not user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
            return fallback
        # Querying the native frame does not map a withdrawn Tk window.
        handle = user32.GetAncestor(win.winfo_id(), 2)
        border = wintypes.RECT()
        if not user32.AdjustWindowRectEx(
            ctypes.byref(border), user32.GetWindowLongW(handle, -16),
            bool(user32.GetMenu(handle)), user32.GetWindowLongW(handle, -20),
        ):
            border = wintypes.RECT()
        area = info.work
        return (area.left, area.top, area.right, area.bottom,
                border.right - border.left, border.bottom - border.top)
    except (AttributeError, OSError, TypeError):
        return fallback


def fit_window_position(win, parent, width, height, x, y):
    left, top, right, bottom, border_width, border_height = _window_work_area(win, parent)
    x = max(left, min(x, right - width - border_width))
    y = max(top, min(y, bottom - height - border_height))
    return x, y


def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def sync_and_focus_existing_window(window, sync_attr=None, *, log_error=None):
    if window is None:
        return False
    try:
        if not window.winfo_exists():
            return False
    except Exception as exc:
        _safe_log(log_error, f"Check existing window failed: {exc!r}")
        return False

    if sync_attr:
        try:
            sync_form = getattr(window, sync_attr, None)
            if sync_form:
                sync_form()
        except Exception as exc:
            _safe_log(log_error, f"Sync existing window failed: {exc!r}")

    try:
        window.deiconify()
        window.lift()
        window.focus_force()
    except Exception as exc:
        _safe_log(log_error, f"Focus existing window failed: {exc!r}")
    return True
