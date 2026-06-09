import ctypes


EXISTING_INSTANCE_ERRORS = (183, 5)


def port_mutex_name(port_name):
    return f"Air724UG_PORT_{port_name}"


def normalize_serial_device_path(port_name):
    port = str(port_name or "").strip()
    if not port:
        return ""
    if port.startswith("\\\\.\\"):
        return port
    return "\\\\.\\" + port


def create_named_mutex(mutex_name):
    return ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)


def close_windows_handle(handle):
    if not handle:
        return
    ctypes.windll.kernel32.CloseHandle(handle)


def release_mutex_handle(handle):
    if not handle:
        return
    try:
        ctypes.windll.kernel32.ReleaseMutex(handle)
    except Exception:
        pass
    try:
        close_windows_handle(handle)
    except Exception:
        pass


def acquire_mutex_with_error(mutex_name):
    handle = create_named_mutex(mutex_name)
    return handle, ctypes.windll.kernel32.GetLastError()


def is_existing_instance_error(last_error):
    return last_error in EXISTING_INSTANCE_ERRORS


def show_message_box(title, message, style=0x30):
    ctypes.windll.user32.MessageBoxW(0, str(message), str(title), style)


def focus_existing_window(window_title):
    """Best-effort restore of a running Tk window with the given title."""
    try:
        user32 = ctypes.windll.user32
        hwnd = user32.FindWindowW(None, window_title)
        if not hwnd:
            return False

        sw_show = 5
        sw_restore = 9
        swp_nosize = 0x0001
        swp_nomove = 0x0002
        swp_showwindow = 0x0040
        hwnd_topmost = -1
        hwnd_notopmost = -2

        try:
            if user32.IsIconic(hwnd):
                user32.ShowWindow(hwnd, sw_restore)
            else:
                user32.ShowWindow(hwnd, sw_show)
                user32.ShowWindow(hwnd, sw_restore)
        except Exception:
            pass

        try:
            user32.SetWindowPos(
                hwnd,
                hwnd_topmost,
                0,
                0,
                0,
                0,
                swp_nomove | swp_nosize | swp_showwindow,
            )
            user32.SetWindowPos(
                hwnd,
                hwnd_notopmost,
                0,
                0,
                0,
                0,
                swp_nomove | swp_nosize | swp_showwindow,
            )
        except Exception:
            pass

        for action in (user32.BringWindowToTop, user32.SetForegroundWindow, user32.SetActiveWindow):
            try:
                action(hwnd)
            except Exception:
                pass

        return True
    except Exception:
        return False


def is_port_locked_by_other(port_name):
    """
    Best-effort exclusive-open probe for a serial port.

    Returns True when Windows refuses an exclusive handle, and False when the
    port can be opened or the probe itself fails.
    """
    try:
        port = normalize_serial_device_path(port_name)
        if not port:
            return True

        generic_read = 0x80000000
        generic_write = 0x40000000
        open_existing = 3
        invalid_handle_value = ctypes.c_void_p(-1).value

        kernel32 = ctypes.windll.kernel32
        create_file = kernel32.CreateFileW
        create_file.argtypes = [
            ctypes.c_wchar_p,
            ctypes.c_ulong,
            ctypes.c_ulong,
            ctypes.c_void_p,
            ctypes.c_ulong,
            ctypes.c_ulong,
            ctypes.c_void_p,
        ]
        create_file.restype = ctypes.c_void_p

        close_handle = kernel32.CloseHandle
        close_handle.argtypes = [ctypes.c_void_p]
        close_handle.restype = ctypes.c_int

        handle = create_file(
            port,
            generic_read | generic_write,
            0,
            None,
            open_existing,
            0,
            None,
        )
        if handle in (None, 0, invalid_handle_value):
            return True

        try:
            return False
        finally:
            close_handle(handle)
    except Exception:
        return False


def request_dpi_awareness():
    try:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass
