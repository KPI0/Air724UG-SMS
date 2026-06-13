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
