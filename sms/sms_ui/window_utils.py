def sync_and_focus_existing_window(window, sync_attr=None):
    if window is None:
        return False
    try:
        if not window.winfo_exists():
            return False
    except Exception:
        return False

    if sync_attr:
        try:
            sync_form = getattr(window, sync_attr, None)
            if sync_form:
                sync_form()
        except Exception:
            pass

    try:
        window.deiconify()
        window.lift()
        window.focus_force()
    except Exception:
        pass
    return True
