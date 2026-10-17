from sms_ui.missed_call_popup import open_missed_call_popup


def _popup_exists(window):
    try:
        return window is not None and bool(window.winfo_exists())
    except Exception:
        return False


def show_missed_call_popup_runtime(
    missed_call,
    *,
    parent,
    current_popup,
    set_popup,
    center_window,
    show_window,
    open_popup=open_missed_call_popup,
):
    caller_num = getattr(missed_call, "caller_num", "")
    started_at = getattr(missed_call, "started_at", None)

    if _popup_exists(current_popup):
        try:
            try:
                missed_count = max(
                    1,
                    int(getattr(current_popup, "missed_call_popup_count", 1)),
                ) + 1
            except (TypeError, ValueError):
                missed_count = 2
            current_popup.missed_call_popup_update(
                caller_num,
                started_at,
                missed_count,
            )
            current_popup.missed_call_popup_count = missed_count
            return "updated"
        except Exception:
            try:
                current_popup.destroy()
            except Exception:
                pass
            set_popup(None)

    popup_holder = {}

    def close_popup():
        popup = popup_holder.get("window")
        try:
            if _popup_exists(popup):
                popup.destroy()
        finally:
            set_popup(None)
            show_window()

    try:
        popup = open_popup(
            parent,
            caller_num,
            started_at,
            center_window,
            close_popup,
        )
        popup.missed_call_popup_count = 1
        popup_holder["window"] = popup
        set_popup(popup)
        return "shown"
    except Exception:
        set_popup(None)
        return "error"


def show_missed_call_popup_app_runtime(
    *,
    missed_call,
    parent,
    get_popup,
    set_popup,
    center_window,
    show_window,
    run_on_ui_thread,
    ui_post,
    is_enabled=lambda: True,
    get_active_call_popup=lambda: None,
    log_error=None,
    show_runtime=show_missed_call_popup_runtime,
):
    def show_on_ui():
        if not is_enabled():
            return "disabled"
        result = show_runtime(
            missed_call,
            parent=parent,
            current_popup=get_popup(),
            set_popup=set_popup,
            center_window=center_window,
            show_window=show_window,
        )
        # A delayed missed-call notice may be processed after the next call
        # has already opened. Both remain visible, with the live call on top.
        popup = get_active_call_popup()
        try:
            if _popup_exists(popup) and popup.winfo_ismapped():
                popup.lift()
        except Exception as exc:
            if log_error is not None:
                log_error(f"Raise active call above missed-call notice failed: {exc!r}")
        return result

    return run_on_ui_thread(show_on_ui, ui_post)
