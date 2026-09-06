from sms_ui.call_popup_runtime import close_call_popup_app_runtime, show_call_popup_app_runtime
from sms_ui.missed_call_popup_runtime import show_missed_call_popup_app_runtime


def set_call_popup_namespace_runtime(namespace, window):
    namespace.__setitem__("current_call_popup", window)


def set_dial_popup_namespace_runtime(namespace, window):
    namespace.__setitem__("current_dial_popup", window)


def mark_dial_popup_connected_namespace_runtime(namespace):
    def mark_on_ui():
        popup = namespace.get("current_dial_popup")
        if popup is None:
            return False
        try:
            if not popup.winfo_exists():
                return False
        except Exception:
            return False
        marker = getattr(popup, "_call_popup_mark_connected", None)
        if not callable(marker):
            return False
        try:
            marker()
            return True
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if callable(log_error):
                try:
                    log_error(f"Mark dial popup connected failed: {exc!r}")
                except Exception:
                    pass
            return False

    return namespace["run_on_ui_thread"](mark_on_ui, namespace["ui_post"])


def mark_call_popup_connected_namespace_runtime(namespace):
    def mark_on_ui():
        popup = namespace.get("current_call_popup")
        if popup is None:
            return False
        try:
            if not popup.winfo_exists():
                return False
        except Exception:
            return False
        marker = getattr(popup, "_call_popup_mark_connected", None)
        if not callable(marker):
            return False
        try:
            marker()
            return True
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if callable(log_error):
                try:
                    log_error(f"Mark incoming call popup connected failed: {exc!r}")
                except Exception:
                    pass
            return False

    return namespace["run_on_ui_thread"](mark_on_ui, namespace["ui_post"])


def finish_dial_popup_namespace_runtime(namespace, message=""):
    def finish_on_ui():
        popup = namespace.get("current_dial_popup")
        if popup is None:
            return False
        try:
            if not popup.winfo_exists():
                return False
        except Exception:
            return False
        marker = getattr(popup, "_call_popup_mark_ended", None)
        if not callable(marker):
            return False
        try:
            marker(message)
            return True
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if callable(log_error):
                try:
                    log_error(f"Mark dial popup ended failed: {exc!r}")
                except Exception:
                    pass
            return False

    return namespace["run_on_ui_thread"](finish_on_ui, namespace["ui_post"])


def set_missed_call_popup_namespace_runtime(namespace, window):
    namespace.__setitem__("current_missed_call_popup", window)


def start_incoming_call_session_namespace_runtime(namespace, caller_num):
    return namespace["INCOMING_CALL_SESSION"].start(caller_num)


def mark_incoming_call_handled_namespace_runtime(namespace):
    return namespace["INCOMING_CALL_SESSION"].mark_handled()


def finish_incoming_call_session_namespace_runtime(namespace):
    return namespace["INCOMING_CALL_SESSION"].finish()


def reset_incoming_call_session_namespace_runtime(namespace):
    return namespace["INCOMING_CALL_SESSION"].reset()


def close_call_popup_namespace_runtime(namespace, *, close_app_runtime=close_call_popup_app_runtime):
    result = close_app_runtime(
        get_popup=lambda: namespace["current_call_popup"],
        set_popup=lambda window: set_call_popup_namespace_runtime(namespace, window),
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
        log_error=namespace.get("log_file_only"),
    )
    if namespace.get("current_dial_popup") is not None:
        dial_result = close_app_runtime(
            get_popup=lambda: namespace.get("current_dial_popup"),
            set_popup=lambda window: set_dial_popup_namespace_runtime(namespace, window),
            run_on_ui_thread=namespace["run_on_ui_thread"],
            ui_post=namespace["ui_post"],
            log_error=namespace.get("log_file_only"),
        )
        if result is None:
            result = dial_result
    return result


def close_missed_call_popup_namespace_runtime(namespace, *, close_app_runtime=close_call_popup_app_runtime):
    return close_app_runtime(
        get_popup=lambda: namespace["current_missed_call_popup"],
        set_popup=lambda window: set_missed_call_popup_namespace_runtime(namespace, window),
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
        log_error=namespace.get("log_file_only"),
    )


def close_phone_popups_namespace_runtime(namespace):
    for callback_name in ("close_call_popup", "close_missed_call_popup"):
        callback = namespace.get(callback_name)
        if not callable(callback):
            continue
        try:
            callback()
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if callable(log_error):
                try:
                    log_error(f"Close {callback_name} after disabling phone popups failed: {exc!r}")
                except Exception:
                    pass


def show_call_popup_namespace_runtime(
    namespace,
    caller_num,
    *,
    show_app_runtime=show_call_popup_app_runtime,
):
    if not namespace.get("CALL_POPUP_ENABLED", True):
        return "disabled"
    return show_app_runtime(
        parent=namespace["root"],
        caller_num=caller_num,
        get_popup=lambda: namespace["current_call_popup"],
        set_popup=lambda window: set_call_popup_namespace_runtime(namespace, window),
        center_window=namespace["center_window"],
        serial_lock=namespace["serial_lock"],
        get_serial=lambda: namespace["serial_obj"],
        port_ui=namespace["port_ui"],
        set_status=namespace["set_status"],
        ui_post=namespace["ui_post"],
        close_popup=namespace["close_call_popup"],
        set_ring_timeout=lambda value: namespace.__setitem__("ring_timeout_target", value),
        run_on_ui_thread=namespace["run_on_ui_thread"],
        is_enabled=lambda: bool(namespace.get("CALL_POPUP_ENABLED", True)),
        mark_call_handled=lambda: mark_incoming_call_handled_namespace_runtime(namespace),
        log_error=namespace.get("log_file_only"),
    )


def show_missed_call_popup_namespace_runtime(
    namespace,
    missed_call,
    *,
    show_app_runtime=show_missed_call_popup_app_runtime,
):
    if not namespace.get("CALL_POPUP_ENABLED", True):
        return "disabled"
    return show_app_runtime(
        missed_call=missed_call,
        parent=namespace["root"],
        get_popup=lambda: namespace["current_missed_call_popup"],
        set_popup=lambda window: set_missed_call_popup_namespace_runtime(namespace, window),
        center_window=namespace["center_window"],
        show_window=namespace["show_window"],
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
        is_enabled=lambda: bool(namespace.get("CALL_POPUP_ENABLED", True)),
    )


def get_serial_call_state_namespace_runtime(namespace):
    return namespace["ring_timeout_target"], namespace["current_dial_num"]


def set_serial_call_state_namespace_runtime(namespace, next_ring_timeout, next_dial_num):
    namespace.__setitem__("ring_timeout_target", next_ring_timeout)
    namespace.__setitem__("current_dial_num", next_dial_num)
