from sms_ui.call_popup_runtime import close_call_popup_app_runtime, show_call_popup_app_runtime


def set_call_popup_namespace_runtime(namespace, window):
    namespace.__setitem__("current_call_popup", window)


def close_call_popup_namespace_runtime(namespace, *, close_app_runtime=close_call_popup_app_runtime):
    return close_app_runtime(
        get_popup=lambda: namespace["current_call_popup"],
        set_popup=lambda window: set_call_popup_namespace_runtime(namespace, window),
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
        log_error=namespace.get("log_file_only"),
    )


def show_call_popup_namespace_runtime(
    namespace,
    caller_num,
    *,
    show_app_runtime=show_call_popup_app_runtime,
):
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
        log_error=namespace.get("log_file_only"),
    )


def get_serial_call_state_namespace_runtime(namespace):
    return namespace["ring_timeout_target"], namespace["current_dial_num"]


def set_serial_call_state_namespace_runtime(namespace, next_ring_timeout, next_dial_num):
    namespace.__setitem__("ring_timeout_target", next_ring_timeout)
    namespace.__setitem__("current_dial_num", next_dial_num)
