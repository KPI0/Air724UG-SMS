import inspect

from sms_ui.serial_debug_window import open_serial_debug_window_dialog


def open_serial_debug_window_runtime(
    parent,
    *,
    get_state,
    set_state,
    serial_queue,
    serial_lock,
    get_serial_obj,
    push_serial_debug,
    port_ui,
    set_status,
    format_connected_status,
    get_port,
    center_window,
    window_title="串口调试",
    log_error=None,
    get_dial_popup=lambda: None,
    set_dial_popup=lambda _window: None,
    open_dialog=open_serial_debug_window_dialog,
):
    current_window, current_text = get_state("window_refs")
    args = (
        parent,
        current_window,
        current_text,
        get_state("debug_enabled"),
        lambda: get_state("drop_count"),
        serial_queue,
        serial_lock,
        get_serial_obj,
        push_serial_debug,
        port_ui,
        set_status,
        format_connected_status,
        get_port,
        lambda phone: set_state("current_dial_num", phone),
        lambda enabled: set_state("debug_enabled", bool(enabled)),
        lambda count: set_state("drop_count", int(count)),
        lambda: set_state("clear_window_refs"),
        center_window,
    )
    kwargs = {"window_title": window_title}
    if log_error is not None:
        kwargs["log_error"] = log_error
    try:
        dialog_parameters = inspect.signature(open_dialog).parameters.values()
        supports_optional_popup_callbacks = any(
            parameter.kind == inspect.Parameter.VAR_KEYWORD
            for parameter in dialog_parameters
        ) or all(
            name in inspect.signature(open_dialog).parameters
            for name in ("get_dial_popup", "set_dial_popup")
        )
    except (TypeError, ValueError):
        supports_optional_popup_callbacks = True
    if supports_optional_popup_callbacks:
        kwargs["get_dial_popup"] = get_dial_popup
        kwargs["set_dial_popup"] = set_dial_popup
    window, text = open_dialog(*args, **kwargs)
    set_state("window_refs", window, text)
    return window, text
