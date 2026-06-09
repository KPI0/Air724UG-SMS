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
    log_error=None,
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
    if log_error is None:
        window, text = open_dialog(*args)
    else:
        window, text = open_dialog(*args, log_error=log_error)
    set_state("window_refs", window, text)
    return window, text
