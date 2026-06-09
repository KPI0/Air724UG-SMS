import queue

from sms_ui.serial_debug_app_runtime import open_serial_debug_window_runtime


def get_serial_debug_state_namespace_runtime(namespace, name):
    if name == "window_refs":
        return namespace["serial_debug_win"], namespace["serial_debug_text"]
    if name == "debug_enabled":
        return namespace["SERIAL_DEBUG_ENABLED"]
    if name == "drop_count":
        return namespace["serial_debug_drop_count"]
    raise KeyError(name)


def set_serial_debug_state_namespace_runtime(namespace, name, *values):
    if name == "window_refs":
        namespace.__setitem__("serial_debug_win", values[0])
        namespace.__setitem__("serial_debug_text", values[1])
    elif name == "clear_window_refs":
        namespace.__setitem__("serial_debug_win", None)
        namespace.__setitem__("serial_debug_text", None)
    elif name == "debug_enabled":
        namespace.__setitem__("SERIAL_DEBUG_ENABLED", bool(values[0]))
    elif name == "drop_count":
        namespace.__setitem__("serial_debug_drop_count", int(values[0]))
    elif name == "current_dial_num":
        namespace.__setitem__("current_dial_num", values[0])
    else:
        raise KeyError(name)


def open_serial_debug_window_namespace_runtime(
    namespace,
    *,
    open_window_runtime=open_serial_debug_window_runtime,
):
    return open_window_runtime(
        namespace["root"],
        get_state=lambda name: get_serial_debug_state_namespace_runtime(namespace, name),
        set_state=lambda name, *values: set_serial_debug_state_namespace_runtime(namespace, name, *values),
        serial_queue=namespace["serial_debug_queue"],
        serial_lock=namespace["serial_lock"],
        get_serial_obj=lambda: namespace["serial_obj"],
        push_serial_debug=namespace["_push_serial_debug"],
        port_ui=namespace["port_ui"],
        set_status=namespace["set_status"],
        format_connected_status=namespace["format_connected_status"],
        get_port=lambda: namespace["PORT"],
        center_window=namespace["center_window"],
    )


def push_serial_debug_namespace_runtime(namespace, raw_line):
    if not namespace["SERIAL_DEBUG_ENABLED"]:
        return None
    if raw_line is None:
        return None
    if not str(raw_line).strip():
        return None

    try:
        namespace["serial_debug_queue"].put_nowait(raw_line)
        return "queued"
    except queue.Full:
        namespace.__setitem__(
            "serial_debug_drop_count",
            namespace["serial_debug_drop_count"] + 1,
        )
        return "full"
