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
    """
    Set serial debug state by name.

    Refactored from nested if-elif chain to dispatch table pattern
    for better maintainability and reduced cyclomatic complexity.
    """
    def _require_values(count, actual):
        if len(actual) < count:
            raise ValueError(f"{name} requires {count} value(s), got {len(actual)}")

    # Dispatch table: cleaner than 6-level if-elif nesting
    handlers = {
        "window_refs": lambda: (
            _require_values(2, values),
            namespace.__setitem__("serial_debug_win", values[0]),
            namespace.__setitem__("serial_debug_text", values[1]),
        ),
        "clear_window_refs": lambda: (
            namespace.__setitem__("serial_debug_win", None),
            namespace.__setitem__("serial_debug_text", None),
        ),
        "debug_enabled": lambda: (
            _require_values(1, values),
            namespace.__setitem__("SERIAL_DEBUG_ENABLED", bool(values[0])),
        ),
        "drop_count": lambda: (
            _require_values(1, values),
            namespace.__setitem__("serial_debug_drop_count", int(values[0])),
        ),
        "current_dial_num": lambda: (
            _require_values(1, values),
            namespace.__setitem__("current_dial_num", values[0]),
        ),
    }

    handler = handlers.get(name)
    if handler is None:
        raise KeyError(name)

    handler()


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
        log_error=namespace.get("log_file_only"),
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
