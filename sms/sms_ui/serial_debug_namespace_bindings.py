from sms_core.namespace_binding import make_namespace_runtime_binder
from sms_ui.serial_debug_namespace_runtime import (
    get_serial_debug_state_namespace_runtime,
    open_serial_debug_window_namespace_runtime,
    set_serial_debug_state_namespace_runtime,
)


def install_serial_debug_namespace_bindings(namespace):
    bind = make_namespace_runtime_binder(namespace, globals())

    namespace.update({
        "get_serial_debug_state": bind("get_serial_debug_state_namespace_runtime"),
        "set_serial_debug_state": bind("set_serial_debug_state_namespace_runtime"),
        "open_serial_debug_window": bind("open_serial_debug_window_namespace_runtime"),
    })
    return namespace
