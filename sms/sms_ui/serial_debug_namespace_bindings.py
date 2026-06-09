from sms_ui.serial_debug_namespace_runtime import (
    get_serial_debug_state_namespace_runtime,
    open_serial_debug_window_namespace_runtime,
    set_serial_debug_state_namespace_runtime,
)


def install_serial_debug_namespace_bindings(namespace):
    def get_serial_debug_state(name):
        return get_serial_debug_state_namespace_runtime(namespace, name)

    def set_serial_debug_state(name, *values):
        return set_serial_debug_state_namespace_runtime(namespace, name, *values)

    def open_serial_debug_window():
        return open_serial_debug_window_namespace_runtime(namespace)

    namespace.update({
        "get_serial_debug_state": get_serial_debug_state,
        "set_serial_debug_state": set_serial_debug_state,
        "open_serial_debug_window": open_serial_debug_window,
    })
    return namespace
