from sms_core.serial_app_runtime import run_serial_reader_namespace_runtime
from sms_core.serial_io_runtime import read_serial_line_safely_runtime, send_call_hangup_runtime
from sms_core.serial_namespace_runtime import (
    find_luat_best_port_namespace_runtime,
    open_and_initialize_serial_namespace_runtime,
    resolve_serial_target_port_namespace_runtime,
    scan_com_ports_all_namespace_runtime,
    schedule_delayed_connected_log_namespace_runtime,
    try_manual_rebind_after_error_namespace_runtime,
    try_rebind_manual_port_namespace_runtime,
)
from sms_core.serial_ports import choose_manual_rebind_candidate, manual_rebind_hint
from sms_core.serial_reconnect import apply_serial_disconnect_effects
from sms_ui.call_popup_namespace_runtime import (
    close_call_popup_namespace_runtime,
    get_serial_call_state_namespace_runtime,
    set_call_popup_namespace_runtime,
    set_serial_call_state_namespace_runtime,
    show_call_popup_namespace_runtime,
)
from sms_ui.repeat_notice_runtime import emit_repeat_notice
from sms_ui.serial_debug_namespace_runtime import push_serial_debug_namespace_runtime


def install_serial_namespace_bindings(namespace):
    namespace.setdefault("choose_manual_rebind_candidate", choose_manual_rebind_candidate)
    namespace.setdefault("manual_rebind_hint", manual_rebind_hint)

    def scan_com_ports_all():
        return scan_com_ports_all_namespace_runtime(namespace)

    def find_luat_best_port():
        return find_luat_best_port_namespace_runtime(namespace)

    def push_serial_debug(raw_line):
        return push_serial_debug_namespace_runtime(namespace, raw_line)

    def try_rebind_manual_port(reason=""):
        return try_rebind_manual_port_namespace_runtime(namespace, reason)

    def rebind_hint_ui(msg):
        return emit_repeat_notice(namespace["_rebind_hint_notice"], msg, namespace["system_ui"])

    def serial_error_ui(msg, repeat_key=""):
        return emit_repeat_notice(
            namespace["_serial_error_notice"],
            msg,
            namespace["system_ui"],
            repeat_key=repeat_key,
        )

    def set_call_popup(window):
        return set_call_popup_namespace_runtime(namespace, window)

    def close_call_popup():
        return close_call_popup_namespace_runtime(namespace)

    def show_call_popup(caller_num):
        return show_call_popup_namespace_runtime(namespace, caller_num)

    def resolve_serial_target_port():
        return resolve_serial_target_port_namespace_runtime(namespace)

    def open_and_initialize_serial(target_port):
        return open_and_initialize_serial_namespace_runtime(namespace, target_port)

    def schedule_delayed_connected_log(port, baud, delay=2):
        return schedule_delayed_connected_log_namespace_runtime(namespace, port, baud, delay=delay)

    def read_serial_line_safely():
        return read_serial_line_safely_runtime(
            namespace["serial_lock"],
            lambda: namespace["serial_obj"],
            namespace["serial"].SerialException,
        )

    def try_manual_rebind_after_error(error):
        return try_manual_rebind_after_error_namespace_runtime(
            namespace,
            error,
            hint_message="🧠 Manual：检测到疑似拔插/端口变化，尝试自动重绑…",
        )

    def send_call_hangup_command():
        return send_call_hangup_runtime(
            namespace["serial_lock"],
            lambda: namespace["serial_obj"],
            namespace["write_serial_command_result"],
        )

    def get_serial_call_state():
        return get_serial_call_state_namespace_runtime(namespace)

    def set_serial_call_state(next_ring_timeout, next_dial_num):
        return set_serial_call_state_namespace_runtime(namespace, next_ring_timeout, next_dial_num)

    def set_serial_log_prefix(value):
        namespace.__setitem__("LOG_PREFIX", value)

    def read_serial():
        return run_serial_reader_namespace_runtime(
            namespace,
            parse_callback_head=namespace["_parse_cloud_sms_callback_head"],
            apply_disconnect_effects=apply_serial_disconnect_effects,
        )

    namespace.update({
        "scan_com_ports_all": scan_com_ports_all,
        "find_luat_best_port": find_luat_best_port,
        "_push_serial_debug": push_serial_debug,
        "try_rebind_manual_port": try_rebind_manual_port,
        "rebind_hint_ui": rebind_hint_ui,
        "serial_error_ui": serial_error_ui,
        "set_call_popup": set_call_popup,
        "close_call_popup": close_call_popup,
        "show_call_popup": show_call_popup,
        "resolve_serial_target_port": resolve_serial_target_port,
        "open_and_initialize_serial": open_and_initialize_serial,
        "schedule_delayed_connected_log": schedule_delayed_connected_log,
        "read_serial_line_safely": read_serial_line_safely,
        "try_manual_rebind_after_error": try_manual_rebind_after_error,
        "send_call_hangup_command": send_call_hangup_command,
        "get_serial_call_state": get_serial_call_state,
        "set_serial_call_state": set_serial_call_state,
        "set_serial_log_prefix": set_serial_log_prefix,
        "read_serial": read_serial,
    })
    return namespace
