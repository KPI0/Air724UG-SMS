from sms_core.serial_app_runtime import run_serial_reader_namespace_runtime
from sms_core.serial_io_runtime import read_serial_line_safely_runtime, send_call_hangup_runtime
from sms_core.namespace_binding import make_namespace_runtime_binder
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
    close_missed_call_popup_namespace_runtime,
    finish_incoming_call_session_namespace_runtime,
    get_serial_call_state_namespace_runtime,
    mark_incoming_call_handled_namespace_runtime,
    reset_incoming_call_session_namespace_runtime,
    set_call_popup_namespace_runtime,
    set_missed_call_popup_namespace_runtime,
    set_serial_call_state_namespace_runtime,
    show_missed_call_popup_namespace_runtime,
    show_call_popup_namespace_runtime,
    start_incoming_call_session_namespace_runtime,
)
from sms_ui.repeat_notice_runtime import emit_repeat_notice
from sms_ui.serial_debug_namespace_runtime import push_serial_debug_namespace_runtime
from sms_ui.ui_log_runtime import system_log_prefix_runtime


def install_serial_namespace_bindings(namespace):
    bind = make_namespace_runtime_binder(namespace, globals())
    namespace.setdefault("choose_manual_rebind_candidate", choose_manual_rebind_candidate)
    namespace.setdefault("manual_rebind_hint", manual_rebind_hint)

    def rebind_hint_ui(msg):
        return emit_repeat_notice(namespace["_rebind_hint_notice"], msg, namespace["system_ui"])

    def serial_error_ui(msg, repeat_key=""):
        return emit_repeat_notice(
            namespace["_serial_error_notice"],
            msg,
            namespace["system_ui"],
            repeat_key=repeat_key,
        )

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

    def set_serial_log_prefix(value):
        if value == "system":
            value = system_log_prefix_runtime(namespace.get("APP_INSTANCE_NUMBER", 1))
        namespace.__setitem__("LOG_PREFIX", value)

    def read_serial():
        return run_serial_reader_namespace_runtime(
            namespace,
            parse_callback_head=namespace["_parse_cloud_sms_callback_head"],
            apply_disconnect_effects=apply_serial_disconnect_effects,
        )

    def notify_cloud_channel_status(serial_connected):
        callback = namespace.get("_notify_cloud_channel_status")
        if callback is None:
            return False
        try:
            return callback(serial_connected)
        except Exception as exc:
            try:
                namespace["log_file_only"](
                    f"Cloud channel status notification failed: {type(exc).__name__}"
                )
            except Exception:
                pass
            return False

    namespace.update({
        "scan_com_ports_all": bind("scan_com_ports_all_namespace_runtime"),
        "find_luat_best_port": bind("find_luat_best_port_namespace_runtime"),
        "_push_serial_debug": bind("push_serial_debug_namespace_runtime"),
        "try_rebind_manual_port": bind("try_rebind_manual_port_namespace_runtime"),
        "rebind_hint_ui": rebind_hint_ui,
        "serial_error_ui": serial_error_ui,
        "set_call_popup": bind("set_call_popup_namespace_runtime"),
        "close_call_popup": bind("close_call_popup_namespace_runtime"),
        "close_missed_call_popup": bind("close_missed_call_popup_namespace_runtime"),
        "show_call_popup": bind("show_call_popup_namespace_runtime"),
        "set_missed_call_popup": bind("set_missed_call_popup_namespace_runtime"),
        "show_missed_call_popup": bind("show_missed_call_popup_namespace_runtime"),
        "start_incoming_call_session": bind("start_incoming_call_session_namespace_runtime"),
        "mark_incoming_call_handled": bind("mark_incoming_call_handled_namespace_runtime"),
        "finish_incoming_call_session": bind("finish_incoming_call_session_namespace_runtime"),
        "reset_incoming_call_session": bind("reset_incoming_call_session_namespace_runtime"),
        "resolve_serial_target_port": bind("resolve_serial_target_port_namespace_runtime"),
        "open_and_initialize_serial": bind("open_and_initialize_serial_namespace_runtime"),
        "schedule_delayed_connected_log": bind("schedule_delayed_connected_log_namespace_runtime"),
        "read_serial_line_safely": read_serial_line_safely,
        "try_manual_rebind_after_error": try_manual_rebind_after_error,
        "send_call_hangup_command": send_call_hangup_command,
        "get_serial_call_state": bind("get_serial_call_state_namespace_runtime"),
        "set_serial_call_state": bind("set_serial_call_state_namespace_runtime"),
        "set_serial_log_prefix": set_serial_log_prefix,
        "read_serial": read_serial,
        "notify_cloud_channel_status": notify_cloud_channel_status,
    })
    return namespace
