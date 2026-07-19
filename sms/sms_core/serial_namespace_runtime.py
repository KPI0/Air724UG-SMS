from sms_core.connected_log_runtime import start_delayed_connected_log_runtime
from sms_core.serial_ports import choose_luat_modem_port
from sms_core.serial_reconnect import is_serial_port_gone_error
from sms_core.serial_reconnect_app_runtime import try_rebind_manual_port_runtime
from sms_core.serial_startup_runtime import (
    open_and_initialize_serial_runtime,
    resolve_serial_target_port_runtime,
)
from sms_core.status_text import format_connected_status


def try_rebind_manual_port_namespace_runtime(
    namespace,
    reason="",
    *,
    rebind_runtime=try_rebind_manual_port_runtime,
):
    return rebind_runtime(
        reason,
        mode=namespace["MODE"],
        current_port=namespace["PORT"],
        baud=namespace["BAUD"],
        find_luat_best_port=namespace["find_luat_best_port"],
        list_ports=namespace["list_ports"].comports,
        choose_candidate=namespace["choose_manual_rebind_candidate"],
        config=namespace["config"],
        save_config=namespace["safe_save_config"],
        set_port=lambda port: namespace.__setitem__("PORT", port),
        system_ui=namespace["system_ui"],
        set_status=namespace["set_status"],
        wake_serial=namespace["serial_wakeup_event"].set,
        reset_rebind_hint=namespace["_rebind_hint_notice"].reset,
        hint_formatter=namespace["manual_rebind_hint"],
    )


def resolve_serial_target_port_namespace_runtime(
    namespace,
    *,
    resolve_runtime=resolve_serial_target_port_runtime,
):
    return resolve_runtime(
        mode=namespace["MODE"],
        current_port=namespace["PORT"],
        reconnect_interval=namespace["RECONNECT_INTERVAL"],
        find_luat_best_port=namespace["find_luat_best_port"],
        list_ports=namespace["list_ports"].comports,
        is_port_locked=namespace["is_port_locked_by_other"],
        auto_connect_ui=namespace["auto_connect_ui"],
        set_status=namespace["set_status"],
        wakeup_wait=namespace["serial_wakeup_event"].wait,
        wakeup_clear=namespace["serial_wakeup_event"].clear,
    )


def open_and_initialize_serial_namespace_runtime(
    namespace,
    target_port,
    *,
    open_runtime=open_and_initialize_serial_runtime,
):
    return open_runtime(
        target_port=target_port,
        baud=namespace["BAUD"],
        mode=namespace["MODE"],
        serial_lock=namespace["serial_lock"],
        open_serial=lambda port, baud: namespace["serial"].Serial(port, baud, timeout=0.3, write_timeout=0.5),
        set_serial_obj=lambda value: namespace.__setitem__("serial_obj", value),
        set_port=lambda value: namespace.__setitem__("PORT", value),
        lock_port_mutex=namespace["lock_port_mutex"],
        set_cloud_imei_query_deadline=lambda value: namespace.__setitem__("cloud_imei_query_deadline", value),
        serial_error_ui=namespace["serial_error_ui"],
        set_status=namespace["set_status"],
    )


def schedule_delayed_connected_log_namespace_runtime(
    namespace,
    port,
    baud,
    *,
    delay=2,
    start_runtime=start_delayed_connected_log_runtime,
):
    generation = int(namespace.get("serial_connection_generation", 0)) + 1
    namespace.__setitem__("serial_connection_generation", generation)
    connection = namespace.get("serial_obj")

    def connection_is_current():
        if generation != namespace.get("serial_connection_generation"):
            return False
        lock = namespace.get("serial_lock")
        try:
            with lock:
                current = namespace.get("serial_obj")
        except Exception:
            current = namespace.get("serial_obj")
        if generation != namespace.get("serial_connection_generation"):
            return False
        if current is not connection or current is None:
            return False
        try:
            return bool(getattr(current, "is_open", True))
        except Exception:
            return False

    def is_stopping():
        if not bool(namespace.get("serial_running", True)):
            return True
        for name in ("serial_stop_event", "TK_SHUTDOWN"):
            event = namespace.get(name)
            try:
                if event is not None and event.is_set():
                    return True
            except Exception:
                pass
        return False

    return start_runtime(
        port,
        baud,
        delay=delay,
        reset_auto_connect_state=namespace["_auto_connect_notice"].reset,
        clear_serial_error_repeat_state=namespace["_serial_error_notice"].clear,
        system_ui=namespace["system_ui"],
        ui_post=namespace["ui_post"],
        root_after=namespace["root"].after,
        get_status=namespace["status_var"].get,
        set_status=namespace["set_status"],
        app_start_mono=namespace["APP_START_MONO"],
        start_ui_delay=namespace["START_UI_DELAY"],
        format_connected_status=format_connected_status,
        connection_is_current=connection_is_current,
        is_stopping=is_stopping,
        log_error=namespace.get("log_file_only"),
    )


def try_manual_rebind_after_error_namespace_runtime(
    namespace,
    error,
    *,
    hint_message,
    is_gone_error=is_serial_port_gone_error,
):
    if namespace["MODE"] != "Manual" or not is_gone_error(error):
        return False
    namespace["rebind_hint_ui"](hint_message)
    return namespace["try_rebind_manual_port"]("端口号变化或设备插拔")


def scan_com_ports_all_namespace_runtime(namespace):
    return [port.device for port in namespace["list_ports"].comports()]


def find_luat_best_port_namespace_runtime(
    namespace,
    *,
    choose_port=choose_luat_modem_port,
):
    return choose_port(
        namespace["list_ports"].comports(),
        remembered_port=namespace["PORT"],
        is_locked=namespace["is_port_locked_by_other"],
    )
