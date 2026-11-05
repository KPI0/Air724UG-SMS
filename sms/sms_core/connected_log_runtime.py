import threading
import time

from sms_core.status_text import format_connected_status as default_format_connected_status
from sms_core.serial_sender import serial_initialization_pending, serial_sms_recovery_error
from sms_core.threading_runtime import start_daemon_thread


PROTECTED_STATUS_WORDS = ("响铃", "通话", "呼叫")


def can_restore_connected_status(namespace, serial, generation):
    """Restore idle UI only for the originating, ready serial connection."""
    try:
        if (serial is None or namespace.get("serial_obj") is not serial
                or namespace.get("serial_connection_generation", 0) != generation
                or namespace.get("serial_running", True) is False
                or namespace.get("is_exiting", False)
                or not serial.is_open
                or serial_initialization_pending(serial)
                or serial_sms_recovery_error(serial)):
            return False
        for name in ("serial_stop_event", "TK_SHUTDOWN"):
            event = namespace.get(name)
            if event is not None and event.is_set():
                return False
        initialization = namespace.get("_serial_initialization_session")
        if (getattr(initialization, "serial", None) is serial
                and getattr(initialization, "has_command_errors", False)):
            return False
        if namespace.get("current_dial_num") or namespace.get("_call_command_pending_incoming"):
            return False
        snapshot = getattr(namespace.get("INCOMING_CALL_SESSION"), "snapshot", None)
        if callable(snapshot) and getattr(snapshot(), "session_token", None) is not None:
            return False
        return True
    except Exception:
        return False


def _connected_task_is_current(connection_is_current, is_stopping):
    try:
        if is_stopping():
            return False
        return bool(connection_is_current())
    except Exception:
        return False


def startup_ui_delay_ms(app_start_mono, start_ui_delay, monotonic=time.monotonic):
    try:
        elapsed = monotonic() - app_start_mono
        return int(max(0.0, start_ui_delay - elapsed) * 1000)
    except Exception:
        return 0


def connected_status_update(current_status, port, format_connected_status=default_format_connected_status):
    current_status = str(current_status or "")
    if any(word in current_status for word in PROTECTED_STATUS_WORDS):
        return None
    return format_connected_status(port), "green"


def run_delayed_connected_log_runtime(
    port,
    baud,
    *,
    delay=2,
    sleep=time.sleep,
    reset_auto_connect_state,
    clear_serial_error_repeat_state,
    system_ui,
    ui_post,
    root_after,
    get_status,
    set_status,
    app_start_mono,
    start_ui_delay,
    monotonic=time.monotonic,
    format_connected_status=default_format_connected_status,
    connection_is_current=lambda: True,
    is_stopping=lambda: False,
):
    try:
        sleep(delay)
        if not _connected_task_is_current(connection_is_current, is_stopping):
            return "stale"
        reset_auto_connect_state()
        clear_serial_error_repeat_state()
        if not _connected_task_is_current(connection_is_current, is_stopping):
            return "stale"
        system_ui(f"🔌 串口已连接：{port} @ {baud}")

        def update_status():
            try:
                if not _connected_task_is_current(connection_is_current, is_stopping):
                    return
                update = connected_status_update(get_status(), port, format_connected_status)
                if update is not None:
                    text, color = update
                    set_status(text, color)
            except Exception:
                pass

        def schedule_update():
            if not _connected_task_is_current(connection_is_current, is_stopping):
                return
            delay_ms = startup_ui_delay_ms(app_start_mono, start_ui_delay, monotonic)
            try:
                root_after(delay_ms, update_status)
            except Exception:
                update_status()

        ui_post(schedule_update)
        return "scheduled"
    except Exception:
        return "error"


def start_delayed_connected_log_runtime(
    port,
    baud,
    *,
    thread_factory=threading.Thread,
    log_error=None,
    **kwargs,
):
    return start_daemon_thread(
        "delayed_connected_log",
        lambda: run_delayed_connected_log_runtime(port, baud, **kwargs),
        log_error=log_error,
        thread_factory=thread_factory,
    )
