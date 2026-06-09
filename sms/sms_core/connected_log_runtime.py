import threading
import time

from sms_core.status_text import format_connected_status as default_format_connected_status


PROTECTED_STATUS_WORDS = ("响铃", "通话", "呼叫")


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
):
    try:
        sleep(delay)
        reset_auto_connect_state()
        clear_serial_error_repeat_state()
        system_ui(f"🔌 串口已连接：{port} @ {baud}")

        def update_status():
            try:
                update = connected_status_update(get_status(), port, format_connected_status)
                if update is not None:
                    text, color = update
                    set_status(text, color)
            except Exception:
                pass

        def schedule_update():
            delay_ms = startup_ui_delay_ms(app_start_mono, start_ui_delay, monotonic)
            try:
                root_after(delay_ms, update_status)
            except Exception:
                update_status()

        ui_post(schedule_update)
    except Exception:
        pass


def start_delayed_connected_log_runtime(
    port,
    baud,
    *,
    thread_factory=threading.Thread,
    **kwargs,
):
    thread = thread_factory(
        target=lambda: run_delayed_connected_log_runtime(port, baud, **kwargs),
        daemon=True,
    )
    thread.start()
    return thread
