from dataclasses import dataclass

from sms_core.cloud_control_settings import (
    CloudControlFormValues,
    CloudControlSettings,
    cloud_control_form_values,
    cloud_control_save_kwargs,
    cloud_control_state,
    normalize_reconnect_interval,
    read_cloud_control_settings,
    update_cloud_control_settings,
    write_cloud_control_settings,
)
from sms_core.cloud_protocol import normalize_cloud_ws_url


@dataclass(frozen=True)
class CloudStartValidation:
    ok: bool
    url: str = ""
    status_text: str = ""
    status_color: str = "#666666"
    warning_title: str = ""
    warning_message: str = ""


def cloud_restarting_status():
    return "🌐 正在重启"


def cloud_stopped_status(enabled):
    return "🌐 已关闭" if not enabled else "🌐 已断开"


def next_cloud_backoff(current_backoff, factor=1.5, max_backoff=60.0):
    try:
        current = max(1.0, float(current_backoff))
    except Exception:
        current = 1.0
    return min(float(max_backoff), current * float(factor))


def cloud_backoff_sleep_ticks(current_backoff, ticks_per_second=10):
    try:
        seconds = max(1, int(current_backoff))
    except Exception:
        seconds = 1
    return seconds * int(ticks_per_second)


def cloud_auto_upload_action(was_public, auto_upload, connected, has_loop, has_conn):
    if auto_upload and connected and has_loop and has_conn:
        return "register"
    if not auto_upload and was_public:
        return "unregister"
    return ""


def cloud_start_thread_action(thread_alive, stop_requested):
    if thread_alive and stop_requested:
        return "restarting"
    if thread_alive:
        return "already_running"
    return "start"


def cloud_restart_attempt_action(restart_seq, current_seq, tk_is_alive, thread_alive, stop_requested):
    if restart_seq != current_seq or not tk_is_alive:
        return "cancel"
    if thread_alive and stop_requested:
        return "wait"
    return "start"


def start_cloud_control_runtime(
    *,
    websockets_available,
    url,
    device_secret,
    reconnect_interval,
    show_errors,
    validate_start,
    set_cloud_status,
    log_missing_dependency,
    show_warning,
    runtime_imei,
    request_device_imei,
    lock,
    get_thread,
    set_thread,
    stop_event,
    thread_factory,
    thread_target,
    start_thread_action=cloud_start_thread_action,
    restarting_status=cloud_restarting_status,
):
    validation = validate_start(websockets_available, url, device_secret)
    if not validation.ok:
        set_cloud_status(validation.status_text, validation.status_color)
        if not websockets_available:
            log_missing_dependency()
        if show_errors:
            show_warning(validation.warning_title, validation.warning_message)
        return False

    if not runtime_imei():
        request_device_imei()

    with lock:
        thread = get_thread()
        action = start_thread_action(
            thread is not None and thread.is_alive(),
            stop_event.is_set(),
        )
        if action == "restarting":
            set_cloud_status(restarting_status(), "#b26a00")
            return False
        if action == "already_running":
            return True

        stop_event.clear()
        thread = thread_factory(
            target=thread_target,
            args=(validation.url, reconnect_interval),
            daemon=True,
        )
        set_thread(thread)
        thread.start()

    return True


def stop_cloud_control_runtime(
    *,
    update_status,
    enabled,
    stop_event,
    set_connected,
    set_authorized,
    reset_serial_log_state,
    get_loop,
    get_ws,
    schedule_unregister_then_close,
    set_ws,
    set_cloud_status,
    run_coroutine_threadsafe,
    stopped_status=cloud_stopped_status,
):
    stop_event.set()
    set_connected(False)
    set_authorized(False)
    reset_serial_log_state()

    try:
        loop = get_loop()
        ws = get_ws()
        if loop is not None and loop.is_running() and ws is not None:
            run_coroutine_threadsafe(schedule_unregister_then_close(ws), loop)
    except Exception:
        pass

    set_ws(None)

    if update_status:
        set_cloud_status(stopped_status(enabled), "#666666")


def restart_cloud_control_runtime(
    *,
    show_errors,
    lock,
    increment_restart_seq,
    get_restart_seq,
    get_thread,
    stop_control,
    tk_alive,
    stop_event,
    set_cloud_status,
    schedule_after,
    ui_post,
    start_control,
    thread_factory,
    restart_attempt_action=cloud_restart_attempt_action,
    restarting_status=cloud_restarting_status,
):
    with lock:
        restart_seq = increment_restart_seq()
        old_thread = get_thread()

    stop_control(update_status=False)

    def wait_and_start():
        try:
            if old_thread is not None and old_thread.is_alive():
                old_thread.join(timeout=2.0)
        except Exception:
            pass

        def try_start():
            try:
                if restart_seq != get_restart_seq() or not tk_alive():
                    return

                with lock:
                    thread = get_thread()
                    action = restart_attempt_action(
                        restart_seq,
                        get_restart_seq(),
                        True,
                        thread is not None and thread.is_alive(),
                        stop_event.is_set(),
                    )

                if action == "wait":
                    set_cloud_status(restarting_status(), "#b26a00")
                    schedule_after(500, try_start)
                    return
                if action == "cancel":
                    return
            except Exception:
                return

            start_control(show_errors=show_errors)

        ui_post(try_start)

    thread = thread_factory(target=wait_and_start, daemon=True)
    thread.start()
    return thread


def validate_cloud_start(websockets_available, url, device_secret):
    if not websockets_available:
        return CloudStartValidation(
            ok=False,
            status_text="🌐 缺少依赖",
            status_color="#cc0000",
            warning_title="云端控制",
            warning_message="当前 Python 环境缺少 websockets 库，无法启动云端控制。",
        )

    normalized_url = normalize_cloud_ws_url(url)
    if not normalized_url:
        return CloudStartValidation(
            ok=False,
            status_text="🌐 未配置",
            status_color="#cc0000",
            warning_title="云端控制",
            warning_message="请先填写 WebSocket 地址。",
        )

    if not (normalized_url.startswith("ws://") or normalized_url.startswith("wss://")):
        return CloudStartValidation(
            ok=False,
            url=normalized_url,
            status_text="🌐 地址错误",
            status_color="#cc0000",
            warning_title="云端控制",
            warning_message="WebSocket 地址必须以 ws:// 或 wss:// 开头。",
        )

    if not str(device_secret or "").strip():
        return CloudStartValidation(
            ok=False,
            url=normalized_url,
            status_text="🌐 密码未配置",
            status_color="#cc0000",
            warning_title="云端控制",
            warning_message="请先设置云端控制密码。",
        )

    return CloudStartValidation(ok=True, url=normalized_url)
