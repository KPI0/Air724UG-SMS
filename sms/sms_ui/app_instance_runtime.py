import sys

from sms_core.windows_runtime import (
    acquire_mutex_with_error,
    close_windows_handle,
    focus_existing_window,
    is_existing_instance_error,
    show_message_box,
)


APP_MUTEX_NAME = "Air724UG_SMS_Monitor_Mutex_V3"


def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def check_single_instance_app_runtime(
    *,
    allow_multi_instance,
    window_title,
    mutex_name=APP_MUTEX_NAME,
    acquire_mutex=acquire_mutex_with_error,
    close_handle=close_windows_handle,
    existing_error=is_existing_instance_error,
    focus_window=focus_existing_window,
    show_message=show_message_box,
    exit_process=sys.exit,
    log_error=None,
):
    if allow_multi_instance:
        return None

    app_mutex, last_error = acquire_mutex(mutex_name)

    if existing_error(last_error):
        if app_mutex:
            try:
                close_handle(app_mutex)
            except Exception as exc:
                _safe_log(log_error, f"Close existing instance mutex failed: {exc!r}")
        if not focus_window(window_title):
            show_message("提示", "程序已经在运行中，请在右下角托盘查看。", 0x30)
        exit_process(0)
        return None

    if not app_mutex:
        show_message(
            "启动失败",
            f"无法创建单实例锁，程序为避免多开将退出。\n错误码：{last_error}",
            0x10,
        )
        exit_process(1)
        return None

    return app_mutex
