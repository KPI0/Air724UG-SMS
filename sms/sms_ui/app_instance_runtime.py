import hashlib
import os
import sys

from sms_core.windows_runtime import (
    acquire_mutex_with_error,
    close_windows_handle,
    focus_existing_window,
    is_existing_instance_error,
    show_message_box,
)


APP_MUTEX_NAME = "Air724UG_SMS_Monitor_Mutex_V3"
APP_INSTANCE_MUTEX_PREFIX = "Air724UG_SMS_Monitor_Instance_V1"
MAX_INSTANCE_NUMBER = 9999


def app_dir_mutex_name(app_dir, *, prefix=APP_MUTEX_NAME):
    if not app_dir:
        return prefix
    normalized = os.path.normcase(os.path.abspath(str(app_dir)))
    digest = hashlib.sha256(normalized.encode("utf-8", "surrogatepass")).hexdigest()[:16]
    return f"{prefix}_{digest}"


def instance_mutex_name(app_dir, instance_number):
    number = max(1, int(instance_number))
    base_name = app_dir_mutex_name(app_dir, prefix=APP_INSTANCE_MUTEX_PREFIX)
    return f"{base_name}_{number}"


def format_instance_window_title(base_title, instance_number):
    number = max(1, int(instance_number or 1))
    return str(base_title) if number == 1 else f"{base_title} {number}"


def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def claim_instance_number_app_runtime(
    *,
    app_dir=None,
    max_instance_number=MAX_INSTANCE_NUMBER,
    acquire_mutex=acquire_mutex_with_error,
    close_handle=close_windows_handle,
    existing_error=is_existing_instance_error,
    log_error=None,
):
    """Claim the smallest available per-app instance number.

    The returned mutex handle must remain open for the process lifetime so
    other processes cannot claim the same number.
    """
    try:
        limit = max(1, int(max_instance_number))
    except (TypeError, ValueError):
        limit = MAX_INSTANCE_NUMBER

    for number in range(1, limit + 1):
        mutex, last_error = acquire_mutex(instance_mutex_name(app_dir, number))
        if existing_error(last_error):
            if mutex:
                try:
                    close_handle(mutex)
                except Exception as exc:
                    _safe_log(log_error, f"Close occupied instance mutex failed: {exc!r}")
            continue

        if mutex:
            return number, mutex

        _safe_log(
            log_error,
            f"Claim instance number {number} failed with error code {last_error}",
        )
        break

    _safe_log(log_error, "No instance number mutex could be claimed; using unnumbered title")
    return 1, None


def is_instance_number_active_app_runtime(
    *,
    app_dir,
    instance_number,
    acquire_mutex=acquire_mutex_with_error,
    close_handle=close_windows_handle,
    existing_error=is_existing_instance_error,
):
    """Probe a numbered instance mutex without retaining the probe handle."""
    handle, last_error = acquire_mutex(instance_mutex_name(app_dir, instance_number))
    try:
        return bool(existing_error(last_error))
    finally:
        if handle:
            close_handle(handle)


def check_single_instance_app_runtime(
    *,
    allow_multi_instance,
    window_title,
    app_dir=None,
    mutex_name=None,
    acquire_mutex=acquire_mutex_with_error,
    close_handle=close_windows_handle,
    existing_error=is_existing_instance_error,
    focus_window=focus_existing_window,
    show_message=show_message_box,
    exit_process=sys.exit,
    log_error=None,
):
    if mutex_name is None:
        mutex_name = app_dir_mutex_name(app_dir)

    app_mutex, last_error = acquire_mutex(mutex_name)

    if allow_multi_instance:
        if app_mutex:
            return app_mutex
        _safe_log(
            log_error,
            f"Create multi-instance presence mutex failed with error code {last_error}",
        )
        show_message(
            "启动失败",
            f"无法创建实例状态锁，程序为避免后续单实例保护失效将退出。\n错误码：{last_error}",
            0x10,
        )
        exit_process(1)
        return None

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
