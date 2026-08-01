import os

from sms_core.windows_runtime import acquire_mutex_with_error
from sms_ui.app_autostart_runtime import set_autostart_runtime
from sms_ui.app_instance_runtime import app_dir_mutex_name
from sms_ui.app_restart_runtime import restart_software_app_runtime
from sms_ui.call_popup_namespace_runtime import close_phone_popups_namespace_runtime
from sms_ui.app_shutdown_runtime import cleanup_and_exit_app_runtime
from sms_ui.settings_runtime import (
    toggle_call_popup_runtime,
    toggle_multi_instance_runtime,
    toggle_popup_runtime,
    toggle_voice_broadcast_runtime,
)


def _registered_worker_threads(namespace):
    threads = []
    for registry_name in (
        "UPDATE_THREAD_REGISTRY",
        "SERIAL_COMMAND_THREAD_REGISTRY",
        "SMS_SEND_THREAD_REGISTRY",
    ):
        registry = namespace.get(registry_name)
        if registry is None:
            continue
        try:
            threads.extend(registry.snapshot())
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if log_error is not None:
                try:
                    log_error(f"Snapshot {registry_name} failed: {exc!r}")
                except Exception:
                    pass
            raise RuntimeError(f"Snapshot {registry_name} failed") from exc
    return tuple(threads)


def _shutdown_worker_threads(namespace):
    return (
        namespace.get("serial_thread"),
        namespace.get("TTS_THREAD"),
        namespace.get("cloud_ws_thread"),
        namespace.get("tray_thread"),
    ) + _registered_worker_threads(namespace)


def set_autostart_namespace_runtime(namespace, enable, *, set_autostart_app_runtime=set_autostart_runtime):
    return set_autostart_app_runtime(
        enable,
        autostart_flag=namespace["AUTOSTART_FLAG"],
        create_startup_shortcut=namespace["create_startup_shortcut"],
        remove_startup_shortcut=namespace["remove_startup_shortcut"],
        system_ui=namespace["system_ui"],
        show_error=lambda title, message: namespace["ui_messagebox"]("error", title, message),
    )


def cleanup_and_exit_namespace_runtime(namespace, *, cleanup_app_runtime=cleanup_and_exit_app_runtime):
    return cleanup_app_runtime(
        root=namespace["root"],
        messagebox=namespace["messagebox"],
        is_exiting=namespace["is_exiting"],
        set_exiting=lambda value: namespace.__setitem__("is_exiting", bool(value)),
        set_serial_running=lambda value: namespace.__setitem__("serial_running", bool(value)),
        shutdown_events=(namespace["TK_SHUTDOWN"],),
        worker_stop_events=(
            namespace["serial_stop_event"],
            namespace["serial_wakeup_event"],
        ),
        tts_stop_event=namespace["TTS_STOP"],
        safe_set_events=namespace["safe_set_events"],
        stop_cloud_control=namespace["stop_cloud_control"],
        safe_close_serial=namespace["safe_close_serial"],
        stop_tray_icon=namespace["stop_tray_icon"],
        flush_log_queue=namespace["flush_log_queue"],
        file_log_queue=namespace["FILE_LOG_Q"],
        file_log_thread=namespace.get("file_log_thread"),
        file_log_stop_event=namespace["file_log_stop"],
        worker_threads=lambda: _shutdown_worker_threads(namespace),
        deferred_worker_stop_events=(namespace["third_push_stop"],),
        deferred_worker_threads=lambda: (namespace.get("third_push_thread"),),
        deferred_worker_queues=(namespace["THIRD_PUSH_Q"],),
        destroy_root=namespace["root"].destroy,
        log_error=namespace.get("log_file_only"),
    )


def toggle_voice_broadcast_namespace_runtime(
    namespace,
    *,
    toggle_runtime=toggle_voice_broadcast_runtime,
):
    return toggle_runtime(
        namespace["VOICE_ENABLED"],
        namespace["config"],
        namespace["safe_save_config"],
        lambda value: namespace.__setitem__("VOICE_ENABLED", bool(value)),
        namespace["update_voice_menu_label"],
        namespace["system_ui"],
        log_error=namespace.get("log_file_only"),
    )


def toggle_multi_instance_namespace_runtime(
    namespace,
    *,
    toggle_runtime=toggle_multi_instance_runtime,
    acquire_mutex=acquire_mutex_with_error,
):
    previous = bool(namespace["ALLOW_MULTI_INSTANCE"])
    enabled = bool(namespace["multi_instance_var"].get())
    pending_mutex = None

    if previous and not enabled and not namespace.get("app_mutex"):
        try:
            pending_mutex, last_error = acquire_mutex(app_dir_mutex_name(namespace["APP_DIR"]))
        except Exception as exc:
            pending_mutex = None
            last_error = repr(exc)
        if not pending_mutex:
            try:
                namespace["multi_instance_var"].set(previous)
            except Exception:
                pass
            try:
                namespace["system_ui"](
                    "❌ 关闭程序多开失败：无法建立单实例保护，设置未更改",
                    "normal",
                )
            except Exception:
                pass
            try:
                namespace["ui_messagebox"](
                    "error",
                    "关闭多开失败",
                    "无法建立单实例保护，请稍后重试或重新启动软件。",
                )
            except Exception:
                pass
            log_error = namespace.get("log_file_only")
            if log_error is not None:
                try:
                    log_error(f"Acquire single-instance mutex while disabling multi-instance failed: {last_error}")
                except Exception:
                    pass
            return None

    try:
        result = toggle_runtime(
            enabled,
            namespace["config"],
            namespace["safe_save_config"],
            lambda value: namespace.__setitem__("ALLOW_MULTI_INSTANCE", bool(value)),
            namespace["system_ui"],
            show_notice=lambda title, message: namespace["ui_messagebox"]("info", title, message),
            log_error=namespace.get("log_file_only"),
        )
    except Exception:
        if pending_mutex:
            try:
                namespace["release_mutex_handle"](pending_mutex)
            except Exception:
                pass
        raise
    if result is None:
        if pending_mutex:
            try:
                namespace["release_mutex_handle"](pending_mutex)
            except Exception:
                pass
        try:
            namespace["multi_instance_var"].set(previous)
        except Exception:
            pass
    elif pending_mutex:
        namespace.__setitem__("app_mutex", pending_mutex)
    return result


def toggle_popup_namespace_runtime(namespace, *, toggle_runtime=toggle_popup_runtime):
    previous = bool(namespace["POPUP_ENABLED"])
    result = toggle_runtime(
        namespace["popup_var"].get(),
        namespace["config"],
        namespace["safe_save_config"],
        lambda value: namespace.__setitem__("POPUP_ENABLED", bool(value)),
        namespace["system_ui"],
        log_error=namespace.get("log_file_only"),
    )
    if result is None:
        try:
            namespace["popup_var"].set(previous)
        except Exception:
            pass
    return result


def toggle_call_popup_namespace_runtime(
    namespace,
    *,
    toggle_runtime=toggle_call_popup_runtime,
):
    previous = bool(namespace["CALL_POPUP_ENABLED"])
    result = toggle_runtime(
        namespace["call_popup_var"].get(),
        namespace["config"],
        namespace["safe_save_config"],
        lambda value: namespace.__setitem__("CALL_POPUP_ENABLED", bool(value)),
        namespace["system_ui"],
        log_error=namespace.get("log_file_only"),
    )
    if result is None:
        try:
            namespace["call_popup_var"].set(previous)
        except Exception:
            pass
    elif result is False:
        close_phone_popups_namespace_runtime(namespace)
    return result


def restart_software_namespace_runtime(
    namespace,
    *,
    restart_app_runtime=restart_software_app_runtime,
):
    os_module = namespace.get("os", os)
    return restart_app_runtime(
        root=namespace["root"],
        messagebox=namespace["messagebox"],
        is_exiting=namespace["is_exiting"],
        set_exiting=lambda value: namespace.__setitem__("is_exiting", bool(value)),
        set_serial_running=lambda value: namespace.__setitem__("serial_running", bool(value)),
        autostart_flag=namespace["AUTOSTART_FLAG"],
        restart_helper_flag=namespace["RESTART_HELPER_FLAG"],
        log_error=namespace["log_file_only"],
        system_ui=namespace["system_ui"],
        stop_tray_icon=namespace["stop_tray_icon"],
        safe_set_events=namespace["safe_set_events"],
        stop_events=(
            namespace["TK_SHUTDOWN"],
            namespace["serial_stop_event"],
            namespace["serial_wakeup_event"],
            namespace["TTS_STOP"],
        ),
        stop_cloud_control=namespace["stop_cloud_control"],
        safe_close_serial=namespace["safe_close_serial"],
        app_mutex=namespace["app_mutex"],
        release_mutex=namespace["release_mutex_handle"],
        flush_log_queue=namespace["flush_log_queue"],
        file_log_queue=namespace["FILE_LOG_Q"],
        file_log_thread=namespace.get("file_log_thread"),
        file_log_stop_event=namespace["file_log_stop"],
        worker_threads=lambda: _shutdown_worker_threads(namespace),
        deferred_stop_events=(namespace["third_push_stop"],),
        deferred_worker_threads=lambda: (namespace.get("third_push_thread"),),
        deferred_worker_queues=(namespace["THIRD_PUSH_Q"],),
        exit_process=os_module._exit,
    )
