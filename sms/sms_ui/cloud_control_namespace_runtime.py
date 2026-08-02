import asyncio
import threading

from sms_core.config_runtime import reload_config_runtime
from sms_core.cloud_runtime import read_cloud_control_settings
from sms_ui.cloud_control_app_runtime import (
    cloud_control_settings_from_values,
    open_cloud_control_values_app_runtime,
    restart_cloud_control_app_runtime,
    save_cloud_control_setting_runtime,
    start_cloud_control_app_runtime,
    stop_cloud_control_app_runtime,
)


def refresh_cloud_control_settings_namespace_runtime(
    namespace,
    *,
    reload_config=reload_config_runtime,
):
    try:
        settings = reload_config(
            config=namespace["config"],
            config_file=namespace["CONFIG_FILE"],
            config_lock=namespace["CONFIG_LOCK"],
            read_values=read_cloud_control_settings,
        )
    except Exception as exc:
        log_error = namespace.get("log_file_only")
        if log_error is not None:
            try:
                log_error(f"Reload cloud-control config failed ({type(exc).__name__})")
            except Exception:
                pass
        return False

    namespace["apply_cloud_control_settings"](settings)
    return True


def start_cloud_control_namespace_runtime(
    namespace,
    *,
    show_errors=False,
    start_app_runtime=start_cloud_control_app_runtime,
):
    return start_app_runtime(
        websockets_available=namespace["websockets"] is not None,
        url=namespace["CLOUD_WS_URL"],
        device_secret=namespace["CLOUD_DEVICE_SECRET"],
        reconnect_interval=namespace["CLOUD_WS_RECONNECT_INTERVAL"],
        show_errors=show_errors,
        set_cloud_status=namespace["set_cloud_status"],
        cloud_log=namespace["_cloud_log"],
        show_warning=namespace["messagebox"].showwarning,
        runtime_imei=namespace["_cloud_runtime_imei"],
        request_device_imei=namespace["request_cloud_device_imei"],
        lock=namespace["cloud_ws_lock"],
        get_thread=lambda: namespace["cloud_ws_thread"],
        set_thread=lambda thread: namespace.__setitem__("cloud_ws_thread", thread),
        stop_event=namespace["cloud_stop_event"],
        thread_factory=namespace.get("threading", threading).Thread,
        thread_target=namespace["_cloud_thread_main"],
    )


def stop_cloud_control_namespace_runtime(
    namespace,
    *,
    update_status=True,
    stop_app_runtime=stop_cloud_control_app_runtime,
):
    return stop_app_runtime(
        update_status=update_status,
        enabled=namespace["CLOUD_CONTROL_ENABLED"],
        stop_event=namespace["cloud_stop_event"],
        set_connected=lambda value: namespace.__setitem__("cloud_connected", bool(value)),
        set_authorized=lambda value: namespace.__setitem__("cloud_device_authorized", bool(value)),
        reset_serial_log_state=namespace["_reset_cloud_serial_log_state"],
        get_loop=lambda: namespace["cloud_ws_loop"],
        get_ws=lambda: namespace["cloud_ws_conn"],
        schedule_unregister_then_close=namespace["_cloud_unregister_then_close"],
        set_ws=lambda value: namespace.__setitem__("cloud_ws_conn", value),
        set_cloud_status=namespace["set_cloud_status"],
        run_coroutine_threadsafe=namespace.get("asyncio", asyncio).run_coroutine_threadsafe,
    )


def restart_cloud_control_namespace_runtime(
    namespace,
    *,
    show_errors=False,
    restart_app_runtime=restart_cloud_control_app_runtime,
):
    return restart_app_runtime(
        show_errors=show_errors,
        lock=namespace["cloud_ws_lock"],
        get_restart_seq=lambda: namespace["cloud_restart_seq"],
        set_restart_seq=lambda value: namespace.__setitem__("cloud_restart_seq", value),
        get_thread=lambda: namespace["cloud_ws_thread"],
        stop_control=namespace["stop_cloud_control"],
        tk_alive=namespace["tk_alive"],
        stop_event=namespace["cloud_stop_event"],
        set_cloud_status=namespace["set_cloud_status"],
        schedule_after=namespace["root"].after,
        ui_post=namespace["ui_post"],
        start_control=namespace["start_cloud_control"],
        thread_factory=namespace.get("threading", threading).Thread,
        cloud_log=namespace["_cloud_log"],
    )


def save_cloud_control_setting_namespace_runtime(
    namespace,
    *,
    enabled=None,
    url=None,
    reconnect_interval=None,
    device_secret=None,
    auto_upload=None,
    save_app_runtime=save_cloud_control_setting_runtime,
):
    return save_app_runtime(
        current_settings=lambda: cloud_control_settings_from_values(
            enabled=namespace["CLOUD_CONTROL_ENABLED"],
            url=namespace["CLOUD_WS_URL"],
            reconnect_interval=namespace["CLOUD_WS_RECONNECT_INTERVAL"],
            device_secret=namespace["CLOUD_DEVICE_SECRET"],
            auto_upload=namespace["CLOUD_AUTO_UPLOAD"],
        ),
        apply_settings=namespace["apply_cloud_control_settings"],
        config=namespace["config"],
        save_config=namespace["safe_save_config"],
        system_ui=namespace["system_ui"],
        enabled=enabled,
        url=url,
        reconnect_interval=reconnect_interval,
        device_secret=device_secret,
        auto_upload=auto_upload,
    )


def cloud_log_namespace_runtime(namespace, message, *, show_main=False):
    shutdown_event = namespace.get("TK_SHUTDOWN")
    try:
        if shutdown_event is not None and shutdown_event.is_set():
            return None
    except Exception:
        pass

    base_msg = f"🌐 {message}"
    try:
        file_msg = namespace["_cloud_repeat_filter"](namespace["_cloud_file_notice"], base_msg)
        if file_msg is not None:
            namespace["log_file_only"](file_msg)
    except Exception:
        namespace["log_file_only"](base_msg)

    if not show_main:
        return None

    try:
        ui_msg = namespace["_cloud_repeat_filter"](namespace["_cloud_main_notice"], base_msg)
        if ui_msg is not None:
            namespace["ui_only"](ui_msg, "normal")
    except Exception:
        namespace["ui_only"](base_msg, "normal")

    return None


def open_cloud_control_window_namespace_runtime(
    namespace,
    *,
    open_values_app_runtime=open_cloud_control_values_app_runtime,
):
    def sync_existing_window(win, attr):
        return namespace["sync_and_focus_existing_window"](
            win,
            attr,
            log_error=namespace.get("log_file_only"),
        )

    return open_values_app_runtime(
        namespace["root"],
        current_window=namespace["cloud_control_win"],
        enabled=namespace["CLOUD_CONTROL_ENABLED"],
        auto_upload=namespace["CLOUD_AUTO_UPLOAD"],
        url=namespace["CLOUD_WS_URL"],
        secret=namespace["CLOUD_DEVICE_SECRET"],
        reconnect_interval=namespace["CLOUD_WS_RECONNECT_INTERVAL"],
        status_var=namespace["cloud_var"],
        refresh_settings=namespace["refresh_cloud_control_settings_from_config"],
        save_setting=namespace["save_cloud_control_setting"],
        is_connected=lambda: namespace["cloud_connected"],
        get_loop=lambda: namespace["cloud_ws_loop"],
        get_ws=lambda: namespace["cloud_ws_conn"],
        send_register=namespace["_cloud_send_register"],
        run_coroutine_threadsafe=namespace.get("asyncio", asyncio).run_coroutine_threadsafe,
        schedule_unregister=namespace["_cloud_schedule_unregister"],
        restart_control=namespace["restart_cloud_control"],
        stop_control=namespace["stop_cloud_control"],
        cloud_log=namespace["_cloud_log"],
        sync_existing_window=sync_existing_window,
        set_window=lambda win: namespace.__setitem__("cloud_control_win", win),
        center_window=namespace["center_window"],
        open_security_settings=lambda: namespace["open_security_settings"](),
        settings_provider=lambda: {
            "enabled": namespace["CLOUD_CONTROL_ENABLED"],
            "auto_upload": namespace["CLOUD_AUTO_UPLOAD"],
            "url": namespace["CLOUD_WS_URL"],
            "secret": namespace["CLOUD_DEVICE_SECRET"],
            "reconnect_interval": namespace["CLOUD_WS_RECONNECT_INTERVAL"],
        },
    )
