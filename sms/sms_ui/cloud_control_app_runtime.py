from sms_core.cloud_runtime import (
    CloudControlSettings,
    cloud_control_state,
    restart_cloud_control_runtime,
    start_cloud_control_runtime,
    stop_cloud_control_runtime,
    update_cloud_control_settings,
    validate_cloud_start,
    write_cloud_control_settings,
)
from sms_ui.cloud_control_window import open_cloud_control_window_runtime


def cloud_control_settings_from_values(enabled, url, reconnect_interval, device_secret, auto_upload):
    return CloudControlSettings(
        enabled=enabled,
        url=url,
        reconnect_interval=reconnect_interval,
        device_secret=device_secret,
        auto_upload=auto_upload,
    )


def save_cloud_control_setting_runtime(
    *,
    current_settings,
    apply_settings,
    config,
    save_config,
    system_ui,
    enabled=None,
    url=None,
    reconnect_interval=None,
    device_secret=None,
    auto_upload=None,
    update_settings=update_cloud_control_settings,
    write_settings=write_cloud_control_settings,
):
    settings = update_settings(
        current_settings(),
        enabled=enabled,
        url=url,
        reconnect_interval=reconnect_interval,
        device_secret=device_secret,
        auto_upload=auto_upload,
    )
    apply_settings(settings)

    try:
        write_settings(config, settings)
        if save_config() is False:
            raise RuntimeError("配置保存失败")
    except Exception as exc:
        system_ui(f"❌ 云端控制配置保存失败：{exc}", "normal")
        return None

    return settings


def start_cloud_control_app_runtime(
    *,
    websockets_available,
    url,
    device_secret,
    reconnect_interval,
    show_errors,
    set_cloud_status,
    cloud_log,
    show_warning,
    runtime_imei,
    request_device_imei,
    lock,
    get_thread,
    set_thread,
    stop_event,
    thread_factory,
    thread_target,
    start_runtime=start_cloud_control_runtime,
):
    return start_runtime(
        websockets_available=websockets_available,
        url=url,
        device_secret=device_secret,
        reconnect_interval=reconnect_interval,
        show_errors=show_errors,
        validate_start=validate_cloud_start,
        set_cloud_status=set_cloud_status,
        log_missing_dependency=lambda: cloud_log("缺少 websockets 库，无法启动云端控制", show_main=True),
        show_warning=show_warning,
        runtime_imei=runtime_imei,
        request_device_imei=request_device_imei,
        lock=lock,
        get_thread=get_thread,
        set_thread=set_thread,
        stop_event=stop_event,
        thread_factory=thread_factory,
        thread_target=thread_target,
        log_error=lambda message: cloud_log(message, show_main=True),
    )


def stop_cloud_control_app_runtime(
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
    stop_runtime=stop_cloud_control_runtime,
):
    return stop_runtime(
        update_status=update_status,
        enabled=enabled,
        stop_event=stop_event,
        set_connected=set_connected,
        set_authorized=set_authorized,
        reset_serial_log_state=reset_serial_log_state,
        get_loop=get_loop,
        get_ws=get_ws,
        schedule_unregister_then_close=schedule_unregister_then_close,
        set_ws=set_ws,
        set_cloud_status=set_cloud_status,
        run_coroutine_threadsafe=run_coroutine_threadsafe,
    )


def restart_cloud_control_app_runtime(
    *,
    show_errors,
    lock,
    get_restart_seq,
    set_restart_seq,
    get_thread,
    stop_control,
    tk_alive,
    stop_event,
    set_cloud_status,
    schedule_after,
    ui_post,
    start_control,
    thread_factory,
    cloud_log=None,
    restart_runtime=restart_cloud_control_runtime,
):
    def increment_restart_seq():
        next_seq = get_restart_seq() + 1
        set_restart_seq(next_seq)
        return next_seq

    return restart_runtime(
        show_errors=show_errors,
        lock=lock,
        increment_restart_seq=increment_restart_seq,
        get_restart_seq=get_restart_seq,
        get_thread=get_thread,
        stop_control=stop_control,
        tk_alive=tk_alive,
        stop_event=stop_event,
        set_cloud_status=set_cloud_status,
        schedule_after=schedule_after,
        ui_post=ui_post,
        start_control=start_control,
        thread_factory=thread_factory,
        log_error=(lambda message: cloud_log(message, show_main=True)) if cloud_log is not None else None,
    )


def cloud_window_connection_state(is_connected, get_loop, get_ws):
    return (bool(is_connected()), get_loop() is not None, get_ws() is not None)


def register_current_cloud_connection(get_loop, get_ws, send_register, run_coroutine_threadsafe):
    loop = get_loop()
    ws = get_ws()
    return run_coroutine_threadsafe(send_register(ws), loop)


def open_cloud_control_app_runtime(
    parent,
    *,
    current_window,
    get_settings,
    status_var,
    refresh_settings,
    save_setting,
    get_connection_state=None,
    register_current=None,
    is_connected=None,
    get_loop=None,
    get_ws=None,
    send_register=None,
    run_coroutine_threadsafe=None,
    schedule_unregister,
    restart_control,
    stop_control,
    cloud_log,
    sync_existing_window,
    set_window,
    center_window,
    open_window_runtime=open_cloud_control_window_runtime,
):
    def current_state():
        settings = get_settings()
        return cloud_control_state(
            settings["enabled"],
            settings["auto_upload"],
            settings["url"],
            settings["secret"],
            settings["reconnect_interval"],
        )

    if get_connection_state is None:
        get_connection_state = lambda: cloud_window_connection_state(is_connected, get_loop, get_ws)
    if register_current is None:
        register_current = lambda: register_current_cloud_connection(
            get_loop,
            get_ws,
            send_register,
            run_coroutine_threadsafe,
        )

    return open_window_runtime(
        parent,
        current_window,
        current_state,
        status_var,
        refresh_settings,
        save_setting,
        get_connection_state,
        register_current,
        schedule_unregister,
        restart_control,
        stop_control,
        cloud_log,
        sync_existing_window,
        set_window,
        center_window,
    )


def open_cloud_control_values_app_runtime(
    parent,
    *,
    current_window,
    enabled,
    auto_upload,
    url,
    secret,
    reconnect_interval,
    status_var,
    refresh_settings,
    save_setting,
    get_connection_state=None,
    register_current=None,
    is_connected=None,
    get_loop=None,
    get_ws=None,
    send_register=None,
    run_coroutine_threadsafe=None,
    schedule_unregister,
    restart_control,
    stop_control,
    cloud_log,
    sync_existing_window,
    set_window,
    center_window,
    open_window_runtime=open_cloud_control_window_runtime,
):
    return open_cloud_control_app_runtime(
        parent,
        current_window=current_window,
        get_settings=lambda: {
            "enabled": enabled,
            "auto_upload": auto_upload,
            "url": url,
            "secret": secret,
            "reconnect_interval": reconnect_interval,
        },
        status_var=status_var,
        refresh_settings=refresh_settings,
        save_setting=save_setting,
        get_connection_state=get_connection_state,
        register_current=register_current,
        is_connected=is_connected,
        get_loop=get_loop,
        get_ws=get_ws,
        send_register=send_register,
        run_coroutine_threadsafe=run_coroutine_threadsafe,
        schedule_unregister=schedule_unregister,
        restart_control=restart_control,
        stop_control=stop_control,
        cloud_log=cloud_log,
        sync_existing_window=sync_existing_window,
        set_window=set_window,
        center_window=center_window,
        open_window_runtime=open_window_runtime,
    )
