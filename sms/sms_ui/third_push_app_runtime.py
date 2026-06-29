from sms_core.third_push import third_push_state
from sms_core.third_push_config import ThirdPushSettings, update_third_push_settings, write_third_push_settings
from sms_core.third_push_runtime import enqueue_third_push_runtime, third_push_worker_runtime
from sms_core.third_push import format_message as format_third_push_message
from sms_core.third_push import send_channel as send_third_push_channel
from sms_ui.third_push_result_runtime import show_third_push_test_result_runtime
from sms_ui.third_push_window import open_third_push_window_runtime


def save_third_push_setting_runtime(
    *,
    current_settings,
    apply_settings,
    config,
    save_config,
    enabled=None,
    sms_enabled=None,
    call_enabled=None,
    notify_type=None,
    settings=None,
    update_settings=update_third_push_settings,
    write_settings=write_third_push_settings,
):
    next_settings = update_settings(
        current_settings(),
        enabled=enabled,
        sms_enabled=sms_enabled,
        call_enabled=call_enabled,
        notify_type=notify_type,
        settings=settings,
    )
    apply_settings(next_settings)

    try:
        write_settings(config, next_settings)
        if save_config() is False:
            raise RuntimeError("配置保存失败")
    except Exception:
        return None

    return next_settings


def third_push_settings_from_values(enabled, sms_enabled, call_enabled, channels, settings):
    return ThirdPushSettings(
        enabled,
        sms_enabled,
        call_enabled,
        channels,
        settings,
    )


def open_third_push_app_runtime(
    *,
    parent,
    current_window,
    enabled,
    sms_enabled,
    call_enabled,
    channels,
    settings,
    refresh_settings,
    save_setting,
    enqueue_push,
    system_ui,
    sync_existing_window,
    set_window,
    center_window,
    open_window_runtime=open_third_push_window_runtime,
):
    def current_state():
        return third_push_state(
            enabled,
            sms_enabled,
            call_enabled,
            channels,
            settings,
        )

    return open_window_runtime(
        parent,
        current_window,
        current_state,
        refresh_settings,
        save_setting,
        enqueue_push,
        system_ui,
        sync_existing_window,
        set_window,
        center_window,
    )


def format_third_push_message_runtime(
    raw_msg,
    template,
    *,
    get_log_prefix,
    variables=None,
    format_message=format_third_push_message,
):
    try:
        return format_message(raw_msg, template, port=get_log_prefix(), variables=variables)
    except TypeError:
        return format_message(raw_msg, template, port=get_log_prefix())


def send_third_push_channel_runtime(
    channel,
    message,
    settings,
    *,
    get_log_prefix,
    app_version,
    send_channel=send_third_push_channel,
):
    return send_channel(
        channel,
        message,
        settings,
        user_agent=f"Air724UG-SMS/{app_version}",
        port=get_log_prefix(),
    )


def third_push_worker_app_runtime(
    *,
    stop_event,
    push_queue,
    get_log_prefix,
    app_version,
    system_ui,
    show_result,
    worker_runtime=third_push_worker_runtime,
):
    return worker_runtime(
        stop_event=stop_event,
        push_queue=push_queue,
        send_channel_func=lambda channel, message, settings: send_third_push_channel_runtime(
            channel,
            message,
            settings,
            get_log_prefix=get_log_prefix,
            app_version=app_version,
        ),
        system_ui=system_ui,
        show_result=show_result,
        format_message_func=lambda raw_msg, template=None, variables=None: format_third_push_message_runtime(
            raw_msg,
            template,
            get_log_prefix=get_log_prefix,
            variables=variables,
        ),
    )


def show_third_push_test_result_app_runtime(
    *,
    root,
    get_current_window,
    messagebox,
    ui_post,
    ok_channels,
    fail_infos,
    show_runtime=show_third_push_test_result_runtime,
):
    return show_runtime(
        root=root,
        current_window=get_current_window(),
        messagebox=messagebox,
        ui_post=ui_post,
        ok_channels=ok_channels,
        fail_infos=fail_infos,
    )


def enqueue_third_push_app_runtime(
    raw_msg,
    *,
    push_queue,
    enabled,
    sms_enabled,
    call_enabled,
    configured_channels,
    current_settings,
    system_ui,
    channels=None,
    settings=None,
    template=None,
    variables=None,
    event_type="sms",
    show_success=False,
    show_result=False,
    enqueue_runtime=enqueue_third_push_runtime,
):
    return enqueue_runtime(
        raw_msg,
        push_queue=push_queue,
        enabled=enabled(),
        sms_enabled=sms_enabled(),
        call_enabled=call_enabled(),
        configured_channels=configured_channels(),
        current_settings=current_settings(),
        channels=channels,
        settings=settings,
        template=template,
        variables=variables,
        event_type=event_type,
        show_success=show_success,
        show_result=show_result,
        system_ui=system_ui,
    )


def open_third_push_values_app_runtime(
    *,
    parent,
    current_window,
    enabled,
    sms_enabled,
    call_enabled,
    channels,
    settings,
    refresh_settings,
    save_setting,
    enqueue_push,
    system_ui,
    sync_existing_window,
    set_window,
    center_window,
    open_app_runtime=open_third_push_app_runtime,
):
    return open_app_runtime(
        parent=parent,
        current_window=current_window(),
        enabled=enabled(),
        sms_enabled=sms_enabled(),
        call_enabled=call_enabled(),
        channels=channels(),
        settings=settings(),
        refresh_settings=refresh_settings,
        save_setting=save_setting,
        enqueue_push=enqueue_push,
        system_ui=system_ui,
        sync_existing_window=sync_existing_window,
        set_window=set_window,
        center_window=center_window,
    )
