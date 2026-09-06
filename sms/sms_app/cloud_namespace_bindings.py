import asyncio
import time

from sms_core.namespace_binding import (
    make_async_namespace_runtime_binder,
    make_namespace_runtime_binder,
)
from sms_core.cloud_connection_runtime import (
    reply_cloud_payload_runtime,
    schedule_cloud_unregister_runtime,
    send_cloud_payload_runtime,
    unregister_then_close_cloud_connection_runtime,
)
from sms_core.cloud_message_namespace_runtime import (
    clear_cloud_sms_event_state_namespace_runtime,
    drain_cloud_sms_event_queue_namespace_runtime,
    drain_cloud_serial_log_queue_namespace_runtime,
    handle_cloud_message_namespace_runtime,
    handle_cloud_call_recording_message_namespace_runtime,
    reset_cloud_serial_log_state_namespace_runtime,
    schedule_cloud_serial_log_drain_namespace_runtime,
    schedule_cloud_sms_event_drain_namespace_runtime,
    schedule_cloud_call_recording_upload_namespace_runtime,
    send_cloud_call_recording_status_namespace_runtime,
    send_cloud_call_event_namespace_runtime,
    send_cloud_call_state_namespace_runtime,
    send_cloud_serial_command_namespace_runtime,
    send_cloud_serial_log_namespace_runtime,
    send_cloud_sms_event_namespace_runtime,
)
from sms_core.cloud_state_namespace_runtime import (
    cloud_auth_matches_namespace_runtime,
    cloud_check_replay_window_namespace_runtime,
    cloud_identity_payload_namespace_runtime,
    cloud_runtime_imei_namespace_runtime,
    cloud_status_payload_namespace_runtime,
    maybe_capture_cloud_device_imei_namespace_runtime,
    notify_cloud_identity_changed_namespace_runtime,
    notify_cloud_channel_status_namespace_runtime,
    request_cloud_device_imei_namespace_runtime,
    set_cloud_device_imei_namespace_runtime,
)
from sms_core.cloud_ws_namespace_runtime import (
    cloud_thread_main_namespace_runtime,
    cloud_ws_main_namespace_runtime,
    send_cloud_register_namespace_runtime,
    send_cloud_session_revoke_namespace_runtime,
    send_cloud_unregister_namespace_runtime,
    wait_cloud_login_ack_namespace_runtime,
)
from sms_ui.cloud_control_namespace_runtime import (
    cloud_log_namespace_runtime,
    open_cloud_control_window_namespace_runtime,
    refresh_cloud_control_settings_namespace_runtime,
    restart_cloud_control_namespace_runtime,
    save_cloud_control_setting_namespace_runtime,
    start_cloud_control_namespace_runtime,
    stop_cloud_control_namespace_runtime,
)


def install_cloud_namespace_bindings(namespace):
    module_globals = globals()
    bind = make_namespace_runtime_binder(namespace, module_globals)
    bind_async = make_async_namespace_runtime_binder(namespace, module_globals)

    def cloud_schedule_unregister(reason="hidden"):
        return schedule_cloud_unregister_runtime(
            reason=reason,
            get_loop=lambda: namespace["cloud_ws_loop"],
            get_ws=lambda: namespace["cloud_ws_conn"],
            is_connected=lambda: namespace["cloud_connected"],
            send_unregister=namespace["_cloud_send_unregister"],
            run_coroutine_threadsafe=asyncio.run_coroutine_threadsafe,
            log_error=namespace.get("_cloud_log"),
        )

    async def cloud_unregister_then_close(ws, reason="disconnect"):
        return await unregister_then_close_cloud_connection_runtime(
            ws,
            reason=reason,
            auto_upload=namespace["CLOUD_AUTO_UPLOAD"],
            send_unregister=namespace["_cloud_send_unregister"],
            send_session_revoke=namespace["_cloud_send_session_revoke"],
            log_error=namespace.get("_cloud_log"),
        )

    def save_cloud_control_setting(
        enabled=None,
        url=None,
        reconnect_interval=None,
        device_secret=None,
        auto_upload=None,
    ):
        return save_cloud_control_setting_namespace_runtime(
            namespace,
            enabled=enabled,
            url=url,
            reconnect_interval=reconnect_interval,
            device_secret=device_secret,
            auto_upload=auto_upload,
        )

    def cloud_log(message, show_main=False):
        return cloud_log_namespace_runtime(namespace, message, show_main=show_main)

    async def cloud_send_payload(ws, payload):
        kwargs = {}
        log_error = namespace.get("log_file_only")
        if log_error is not None:
            kwargs["log_error"] = log_error
        return await send_cloud_payload_runtime(ws, payload, **kwargs)

    def cloud_now_ts():
        return int(time.time())

    async def cloud_check_replay_window(ws, data, mark_seen=True):
        return await cloud_check_replay_window_namespace_runtime(
            namespace,
            ws,
            data,
            mark_seen=mark_seen,
        )

    async def cloud_reply(ws, payload):
        return await reply_cloud_payload_runtime(
            ws,
            payload,
            identity_payload=namespace["_cloud_identity_payload"],
            log=namespace["_cloud_log"],
        )

    namespace.update({
        "_cloud_runtime_imei": bind("cloud_runtime_imei_namespace_runtime"),
        "_cloud_identity_payload": bind("cloud_identity_payload_namespace_runtime"),
        "_cloud_wait_login_ack": bind_async(
            "wait_cloud_login_ack_namespace_runtime",
            positional_keywords=("timeout",),
            positional_prefix_count=1,
        ),
        "_cloud_send_register": bind_async("send_cloud_register_namespace_runtime"),
        "_cloud_send_session_revoke": bind_async(
            "send_cloud_session_revoke_namespace_runtime",
            positional_keywords=("reason",),
            positional_prefix_count=1,
        ),
        "_cloud_send_unregister": bind_async(
            "send_cloud_unregister_namespace_runtime",
            positional_keywords=("reason",),
            positional_prefix_count=1,
        ),
        "_cloud_schedule_unregister": cloud_schedule_unregister,
        "_cloud_unregister_then_close": cloud_unregister_then_close,
        "_notify_cloud_identity_changed": bind("notify_cloud_identity_changed_namespace_runtime"),
        "_notify_cloud_channel_status": bind("notify_cloud_channel_status_namespace_runtime"),
        "_set_cloud_device_imei": bind(
            "set_cloud_device_imei_namespace_runtime",
            positional_keywords=("source",),
            positional_prefix_count=1,
        ),
        "save_cloud_control_setting": save_cloud_control_setting,
        "_cloud_log": cloud_log,
        "_cloud_send_payload": cloud_send_payload,
        "_clear_cloud_sms_event_state": bind("clear_cloud_sms_event_state_namespace_runtime"),
        "_cloud_drain_sms_event_queue": bind_async("drain_cloud_sms_event_queue_namespace_runtime"),
        "_schedule_cloud_sms_event_drain": bind("schedule_cloud_sms_event_drain_namespace_runtime"),
        "_schedule_cloud_call_recording_upload": bind(
            "schedule_cloud_call_recording_upload_namespace_runtime"
        ),
        "_send_cloud_call_recording_status": bind(
            "send_cloud_call_recording_status_namespace_runtime"
        ),
        "_reset_cloud_serial_log_state": bind("reset_cloud_serial_log_state_namespace_runtime"),
        "_cloud_drain_serial_log_queue": bind_async("drain_cloud_serial_log_queue_namespace_runtime"),
        "_schedule_cloud_serial_log_drain": bind("schedule_cloud_serial_log_drain_namespace_runtime"),
        "_cloud_send_serial_log": bind("send_cloud_serial_log_namespace_runtime"),
        "_cloud_send_sms_event": bind("send_cloud_sms_event_namespace_runtime"),
        "_cloud_send_call_event": bind("send_cloud_call_event_namespace_runtime"),
        "_cloud_send_call_state": bind("send_cloud_call_state_namespace_runtime"),
        "_cloud_auth_matches": bind("cloud_auth_matches_namespace_runtime"),
        "request_cloud_device_imei": bind("request_cloud_device_imei_namespace_runtime"),
        "_maybe_capture_cloud_device_imei": bind("maybe_capture_cloud_device_imei_namespace_runtime"),
        "_cloud_send_status_payload": bind("cloud_status_payload_namespace_runtime"),
        "_cloud_now_ts": cloud_now_ts,
        "_cloud_check_replay_window": cloud_check_replay_window,
        "_cloud_send_serial_command": bind("send_cloud_serial_command_namespace_runtime"),
        "_cloud_reply": cloud_reply,
        "_handle_cloud_call_recording_message": bind(
            "handle_cloud_call_recording_message_namespace_runtime"
        ),
        "_handle_cloud_message": bind_async("handle_cloud_message_namespace_runtime"),
        "_cloud_ws_main": bind_async("cloud_ws_main_namespace_runtime"),
        "_cloud_thread_main": bind("cloud_thread_main_namespace_runtime"),
        "start_cloud_control": bind(
            "start_cloud_control_namespace_runtime",
            positional_keywords=("show_errors",),
        ),
        "stop_cloud_control": bind(
            "stop_cloud_control_namespace_runtime",
            positional_keywords=("update_status",),
        ),
        "restart_cloud_control": bind(
            "restart_cloud_control_namespace_runtime",
            positional_keywords=("show_errors",),
        ),
        "refresh_cloud_control_settings_from_config": bind(
            "refresh_cloud_control_settings_namespace_runtime"
        ),
        "open_cloud_control_window": bind("open_cloud_control_window_namespace_runtime"),
    })
    return namespace
