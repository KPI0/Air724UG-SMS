import asyncio
import time

from sms_core.cloud_connection_runtime import (
    reply_cloud_payload_runtime,
    schedule_cloud_unregister_runtime,
    send_cloud_payload_runtime,
    unregister_then_close_cloud_connection_runtime,
)
from sms_core.cloud_message_namespace_runtime import (
    drain_cloud_serial_log_queue_namespace_runtime,
    handle_cloud_message_namespace_runtime,
    reset_cloud_serial_log_state_namespace_runtime,
    schedule_cloud_serial_log_drain_namespace_runtime,
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
    request_cloud_device_imei_namespace_runtime,
    set_cloud_device_imei_namespace_runtime,
)
from sms_core.cloud_ws_namespace_runtime import (
    cloud_thread_main_namespace_runtime,
    cloud_ws_main_namespace_runtime,
    send_cloud_register_namespace_runtime,
    send_cloud_unregister_namespace_runtime,
    wait_cloud_login_ack_namespace_runtime,
)
from sms_ui.cloud_control_namespace_runtime import (
    cloud_log_namespace_runtime,
    open_cloud_control_window_namespace_runtime,
    restart_cloud_control_namespace_runtime,
    save_cloud_control_setting_namespace_runtime,
    start_cloud_control_namespace_runtime,
    stop_cloud_control_namespace_runtime,
)


def install_cloud_namespace_bindings(namespace):
    def cloud_runtime_imei():
        return cloud_runtime_imei_namespace_runtime(namespace)

    def cloud_identity_payload():
        return cloud_identity_payload_namespace_runtime(namespace)

    async def cloud_wait_login_ack(ws, timeout=8.0):
        return await wait_cloud_login_ack_namespace_runtime(namespace, ws, timeout=timeout)

    async def cloud_send_register(ws):
        return await send_cloud_register_namespace_runtime(namespace, ws)

    async def cloud_send_unregister(ws, reason="hidden"):
        return await send_cloud_unregister_namespace_runtime(namespace, ws, reason=reason)

    def cloud_schedule_unregister(reason="hidden"):
        return schedule_cloud_unregister_runtime(
            reason=reason,
            get_loop=lambda: namespace["cloud_ws_loop"],
            get_ws=lambda: namespace["cloud_ws_conn"],
            is_connected=lambda: namespace["cloud_connected"],
            send_unregister=namespace["_cloud_send_unregister"],
            run_coroutine_threadsafe=asyncio.run_coroutine_threadsafe,
        )

    async def cloud_unregister_then_close(ws, reason="disconnect"):
        return await unregister_then_close_cloud_connection_runtime(
            ws,
            reason=reason,
            auto_upload=namespace["CLOUD_AUTO_UPLOAD"],
            send_unregister=namespace["_cloud_send_unregister"],
        )

    def notify_cloud_identity_changed():
        return notify_cloud_identity_changed_namespace_runtime(namespace)

    def set_cloud_device_imei(imei, source=""):
        return set_cloud_device_imei_namespace_runtime(namespace, imei, source=source)

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
        return await send_cloud_payload_runtime(ws, payload)

    def reset_cloud_serial_log_state():
        return reset_cloud_serial_log_state_namespace_runtime(namespace)

    async def cloud_drain_serial_log_queue(ws):
        return await drain_cloud_serial_log_queue_namespace_runtime(namespace, ws)

    def schedule_cloud_serial_log_drain(loop, ws):
        return schedule_cloud_serial_log_drain_namespace_runtime(namespace, loop, ws)

    def cloud_send_serial_log(line):
        return send_cloud_serial_log_namespace_runtime(namespace, line)

    def cloud_send_sms_event(callback_head, full_msg):
        return send_cloud_sms_event_namespace_runtime(namespace, callback_head, full_msg)

    def cloud_auth_matches(data):
        return cloud_auth_matches_namespace_runtime(namespace, data)

    def request_cloud_device_imei():
        return request_cloud_device_imei_namespace_runtime(namespace)

    def maybe_capture_cloud_device_imei(line):
        return maybe_capture_cloud_device_imei_namespace_runtime(namespace, line)

    def cloud_send_status_payload():
        return cloud_status_payload_namespace_runtime(namespace)

    def cloud_now_ts():
        return int(time.time())

    async def cloud_check_replay_window(ws, data, mark_seen=True):
        return await cloud_check_replay_window_namespace_runtime(
            namespace,
            ws,
            data,
            mark_seen=mark_seen,
        )

    def cloud_send_serial_command(command):
        return send_cloud_serial_command_namespace_runtime(namespace, command)

    async def cloud_reply(ws, payload):
        return await reply_cloud_payload_runtime(
            ws,
            payload,
            identity_payload=namespace["_cloud_identity_payload"],
            log=namespace["_cloud_log"],
        )

    async def handle_cloud_message(ws, message):
        return await handle_cloud_message_namespace_runtime(namespace, ws, message)

    async def cloud_ws_main(url, reconnect_interval):
        return await cloud_ws_main_namespace_runtime(namespace, url, reconnect_interval)

    def cloud_thread_main(url, reconnect_interval):
        return cloud_thread_main_namespace_runtime(namespace, url, reconnect_interval)

    def start_cloud_control(show_errors=False):
        return start_cloud_control_namespace_runtime(namespace, show_errors=show_errors)

    def stop_cloud_control(update_status=True):
        return stop_cloud_control_namespace_runtime(namespace, update_status=update_status)

    def restart_cloud_control(show_errors=False):
        return restart_cloud_control_namespace_runtime(namespace, show_errors=show_errors)

    def open_cloud_control_window():
        return open_cloud_control_window_namespace_runtime(namespace)

    namespace.update({
        "_cloud_runtime_imei": cloud_runtime_imei,
        "_cloud_identity_payload": cloud_identity_payload,
        "_cloud_wait_login_ack": cloud_wait_login_ack,
        "_cloud_send_register": cloud_send_register,
        "_cloud_send_unregister": cloud_send_unregister,
        "_cloud_schedule_unregister": cloud_schedule_unregister,
        "_cloud_unregister_then_close": cloud_unregister_then_close,
        "_notify_cloud_identity_changed": notify_cloud_identity_changed,
        "_set_cloud_device_imei": set_cloud_device_imei,
        "save_cloud_control_setting": save_cloud_control_setting,
        "_cloud_log": cloud_log,
        "_cloud_send_payload": cloud_send_payload,
        "_reset_cloud_serial_log_state": reset_cloud_serial_log_state,
        "_cloud_drain_serial_log_queue": cloud_drain_serial_log_queue,
        "_schedule_cloud_serial_log_drain": schedule_cloud_serial_log_drain,
        "_cloud_send_serial_log": cloud_send_serial_log,
        "_cloud_send_sms_event": cloud_send_sms_event,
        "_cloud_auth_matches": cloud_auth_matches,
        "request_cloud_device_imei": request_cloud_device_imei,
        "_maybe_capture_cloud_device_imei": maybe_capture_cloud_device_imei,
        "_cloud_send_status_payload": cloud_send_status_payload,
        "_cloud_now_ts": cloud_now_ts,
        "_cloud_check_replay_window": cloud_check_replay_window,
        "_cloud_send_serial_command": cloud_send_serial_command,
        "_cloud_reply": cloud_reply,
        "_handle_cloud_message": handle_cloud_message,
        "_cloud_ws_main": cloud_ws_main,
        "_cloud_thread_main": cloud_thread_main,
        "start_cloud_control": start_cloud_control,
        "stop_cloud_control": stop_cloud_control,
        "restart_cloud_control": restart_cloud_control,
        "open_cloud_control_window": open_cloud_control_window,
    })
    return namespace
