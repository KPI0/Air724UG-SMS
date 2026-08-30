import inspect
import time

from sms_core.cloud_message_runtime import (
    send_cloud_register_runtime,
    send_cloud_session_revoke_runtime,
    send_cloud_unregister_runtime,
)
from sms_core.cloud_ws_runtime import (
    cloud_thread_main_runtime,
    cloud_ws_main_app_runtime,
    wait_cloud_login_ack_runtime,
)


async def wait_cloud_login_ack_namespace_runtime(
    namespace,
    ws,
    *,
    timeout=8.0,
    wait_runtime=wait_cloud_login_ack_runtime,
):
    result = await wait_runtime(
        ws,
        stop_event=namespace["cloud_stop_event"],
        set_authorized=lambda value: namespace.__setitem__("cloud_device_authorized", bool(value)),
        set_auth_status_from_ack=namespace["set_cloud_auth_status_from_ack"],
        log=namespace["_cloud_log"],
        safe_preview=namespace["_cloud_safe_preview"],
        timeout=timeout,
        monotonic=namespace.get("time", time).monotonic,
    )
    if namespace.get("cloud_device_authorized"):
        namespace["cloud_authenticated_secret"] = namespace.get("CLOUD_DEVICE_SECRET", "")
        namespace["cloud_authenticated_ws_url"] = namespace.get("CLOUD_WS_URL", "")
    return result


async def send_cloud_register_namespace_runtime(
    namespace,
    ws,
    *,
    send_runtime=send_cloud_register_runtime,
):
    serial_connected = False
    try:
        with namespace["serial_lock"]:
            serial_obj = namespace.get("serial_obj")
            serial_connected = bool(serial_obj is not None and serial_obj.is_open)
    except Exception:
        serial_connected = False
    try:
        serial_generation = max(0, int(namespace.get("serial_connection_generation", 0)))
    except (TypeError, ValueError):
        serial_generation = 0
    authenticated_url = str(namespace.get("cloud_authenticated_ws_url") or "").strip()
    current_url = str(namespace.get("CLOUD_WS_URL") or "").strip()
    previous_session_secret = (
        namespace.get("cloud_authenticated_secret", "")
        if authenticated_url and authenticated_url == current_url
        else ""
    )
    send_kwargs = dict(
        auto_upload=namespace["CLOUD_AUTO_UPLOAD"],
        build_payload=namespace["_cloud_build_register_payload"],
        timestamp=namespace["_cloud_now_ts"],
        identity_payload=namespace["_cloud_identity_payload"],
        secret=namespace["CLOUD_DEVICE_SECRET"],
        serial_port=namespace["PORT"],
        serial_baud=namespace["BAUD"],
        serial_mode=namespace["MODE"],
        runtime_imei=namespace["_cloud_runtime_imei"],
        log=namespace["_cloud_log"],
        previous_session_secret=previous_session_secret,
        serial_connected=serial_connected,
        control_available=serial_connected,
        serial_connection_generation=serial_generation,
    )
    try:
        signature = inspect.signature(send_runtime)
        accepts_kwargs = any(
            parameter.kind == inspect.Parameter.VAR_KEYWORD
            for parameter in signature.parameters.values()
        )
        if not accepts_kwargs:
            for key in ("serial_connected", "control_available", "serial_connection_generation"):
                if key not in signature.parameters:
                    send_kwargs.pop(key, None)
    except (TypeError, ValueError):
        pass
    result = await send_runtime(
        ws,
        **send_kwargs,
    )
    if result in (True, "sent"):
        namespace["cloud_last_channel_status"] = (ws, serial_generation, serial_connected)
    return result


async def send_cloud_unregister_namespace_runtime(
    namespace,
    ws,
    *,
    reason="hidden",
    send_runtime=send_cloud_unregister_runtime,
):
    return await send_runtime(
        ws,
        reason=reason,
        build_payload=namespace["_cloud_build_unregister_payload"],
        timestamp=namespace["_cloud_now_ts"],
        identity_payload=namespace["_cloud_identity_payload"],
        secret=namespace["CLOUD_DEVICE_SECRET"],
        serial_port=namespace["PORT"],
        serial_baud=namespace["BAUD"],
        serial_mode=namespace["MODE"],
        runtime_imei=namespace["_cloud_runtime_imei"],
        log=namespace["_cloud_log"],
    )


async def send_cloud_session_revoke_namespace_runtime(
    namespace,
    ws,
    *,
    reason="disconnect",
    send_runtime=send_cloud_session_revoke_runtime,
):
    return await send_runtime(
        ws,
        reason=reason,
        build_payload=namespace["_cloud_build_session_revoke_payload"],
        timestamp=namespace["_cloud_now_ts"],
        identity_payload=namespace["_cloud_identity_payload"],
        log=namespace["_cloud_log"],
    )


async def cloud_ws_main_namespace_runtime(
    namespace,
    url,
    reconnect_interval,
    *,
    ws_main_runtime=cloud_ws_main_app_runtime,
):
    return await ws_main_runtime(
        url,
        reconnect_interval,
        stop_event=namespace["cloud_stop_event"],
        runtime_imei=namespace["_cloud_runtime_imei"],
        request_cloud_device_imei=namespace["request_cloud_device_imei"],
        set_cloud_status=namespace["set_cloud_status"],
        log=namespace["_cloud_log"],
        connect=namespace["websockets"].connect,
        set_ws=lambda value: namespace.__setitem__("cloud_ws_conn", value),
        set_connected=lambda value: namespace.__setitem__("cloud_connected", bool(value)),
        set_authorized=lambda value: namespace.__setitem__("cloud_device_authorized", bool(value)),
        reset_serial_log_state=namespace["_reset_cloud_serial_log_state"],
        send_register=namespace["_cloud_send_register"],
        wait_login_ack=namespace["_cloud_wait_login_ack"],
        handle_message=namespace["_handle_cloud_message"],
        cloud_control_enabled=lambda: namespace["CLOUD_CONTROL_ENABLED"],
        monotonic=namespace.get("time", time).monotonic,
        schedule_pending_sms_events=namespace.get(
            "_schedule_cloud_sms_event_drain",
            lambda: None,
        ),
    )


def cloud_thread_main_namespace_runtime(
    namespace,
    url,
    reconnect_interval,
    *,
    thread_main_runtime=cloud_thread_main_runtime,
):
    return thread_main_runtime(
        url,
        reconnect_interval,
        lock=namespace["cloud_ws_lock"],
        set_loop=lambda value: namespace.__setitem__("cloud_ws_loop", value),
        get_thread=lambda: namespace["cloud_ws_thread"],
        set_thread=lambda value: namespace.__setitem__("cloud_ws_thread", value),
        run_main=namespace["_cloud_ws_main"],
        log=namespace["_cloud_log"],
    )
