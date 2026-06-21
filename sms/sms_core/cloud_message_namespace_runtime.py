import asyncio
import inspect
from datetime import datetime

from sms_core.cloud_message_runtime import (
    handle_cloud_message_runtime,
    send_cloud_serial_command_runtime,
    send_cloud_sms_event_runtime,
)
from sms_core.cloud_serial_log_runtime import (
    drain_cloud_serial_log_queue,
    reset_cloud_serial_log_state,
    schedule_cloud_serial_log_drain,
    send_cloud_serial_log_runtime,
)


def _call_with_optional_log_error(func, *args, log_error=None, **kwargs):
    if log_error is not None:
        try:
            signature = inspect.signature(func)
            supports_log_error = (
                "log_error" in signature.parameters
                or any(param.kind == inspect.Parameter.VAR_KEYWORD for param in signature.parameters.values())
            )
        except (TypeError, ValueError):
            supports_log_error = True
        if supports_log_error:
            kwargs["log_error"] = log_error
    return func(*args, **kwargs)


def reset_cloud_serial_log_state_namespace_runtime(namespace):
    return _call_with_optional_log_error(
        reset_cloud_serial_log_state,
        namespace["CLOUD_SERIAL_LOG_Q"],
        namespace["CLOUD_SERIAL_LOG_DRAIN_STATE"],
        log_error=namespace.get("log_file_only"),
    )


async def drain_cloud_serial_log_queue_namespace_runtime(namespace, ws):
    return await _call_with_optional_log_error(
        drain_cloud_serial_log_queue,
        ws,
        log_queue=namespace["CLOUD_SERIAL_LOG_Q"],
        batch_size=namespace["CLOUD_SERIAL_LOG_DRAIN_BATCH"],
        state=namespace["CLOUD_SERIAL_LOG_DRAIN_STATE"],
        is_current_connection=lambda current_ws: current_ws is namespace["cloud_ws_conn"],
        is_connected=lambda: namespace["cloud_connected"],
        log_error=namespace.get("log_file_only"),
    )


def schedule_cloud_serial_log_drain_namespace_runtime(namespace, loop, ws):
    return _call_with_optional_log_error(
        schedule_cloud_serial_log_drain,
        loop,
        ws,
        state=namespace["CLOUD_SERIAL_LOG_DRAIN_STATE"],
        drain_coro_factory=lambda current_ws: namespace["_cloud_drain_serial_log_queue"](current_ws),
        log_error=namespace.get("log_file_only"),
    )


def send_cloud_serial_log_namespace_runtime(
    namespace,
    line,
    *,
    send_runtime=send_cloud_serial_log_runtime,
):
    def build_payload(text):
        return namespace["_cloud_build_serial_log_payload"](
            text,
            namespace["_cloud_now_ts"](),
            namespace["_cloud_identity_payload"](),
            namespace["PORT"],
            namespace["BAUD"],
        )

    return send_runtime(
        line,
        authorized=namespace["cloud_device_authorized"],
        get_loop=lambda: namespace["cloud_ws_loop"],
        get_ws=lambda: namespace["cloud_ws_conn"],
        is_connected=lambda: namespace["cloud_connected"],
        runtime_imei=namespace["_cloud_runtime_imei"],
        build_payload=build_payload,
        log_queue=namespace["CLOUD_SERIAL_LOG_Q"],
        schedule_drain=lambda loop, ws: namespace["_schedule_cloud_serial_log_drain"](loop, ws),
        log_error=namespace.get("log_file_only"),
    )


def send_cloud_sms_event_namespace_runtime(
    namespace,
    callback_head,
    full_msg,
    *,
    send_runtime=send_cloud_sms_event_runtime,
):
    return send_runtime(
        callback_head,
        full_msg,
        authorized=namespace["cloud_device_authorized"],
        get_loop=lambda: namespace["cloud_ws_loop"],
        get_ws=lambda: namespace["cloud_ws_conn"],
        is_connected=lambda: namespace["cloud_connected"],
        runtime_imei=namespace["_cloud_runtime_imei"],
        build_payload=namespace["_cloud_build_sms_event_payload"],
        send_payload=namespace["_cloud_send_payload"],
        timestamp=namespace["_cloud_now_ts"],
        identity_payload=namespace["_cloud_identity_payload"],
        run_coroutine_threadsafe=namespace.get("asyncio", asyncio).run_coroutine_threadsafe,
    )


def send_cloud_serial_command_namespace_runtime(
    namespace,
    command,
    command_data=None,
    *,
    send_runtime=send_cloud_serial_command_runtime,
):
    return send_runtime(
        command,
        command_meta=command_data,
        serial_lock=namespace["serial_lock"],
        get_serial=lambda: namespace["serial_obj"],
        write_command_result=namespace["write_serial_command_result"],
        push_serial_debug=namespace.get("_push_serial_debug"),
        port_ui=namespace.get("port_ui"),
        log=namespace["_cloud_log"],
    )


async def handle_cloud_message_namespace_runtime(
    namespace,
    ws,
    message,
    *,
    handle_runtime=handle_cloud_message_runtime,
):
    loop = asyncio.get_running_loop()
    return await handle_runtime(
        message,
        is_authorized=lambda: namespace["cloud_device_authorized"],
        set_authorized=lambda value: namespace.__setitem__("cloud_device_authorized", bool(value)),
        reply=lambda payload: namespace["_cloud_reply"](ws, payload),
        log=namespace["_cloud_log"],
        set_auth_status_from_ack=namespace["set_cloud_auth_status_from_ack"],
        set_cloud_status=namespace["set_cloud_status"],
        check_replay_window=lambda data, mark_seen=True: namespace["_cloud_check_replay_window"](
            ws,
            data,
            mark_seen=mark_seen,
        ),
        auth_matches=namespace["_cloud_auth_matches"],
        time_text=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        timestamp=namespace["_cloud_now_ts"],
        status_payload=namespace["_cloud_send_status_payload"],
        send_serial_command=lambda command, command_data=None: loop.run_in_executor(
            None,
            namespace["_cloud_send_serial_command"],
            command,
            command_data,
        ),
        show_window=namespace["show_window"],
        hide_window=namespace["hide_window"],
    )
