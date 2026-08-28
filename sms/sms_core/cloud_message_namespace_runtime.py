import asyncio
import inspect
import threading
from datetime import datetime

from sms_core.cloud_message_runtime import (
    handle_cloud_message_runtime,
    send_cloud_call_event_runtime,
    send_cloud_serial_command_runtime,
    send_cloud_sms_event_runtime,
)
from sms_core.cloud_serial_log_runtime import (
    drain_cloud_serial_log_queue,
    reset_cloud_serial_log_state,
    schedule_cloud_serial_log_drain,
    send_cloud_serial_log_runtime,
)
from sms_core.cloud_sms_event_runtime import (
    clear_cloud_sms_event_state,
    drain_cloud_sms_event_queue,
    enqueue_cloud_sms_event_runtime,
    schedule_cloud_sms_event_drain,
)
from sms_core.serial_debug import build_own_number_commands
from sms_core.serial_sender import (
    AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT,
    DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR,
    DEFAULT_SMS_PDU_SEND_COORDINATOR,
    start_registered_serial_worker,
    write_serial_command_sequence_confirmed_locked,
    write_text_sms_pdu_locked,
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


def clear_cloud_sms_event_state_namespace_runtime(namespace):
    return _call_with_optional_log_error(
        clear_cloud_sms_event_state,
        namespace["CLOUD_SMS_EVENT_Q"],
        namespace["CLOUD_SMS_EVENT_DRAIN_STATE"],
        log_error=namespace.get("log_file_only"),
    )


async def drain_cloud_sms_event_queue_namespace_runtime(namespace, ws, *, generation=None):
    return await _call_with_optional_log_error(
        drain_cloud_sms_event_queue,
        ws,
        event_queue=namespace["CLOUD_SMS_EVENT_Q"],
        batch_size=namespace["CLOUD_SMS_EVENT_DRAIN_BATCH"],
        state=namespace["CLOUD_SMS_EVENT_DRAIN_STATE"],
        is_current_connection=lambda current_ws: current_ws is namespace["cloud_ws_conn"],
        is_connected=lambda: namespace["cloud_connected"],
        is_authorized=lambda: namespace["cloud_device_authorized"],
        send_payload=namespace["_cloud_send_payload"],
        log_error=namespace.get("log_file_only"),
        generation=generation,
    )


def schedule_cloud_sms_event_drain_namespace_runtime(namespace, loop=None, ws=None):
    loop = namespace["cloud_ws_loop"] if loop is None else loop
    ws = namespace["cloud_ws_conn"] if ws is None else ws
    if (
        loop is None
        or not loop.is_running()
        or ws is None
        or ws is not namespace["cloud_ws_conn"]
        or not namespace["cloud_connected"]
        or not namespace["cloud_device_authorized"]
        or namespace["CLOUD_SMS_EVENT_Q"].empty()
    ):
        return False
    return _call_with_optional_log_error(
        schedule_cloud_sms_event_drain,
        loop,
        ws,
        state=namespace["CLOUD_SMS_EVENT_DRAIN_STATE"],
        drain_coro_factory=lambda current_ws, generation: namespace[
            "_cloud_drain_sms_event_queue"
        ](current_ws, generation=generation),
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
    metadata=None,
    *,
    send_runtime=send_cloud_sms_event_runtime,
):
    return send_runtime(
        callback_head,
        full_msg,
        metadata=metadata,
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
        enabled=namespace.get("CLOUD_CONTROL_ENABLED", True),
        enqueue_payload=lambda payload, loop, ws, can_send: enqueue_cloud_sms_event_runtime(
            payload,
            event_queue=namespace["CLOUD_SMS_EVENT_Q"],
            can_send=can_send,
            loop=loop,
            ws=ws,
            schedule_drain=lambda next_loop, next_ws: namespace["_schedule_cloud_sms_event_drain"](
                next_loop,
                next_ws,
            ),
            log_error=namespace.get("log_file_only"),
            state=namespace.get("CLOUD_SMS_EVENT_DRAIN_STATE"),
            is_enabled=lambda: namespace.get("CLOUD_CONTROL_ENABLED", True),
        ),
    )


def send_cloud_call_event_namespace_runtime(
    namespace,
    caller,
    message,
    *,
    blocked=False,
    block_reason="",
    send_runtime=send_cloud_call_event_runtime,
):
    return send_runtime(
        caller,
        message,
        blocked=blocked,
        block_reason=block_reason,
        authorized=namespace["cloud_device_authorized"],
        get_loop=lambda: namespace["cloud_ws_loop"],
        get_ws=lambda: namespace["cloud_ws_conn"],
        is_connected=lambda: namespace["cloud_connected"],
        runtime_imei=namespace["_cloud_runtime_imei"],
        build_payload=namespace["_cloud_build_call_event_payload"],
        send_payload=namespace["_cloud_send_payload"],
        timestamp=namespace["_cloud_now_ts"],
        identity_payload=namespace["_cloud_identity_payload"],
        run_coroutine_threadsafe=namespace.get("asyncio", asyncio).run_coroutine_threadsafe,
        enabled=namespace.get("CLOUD_CONTROL_ENABLED", True),
        enqueue_payload=lambda payload, loop, ws, can_send: enqueue_cloud_sms_event_runtime(
            payload,
            event_queue=namespace["CLOUD_SMS_EVENT_Q"],
            can_send=can_send,
            loop=loop,
            ws=ws,
            schedule_drain=lambda next_loop, next_ws: namespace["_schedule_cloud_sms_event_drain"](
                next_loop,
                next_ws,
            ),
            log_error=namespace.get("log_file_only"),
            state=namespace.get("CLOUD_SMS_EVENT_DRAIN_STATE"),
            is_enabled=lambda: namespace.get("CLOUD_CONTROL_ENABLED", True),
        ),
    )


def send_cloud_serial_command_namespace_runtime(
    namespace,
    command,
    command_data=None,
    *,
    send_runtime=send_cloud_serial_command_runtime,
):
    sms_coordinator = namespace.get(
        "SMS_SEND_COORDINATOR",
        DEFAULT_SMS_PDU_SEND_COORDINATOR,
    )
    command_coordinator = namespace.get(
        "SERIAL_COMMAND_RESPONSE_COORDINATOR",
        DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR,
    )

    def execute_confirmed_command(_serial_obj, next_command):
        return write_serial_command_sequence_confirmed_locked(
            namespace["serial_lock"],
            lambda: namespace["serial_obj"],
            (next_command,),
            response_coordinator=command_coordinator,
            response_timeout=AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT,
        )

    result = send_runtime(
        command,
        command_meta=command_data,
        serial_lock=namespace["serial_lock"],
        get_serial=lambda: namespace["serial_obj"],
        write_command_result=execute_confirmed_command,
        push_serial_debug=namespace.get("_push_serial_debug"),
        port_ui=namespace.get("port_ui"),
        log=namespace["_cloud_log"],
        allow_sensitive_commands=namespace.get(
            "CLOUD_SENSITIVE_COMMAND_PERMISSIONS",
            {},
        ),
        send_sms_transaction=lambda phone, message: write_text_sms_pdu_locked(
            namespace["serial_lock"],
            lambda: namespace["serial_obj"],
            phone,
            message,
            push_debug=namespace.get("_push_serial_debug"),
            port_ui=namespace.get("port_ui"),
            response_coordinator=sms_coordinator,
            command_response_coordinator=command_coordinator,
        ),
        set_own_number_transaction=lambda phone: write_serial_command_sequence_confirmed_locked(
            namespace["serial_lock"],
            lambda: namespace["serial_obj"],
            build_own_number_commands(phone),
            response_coordinator=command_coordinator,
        ),
    )
    return result


def apply_cloud_modem_health_namespace_runtime(namespace, result):
    health = namespace.get("CLOUD_MODEM_HEALTH")
    if health is None or not isinstance(result, tuple) or len(result) < 2:
        return result
    ok, info = bool(result[0]), str(result[1] or "")
    health_snapshot = health.record(ok, info)
    if health_snapshot.get("request_reconnect"):
        namespace["_cloud_log"](
            "连续 AT 响应超时，Modem 无响应，正在安全重连串口",
            show_main=True,
        )
        try:
            namespace["safe_close_serial"]()
        finally:
            namespace["serial_wakeup_event"].set()
    metadata = {
        "modem_unresponsive": bool(health_snapshot.get("modem_unresponsive")),
        "consecutive_at_timeouts": int(health_snapshot.get("consecutive_at_timeouts") or 0),
    }
    return ok, info, metadata


async def run_registered_cloud_serial_command_namespace_runtime(
    namespace,
    loop,
    command,
    command_data=None,
    *,
    start_worker=start_registered_serial_worker,
):
    result_future = loop.create_future()

    def finish(result):
        if not result_future.done():
            result_future.set_result(
                apply_cloud_modem_health_namespace_runtime(namespace, result)
            )

    def worker():
        try:
            result = namespace["_cloud_send_serial_command"](command, command_data)
        except Exception as exc:
            result = (False, f"发送失败：{exc}")
        try:
            loop.call_soon_threadsafe(finish, result)
        except Exception:
            pass

    try:
        start_worker(
            "cloud_serial_command",
            worker,
            log_error=namespace.get("log_file_only"),
            thread_registry=namespace["SERIAL_COMMAND_THREAD_REGISTRY"],
            thread_factory=namespace.get("threading", threading).Thread,
        )
    except Exception as exc:
        return False, f"发送失败：{exc}"

    return await result_future


async def handle_cloud_message_namespace_runtime(
    namespace,
    ws,
    message,
    *,
    handle_runtime=handle_cloud_message_runtime,
):
    loop = asyncio.get_running_loop()

    def set_authorized(value):
        authorized = bool(value)
        namespace["cloud_device_authorized"] = authorized
        if authorized:
            namespace["cloud_authenticated_secret"] = namespace.get("CLOUD_DEVICE_SECRET", "")
            namespace["cloud_authenticated_ws_url"] = namespace.get("CLOUD_WS_URL", "")

    return await handle_runtime(
        message,
        is_authorized=lambda: namespace["cloud_device_authorized"],
        set_authorized=set_authorized,
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
        send_serial_command=lambda command, command_data=None: run_registered_cloud_serial_command_namespace_runtime(
            namespace,
            loop,
            command,
            command_data,
        ),
        show_window=namespace["show_window"],
        hide_window=namespace["hide_window"],
    )
