import asyncio
import threading
import time
from contextlib import nullcontext
from sms_core.cloud_command_context import set_cloud_command_authorized

from sms_core.cloud_imei_runtime import (
    IMEI_QUERY_WINDOW_SECONDS,
    maybe_capture_cloud_device_imei_runtime,
    notify_cloud_identity_changed_runtime,
    request_cloud_device_imei_runtime,
    set_cloud_device_imei_runtime,
)
from sms_core.cloud_message_runtime import cloud_status_payload_runtime
from sms_core.cloud_connection_runtime import schedule_cloud_payload_runtime
from sms_core.cloud_payloads import build_channel_status_payload
from sms_core.cloud_sms_event_runtime import interrupt_cloud_sms_event_drain
from sms_core import serial_sender


def cloud_runtime_imei_namespace_runtime(namespace):
    if not namespace["cloud_imei_verified"]:
        return ""
    return namespace["_normalize_imei"](namespace["CLOUD_DEVICE_IMEI"])


def cloud_identity_payload_namespace_runtime(namespace):
    return namespace["_cloud_identity_payload_core"](namespace["_cloud_runtime_imei"](), namespace["APP_VERSION"])


def notify_cloud_identity_changed_namespace_runtime(
    namespace,
    *,
    notify_runtime=notify_cloud_identity_changed_runtime,
):
    return notify_runtime(
        get_loop=lambda: namespace["cloud_ws_loop"],
        get_ws=lambda: namespace["cloud_ws_conn"],
        is_connected=lambda: namespace["cloud_connected"],
        runtime_imei=namespace["_cloud_runtime_imei"],
        send_register=namespace["_cloud_send_register"],
        run_coroutine_threadsafe=namespace.get("asyncio", asyncio).run_coroutine_threadsafe,
    )


def set_cloud_device_imei_namespace_runtime(
    namespace,
    imei,
    *,
    source="",
    set_imei_runtime=set_cloud_device_imei_runtime,
):
    def set_device_imei(value):
        event_state = namespace.get("CLOUD_SMS_EVENT_DRAIN_STATE")
        with event_state.lock if event_state is not None else nullcontext():
            # Authorization belongs to the previous identity until its new
            # login ACK arrives. Keep queued events and their generation.
            set_cloud_command_authorized(namespace, False)
            namespace["CLOUD_DEVICE_IMEI"] = value
            interrupt_cloud_sms_event_drain(event_state, lock_held=True)
        reset_logs = namespace.get("_reset_cloud_serial_log_state")
        if callable(reset_logs):
            reset_logs()
        set_status = namespace.get("set_cloud_status")
        if namespace.get("cloud_connected") and callable(set_status):
            set_status("🌐 等待授权" if value else "🌐 等待读取IMEI", "#b26a00")

    return set_imei_runtime(
        imei,
        current_imei=lambda: namespace["CLOUD_DEVICE_IMEI"],
        normalize_imei=namespace["_normalize_imei"],
        set_device_imei=set_device_imei,
        set_verified=lambda value: namespace.__setitem__("cloud_imei_verified", bool(value)),
        log=namespace["_cloud_log"],
        notify_identity_changed=namespace["_notify_cloud_identity_changed"],
        source=source,
    )


def request_cloud_device_imei_namespace_runtime(
    namespace,
    *,
    request_runtime=request_cloud_device_imei_runtime,
):
    def confirmed_imei_query():
        with serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK:
            with namespace["serial_lock"]:
                serial_obj = namespace.get("serial_obj")
                generation = namespace.get("serial_connection_generation", 0)
            attempted = False
            query_deadline = None
            responses = []
            window_owner = object()
            invalidated = False
            unconfirmed_error = "IMEI查询结果未获确认，已暂停此串口发送，请重启设备后重新连接"

            def get_original_serial():
                nonlocal invalidated
                invalidated = invalidated or bool(
                    namespace.get("serial_obj") is not serial_obj
                    or namespace.get("serial_connection_generation", 0) != generation
                    or namespace.get("serial_running", True) is False
                    or namespace.get("is_exiting")
                    or any(event is not None and event.is_set() for event in (
                        namespace.get("serial_stop_event"), namespace.get("TK_SHUTDOWN"),
                    ))
                )
                return None if invalidated else serial_obj

            def before_write():
                nonlocal attempted, query_deadline
                if get_original_serial() is not serial_obj or serial_obj is None:
                    return False
                query_deadline = namespace.get("time", time).monotonic() + IMEI_QUERY_WINDOW_SECONDS
                namespace["_cloud_imei_query_owner"] = window_owner
                namespace["cloud_imei_query_deadline"] = query_deadline
                attempted = True
                return True

            try:
                result = serial_sender.write_serial_command_confirmed_locked(
                    namespace["serial_lock"], get_original_serial, "AT+CGSN",
                    response_coordinator=namespace.get(
                        "SERIAL_COMMAND_RESPONSE_COORDINATOR",
                        serial_sender.DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR,
                    ),
                    response_timeout=IMEI_QUERY_WINDOW_SECONDS,
                    before_write=before_write, on_modem_response=responses.append,
                    isolate_unconfirmed=True, unconfirmed_error=unconfirmed_error,
                )
            except Exception as exc:
                result = serial_sender.SerialCommandResult(False, f"读取IMEI失败：{type(exc).__name__}")
            finally:
                with namespace["serial_lock"], serial_sender.DEFAULT_SERIAL_WRITE_LOCK:
                    if (namespace.get("_cloud_imei_query_owner") is window_owner
                            and namespace.get("cloud_imei_query_deadline") == query_deadline):
                        namespace["cloud_imei_query_deadline"] = 0.0
                    if namespace.get("_cloud_imei_query_owner") is window_owner:
                        namespace.pop("_cloud_imei_query_owner", None)
                    # Keep the transaction owned until an uncertain query can
                    # no longer lend its delayed terminal to the next command.
                    if attempted and not any(getattr(response, "line", "") for response in responses):
                        serial_obj._sms_recovery_error = unconfirmed_error
            recovery_error = serial_sender.serial_sms_recovery_error(serial_obj)
            if not result.ok and recovery_error and recovery_error not in result.error:
                return serial_sender.SerialCommandResult(False, f"{result.error}；{recovery_error}")
            return result

    return request_runtime(
        serial_lock=namespace["serial_lock"],
        get_serial=lambda: namespace["serial_obj"],
        write_command_result=namespace["write_serial_command_result"],
        set_query_deadline=lambda deadline: namespace.__setitem__("cloud_imei_query_deadline", deadline),
        get_query_deadline=lambda: namespace["cloud_imei_query_deadline"],
        cloud_log=namespace["_cloud_log"],
        monotonic=namespace.get("time", time).monotonic,
        push_serial_debug=namespace["_push_serial_debug"],
        thread_factory=namespace.get("threading", threading).Thread,
        thread_registry=namespace["SERIAL_COMMAND_THREAD_REGISTRY"],
        confirmed_imei_query=confirmed_imei_query,
    )


def maybe_capture_cloud_device_imei_namespace_runtime(
    namespace,
    line,
    *,
    capture_runtime=maybe_capture_cloud_device_imei_runtime,
):
    # Keep reading and consuming the window atomic with query creation or
    # rollback. A stale capture must not clear a newer request's deadline.
    with namespace["serial_lock"]:
        return capture_runtime(
            line,
            query_deadline=namespace["cloud_imei_query_deadline"],
            set_query_deadline=lambda value: namespace.__setitem__("cloud_imei_query_deadline", value),
            imei_regex=namespace["IMEI_REGEX"],
            set_device_imei=namespace["_set_cloud_device_imei"],
            monotonic=namespace.get("time", time).monotonic,
        )


def cloud_auth_matches_namespace_runtime(namespace, data):
    ok, reason = namespace["_cloud_auth_match_result"](
        data,
        namespace["_cloud_runtime_imei"](),
        namespace["CLOUD_DEVICE_SECRET"],
    )
    if not ok:
        namespace["_cloud_log"](reason)
    return ok


def cloud_status_payload_namespace_runtime(
    namespace,
    *,
    status_runtime=cloud_status_payload_runtime,
):
    return status_runtime(
        serial_lock=namespace["serial_lock"],
        get_serial=lambda: namespace["serial_obj"],
        build_payload=namespace["_cloud_build_status_payload"],
        timestamp=namespace["_cloud_now_ts"],
        identity_payload=namespace["_cloud_identity_payload"],
        cloud_connected=namespace["cloud_connected"],
        serial_port=namespace["PORT"],
        serial_baud=namespace["BAUD"],
        serial_mode=namespace["MODE"],
    )


def notify_cloud_channel_status_namespace_runtime(namespace, serial_connected=None):
    """Publish the desktop transport's serial availability to the server."""
    if serial_connected is None:
        try:
            with namespace["serial_lock"]:
                serial_obj = namespace.get("serial_obj")
                serial_connected = bool(serial_obj is not None and serial_obj.is_open)
        except Exception:
            serial_connected = False
    serial_connected = bool(serial_connected)
    try:
        generation = max(0, int(namespace.get("serial_connection_generation", 0)))
    except (TypeError, ValueError):
        generation = 0
    ws = namespace.get("cloud_ws_conn")
    state_key = (ws, generation, serial_connected)
    if namespace.get("cloud_last_channel_status") == state_key:
        return False
    if not namespace.get("cloud_connected") or not namespace.get("cloud_device_authorized"):
        return False
    build_payload = namespace.get("_cloud_build_channel_status_payload", build_channel_status_payload)
    payload = build_payload(
        namespace["_cloud_now_ts"](),
        namespace["_cloud_identity_payload"](),
        serial_connected,
        serial_connected,
        generation,
    )
    scheduled = schedule_cloud_payload_runtime(
        payload,
        get_loop=lambda: namespace.get("cloud_ws_loop"),
        get_ws=lambda: namespace.get("cloud_ws_conn"),
        is_connected=lambda: namespace.get("cloud_connected") and namespace.get("cloud_device_authorized"),
        send_payload=namespace["_cloud_send_payload"],
        run_coroutine_threadsafe=namespace.get("asyncio", asyncio).run_coroutine_threadsafe,
        log_error=namespace.get("log_file_only"),
    )
    if scheduled:
        namespace["cloud_last_channel_status"] = state_key
    return scheduled


async def cloud_check_replay_window_namespace_runtime(namespace, ws, data, *, mark_seen=True):
    result = namespace["_cloud_check_replay_window_core"](
        data,
        namespace["cloud_replay_seen"],
        namespace["_cloud_now_ts"](),
        namespace["CLOUD_REPLAY_WINDOW_SECONDS"],
        namespace["CLOUD_REPLAY_CACHE_MAX"],
        mark_seen=mark_seen,
    )
    if result.ok:
        return True
    if result.log_message:
        namespace["_cloud_log"](result.log_message)
    if result.payload is not None:
        await namespace["_cloud_reply"](ws, result.payload)
    return False
