import asyncio
import threading
import time

from sms_core.cloud_imei_runtime import (
    maybe_capture_cloud_device_imei_runtime,
    notify_cloud_identity_changed_runtime,
    request_cloud_device_imei_runtime,
    set_cloud_device_imei_runtime,
)
from sms_core.cloud_message_runtime import cloud_status_payload_runtime


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
    return set_imei_runtime(
        imei,
        current_imei=lambda: namespace["CLOUD_DEVICE_IMEI"],
        normalize_imei=namespace["_normalize_imei"],
        set_device_imei=lambda value: namespace.__setitem__("CLOUD_DEVICE_IMEI", value),
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
    return request_runtime(
        serial_lock=namespace["serial_lock"],
        get_serial=lambda: namespace["serial_obj"],
        write_command_result=namespace["write_serial_command_result"],
        set_query_deadline=lambda deadline: namespace.__setitem__("cloud_imei_query_deadline", deadline),
        cloud_log=namespace["_cloud_log"],
        monotonic=namespace.get("time", time).monotonic,
        push_serial_debug=namespace["_push_serial_debug"],
        thread_factory=namespace.get("threading", threading).Thread,
    )


def maybe_capture_cloud_device_imei_namespace_runtime(
    namespace,
    line,
    *,
    capture_runtime=maybe_capture_cloud_device_imei_runtime,
):
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
