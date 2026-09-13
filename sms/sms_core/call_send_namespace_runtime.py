"""Record only modem-confirmed voice calls sent from the local serial editor."""
import re
import time
import uuid

from sms_core.cloud_payloads import build_outgoing_call_event_payload
from sms_core.cloud_sms_event_runtime import enqueue_cloud_sms_event_runtime
from sms_core.serial_sender import (
    AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT,
    CALL_TERMINAL_RESPONSE_RE,
    DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR,
    DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY,
    send_command_async,
    start_registered_serial_worker,
    write_serial_command_confirmed_locked,
)


VOICE_DIAL_RE = re.compile(r"[ \t]*ATD(\+?[0-9]+);?[ \t]*(?:\r\n|\r|\n)?\Z", re.IGNORECASE)


def send_local_serial_command_namespace_runtime(
    namespace, command, append_crlf=True, *,
    start_worker=start_registered_serial_worker,
    send_raw=send_command_async,
    response_timeout=AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT,
):
    raw_command = str(command or "")
    match = VOICE_DIAL_RE.fullmatch(raw_command)
    terminated = bool(append_crlf or raw_command.endswith(("\r", "\n")))
    if not match or not terminated:
        # Preserve arbitrary AT, USSD/MMI and unterminated raw editor writes.
        return send_raw(
            namespace["serial_lock"], lambda: namespace["serial_obj"], command,
            append_crlf=append_crlf, push_debug=namespace["_push_serial_debug"],
            log_error=namespace.get("log_file_only"),
        )

    with namespace["serial_lock"]:
        serial = namespace.get("serial_obj")
        generation = namespace.get("serial_connection_generation", 0)
        identity = dict(namespace["_cloud_identity_payload"]())
        imei = str(identity.get("imei") or "").strip()
    source_id = "desktop_dial_" + uuid.uuid4().hex
    attempted_at = None
    context_changed = False

    def get_captured_serial():
        nonlocal attempted_at, context_changed
        stopped = any(
            event is not None and event.is_set()
            for event in (namespace.get("serial_stop_event"), namespace.get("TK_SHUTDOWN"))
        )
        if (
            stopped or namespace.get("serial_running", True) is False
            or namespace.get("serial_obj") is not serial
            or namespace.get("serial_connection_generation", 0) != generation
            or (imei and str(namespace["_cloud_runtime_imei"]() or "") != imei)
        ):
            context_changed = True
            return None
        if attempted_at is None:
            attempted_at = int(namespace["_cloud_now_ts"]())
        return serial

    def record_attempt():
        if not namespace.get("CLOUD_CONTROL_ENABLED", True):
            return "disabled"
        if not imei:
            namespace["port_ui"]("拨号已执行，但设备身份尚未就绪，未上传呼出记录", "normal")
            return "missing_imei"
        payload = build_outgoing_call_event_payload(
            match.group(1), attempted_at, identity, source_id,
            time_text=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(attempted_at)),
        )
        loop, ws = namespace.get("cloud_ws_loop"), namespace.get("cloud_ws_conn")
        can_send = bool(
            loop is not None and loop.is_running() and ws is not None
            and namespace.get("cloud_connected") and namespace.get("cloud_device_authorized")
            and str(namespace["_cloud_runtime_imei"]() or "") == imei
        )
        return enqueue_cloud_sms_event_runtime(
            payload, event_queue=namespace["CLOUD_SMS_EVENT_Q"], can_send=can_send,
            loop=loop, ws=ws, schedule_drain=namespace["_schedule_cloud_sms_event_drain"],
            log_error=namespace.get("log_file_only"),
            state=namespace.get("CLOUD_SMS_EVENT_DRAIN_STATE"),
            is_enabled=lambda: namespace.get("CLOUD_CONTROL_ENABLED", True),
        )

    def worker():
        replies = []
        try:
            result = write_serial_command_confirmed_locked(
                namespace["serial_lock"], get_captured_serial, raw_command,
                response_coordinator=namespace.get(
                    "SERIAL_COMMAND_RESPONSE_COORDINATOR", DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR,
                ),
                append_crlf=append_crlf, response_timeout=response_timeout,
                push_debug=namespace["_push_serial_debug"], on_modem_response=replies.append,
            )
        except Exception:
            namespace["port_ui"]("拨号执行异常，未生成呼出记录，请检查串口状态", "normal")
            return
        # A serial-write error could contain any text. Only the waiter's
        # trusted modem frame may prove an unanswered/busy dial was attempted.
        response = replies[0] if replies else None
        confirmed = bool(response is not None and response.line and (
            response.ok is True or CALL_TERMINAL_RESPONSE_RE.fullmatch(response.error.strip())
        ))
        if not confirmed:
            message = (
                "串口连接或设备身份已变化，已取消旧拨号"
                if context_changed else "拨号未获 Modem 确认，未生成呼出记录"
            )
            if not result.ok:
                namespace["port_ui"](message, "normal")
            return
        try:
            recorded = record_attempt()
            if recorded == "queue_full":
                raise RuntimeError("event queue full")
        except Exception:
            namespace["port_ui"]("拨号已执行，但呼出记录上传遇到问题；请勿为补记录重复拨号", "normal")

    return start_worker(
        "serial_local_dial", worker, log_error=namespace.get("log_file_only"),
        thread_registry=namespace.get("SERIAL_COMMAND_THREAD_REGISTRY", DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY),
    )
