"""Record only modem-confirmed voice calls sent from the local serial editor."""
import re
import time
import uuid

from sms_core.cloud_payloads import build_outgoing_call_event_payload
from sms_core.call_command_context import (
    SerialCallStateSnapshot, call_state_lock, call_with_state_snapshot,
    capture_call_command_context, invalidate_call_commands,
)
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
VOICE_CONTROL_RE = re.compile(r"[ \t]*AT[AH];?[ \t]*(?:\r\n|\r|\n)?\Z", re.IGNORECASE)


def _send_local_call_control(namespace, command, append_crlf, start_worker, response_timeout,
                             dial_popup, on_result):
    serial = namespace.get("serial_obj")
    generation = namespace.get("serial_connection_generation", 0)
    with call_state_lock(namespace):
        context = capture_call_command_context(namespace, command)
        request = namespace.get("_local_dial_request")
        if dial_popup is not None and (
            request is None or request.popup is not dial_popup
            or namespace.get("_dial_state_generation") is not request
            or request.finished
        ):
            return None
        local_hangup = (context.command.upper().rstrip(";") == "ATH" and request is not None
                       and (dial_popup is not None or namespace.get("current_dial_num") == request.number))
        if local_hangup and request is namespace.get("_dial_state_generation"):
            if request.finished or (request.command_context is not None
                                    and not request.command_context.is_current()):
                return None
        else:
            request = None
        if request is not None and request.serial_context is not None:
            serial, generation = request.serial_context

    def deliver(ok, result_context):
        if on_result is None:
            return
        def complete():
            on_result(ok, result_context.is_current())
        post = namespace.get("run_on_ui_thread")
        if callable(post):
            post(complete, namespace.get("ui_post"))
        else:
            complete()

    with call_state_lock(namespace):
        if not context.is_current():
            return None
        if request is not None and not request.submitted:
            request.finished = True
            namespace["current_dial_num"] = ""
            invalidate_call_commands(namespace)
            completion_context = capture_call_command_context(namespace, command)
        else:
            completion_context = None
    if completion_context is not None:
        deliver(True, completion_context)
        return None

    with namespace["serial_lock"]:
        captured_connection_current = (
            namespace.get("serial_obj") is serial
            and namespace.get("serial_connection_generation", 0) == generation)
    if not captured_connection_current:
        namespace["port_ui"]("串口连接已变化，已取消旧通话指令", "normal")
        deliver(False, context)
        return None

    def get_serial():
        stopped = any(event is not None and event.is_set() for event in (
            namespace.get("serial_stop_event"), namespace.get("TK_SHUTDOWN")))
        if (stopped or namespace.get("serial_running", True) is False
                or namespace.get("serial_obj") is not serial
                or namespace.get("serial_connection_generation", 0) != generation):
            return None
        return serial

    def worker():
        try:
            result = write_serial_command_confirmed_locked(
                namespace["serial_lock"], get_serial, command,
                response_coordinator=namespace.get(
                    "SERIAL_COMMAND_RESPONSE_COORDINATOR", DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR),
                append_crlf=append_crlf, response_timeout=response_timeout,
                push_debug=namespace["_push_serial_debug"],
                write_guard=context.write_guard, before_write=context.before_write,
            )
            with call_state_lock(namespace):
                result_context = context
                if result.ok and context.is_current() and context.command.upper().rstrip(";") == "ATH":
                    if request is not None and namespace.get("_dial_state_generation") is request:
                        request.finished = True
                    namespace["current_dial_num"] = ""
                    invalidate_call_commands(namespace)
                    result_context = capture_call_command_context(namespace, command)
            if not result.ok:
                namespace["port_ui"]("通话指令未完成，通话或串口可能已变化，请检查当前状态", "normal")
            deliver(result.ok, result_context)
        except Exception:
            namespace["port_ui"]("通话指令执行异常，请检查串口状态", "normal")
            deliver(False, context)

    try:
        return start_worker("serial_local_call_control", worker,
            log_error=namespace.get("log_file_only"), thread_registry=namespace.get(
                "SERIAL_COMMAND_THREAD_REGISTRY", DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY))
    except Exception:
        namespace["port_ui"]("通话指令任务未能启动，请稍后重试", "normal")
        deliver(False, context)
        return None


def send_local_serial_command_namespace_runtime(
    namespace, command, append_crlf=True, *,
    start_worker=start_registered_serial_worker,
    send_raw=send_command_async,
    response_timeout=AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT,
    dial_popup=None,
    on_result=None,
):
    raw_command = str(command or "")
    match = VOICE_DIAL_RE.fullmatch(raw_command)
    terminated = bool(append_crlf or raw_command.endswith(("\r", "\n")))
    if terminated and VOICE_CONTROL_RE.fullmatch(raw_command):
        return _send_local_call_control(namespace, raw_command, append_crlf, start_worker,
                                        response_timeout, dial_popup, on_result)
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
    with call_state_lock(namespace):
        call_context = capture_call_command_context(namespace, raw_command)
        request = namespace.get("_local_dial_request")
        if (request is not None and request is namespace.get("_dial_state_generation")
                and not request.queued and not request.finished
                and request.number == match.group(1)):
            request.queued = True
            call_context.local_dial_request = request
            request.command_context = call_context
            request.serial_context = (serial, generation)
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

    def finish_failed_popup():
        callback = namespace.get("finish_dial_call")
        with call_state_lock(namespace):
            if (not callable(callback) or call_context.local_dial_request is None
                    or namespace.get("_dial_state_generation") is not call_context.local_dial_request):
                return
            snapshot = SerialCallStateSnapshot(namespace)
        call_with_state_snapshot(
            callback, "📞 拨号失败，请检查串口状态", call_state_snapshot=snapshot,
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
                write_guard=call_context.write_guard,
                before_write=call_context.before_write,
            )
            call_context.finish(result.ok)
            if not result.ok:
                finish_failed_popup()
        except Exception:
            call_context.finish(False)
            finish_failed_popup()
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

    try:
        return start_worker(
            "serial_local_dial", worker, log_error=namespace.get("log_file_only"),
            thread_registry=namespace.get("SERIAL_COMMAND_THREAD_REGISTRY", DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY),
        )
    except Exception:
        call_context.finish(False)
        finish_failed_popup()
        namespace["port_ui"]("拨号任务未能启动，请稍后重试", "normal")
        return None
