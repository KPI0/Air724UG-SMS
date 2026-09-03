import hashlib
import hmac
import json
import re

from sms_core.cloud_command_security import (
    CLOUD_SEND_SMS_TRANSACTION_COMMAND,
    CLOUD_SET_OWN_NUMBER_TRANSACTION_COMMAND,
    cloud_command_batch_error,
    cloud_command_control_char_error,
    cloud_sensitive_command_block_message,
    is_sensitive_cloud_command_allowed,
    sensitive_cloud_command_decision,
)
from sms_core.cloud_messages import (
    attach_cloud_task_ids,
    cloud_auth_failed_payload,
    cloud_command_started_payload,
    cloud_unauthorized_payload,
    dispatch_cloud_action,
    is_cloud_auth_ack_type,
    parse_cloud_message,
)
from sms_core.cloud_protocol import auth_status_from_ack
from sms_core.cloud_security import (
    HIDDEN_SENSITIVE_COMMAND,
    HIDDEN_SMS_COMMAND,
    HIDDEN_SMS_META,
    safe_preview,
)


CLOUD_SMS_PHONE_RE = re.compile(r"(?:\+\d{7,15}|\d{3,20})\Z")
CLOUD_OWN_NUMBER_RE = re.compile(r"\+[1-9]\d{6,14}\Z")
CLOUD_SMS_MESSAGE_MAX_LENGTH = 70
DEVICE_SESSION_REVOKE_PROOF_CONTEXT = b"air724ug-sms:device-session-revoke:v1"


def cloud_session_revoke_proof(secret, imei):
    secret_bytes = str(secret or "").strip().encode("utf-8")
    imei_bytes = str(imei or "").strip().encode("ascii", "ignore")
    if not secret_bytes or not imei_bytes:
        return ""
    return hmac.new(
        secret_bytes,
        DEVICE_SESSION_REVOKE_PROOF_CONTEXT + b":" + imei_bytes,
        hashlib.sha256,
    ).hexdigest()


def _cloud_transaction_parameters(command, command_meta):
    meta = command_meta if isinstance(command_meta, dict) else {}
    normalized_command = str(command or "").strip().upper()
    if normalized_command == CLOUD_SEND_SMS_TRANSACTION_COMMAND:
        phone = str(meta.get("sms_phone") or "").strip()
        message = str(meta.get("sms_message") or "")
        if not CLOUD_SMS_PHONE_RE.fullmatch(phone):
            return None, "短信接收号码格式不正确"
        if len(message) > CLOUD_SMS_MESSAGE_MAX_LENGTH:
            return None, f"短信正文不能超过 {CLOUD_SMS_MESSAGE_MAX_LENGTH} 个字符"
        return (phone, message), ""
    if normalized_command == CLOUD_SET_OWN_NUMBER_TRANSACTION_COMMAND:
        phone = str(meta.get("own_number") or "").strip()
        if not CLOUD_OWN_NUMBER_RE.fullmatch(phone):
            return None, "本机号码格式不正确"
        return (phone,), ""
    return (), ""


def _cloud_transaction_result(result, fallback_error):
    if hasattr(result, "ok"):
        return bool(result.ok), str(getattr(result, "error", "") or fallback_error)
    return bool(result), "" if result else fallback_error


async def send_cloud_register_runtime(
    ws,
    *,
    auto_upload,
    build_payload,
    timestamp,
    identity_payload,
    secret,
    serial_port,
    serial_baud,
    serial_mode,
    runtime_imei,
    log,
    serialize_payload=None,
    previous_session_secret="",
    serial_connected=None,
    control_available=None,
    serial_connection_generation=None,
):
    payload = build_payload(
        auto_upload,
        timestamp(),
        identity_payload(),
        secret,
        serial_port,
        serial_baud,
        serial_mode,
    )
    if isinstance(payload, dict):
        if serial_connected is not None:
            payload["serial_connected"] = bool(serial_connected)
        if control_available is not None:
            payload["control_available"] = bool(control_available)
        if serial_connection_generation is not None:
            try:
                payload["serial_connection_generation"] = max(0, int(serial_connection_generation))
            except (TypeError, ValueError):
                payload["serial_connection_generation"] = 0
        previous_secret = str(previous_session_secret or "").strip()
        current_secret = str(secret or "").strip()
        if previous_secret and not hmac.compare_digest(previous_secret, current_secret):
            proof = cloud_session_revoke_proof(previous_secret, runtime_imei())
            if proof:
                payload["previous_session_proof"] = proof
    serialize_payload = serialize_payload or (lambda payload: json.dumps(payload, ensure_ascii=False))
    try:
        await ws.send(serialize_payload(payload))
        if auto_upload:
            log(f"已上报设备IMEI：{runtime_imei()}", show_main=True)
        else:
            log(f"隐身模式：已注册路由IMEI（不公开设备列表，日志继续上传）：{runtime_imei()}", show_main=True)
        return "sent"
    except Exception as exc:
        log(f"上报设备身份失败：{exc}")
        return "error"


async def send_cloud_unregister_runtime(
    ws,
    *,
    reason="hidden",
    build_payload,
    timestamp,
    identity_payload,
    secret,
    serial_port,
    serial_baud,
    serial_mode,
    runtime_imei,
    log,
    serialize_payload=None,
):
    payload = build_payload(
        reason,
        timestamp(),
        identity_payload(),
        secret,
        serial_port,
        serial_baud,
        serial_mode,
    )
    serialize_payload = serialize_payload or (lambda payload: json.dumps(payload, ensure_ascii=False))
    try:
        await ws.send(serialize_payload(payload))
        log(f"已通知云端设备离线：{runtime_imei()}")
        return "sent"
    except Exception as exc:
        log(f"通知云端设备离线失败：{exc}")
        return "error"


async def send_cloud_session_revoke_runtime(
    ws,
    *,
    reason="disconnect",
    build_payload,
    timestamp,
    identity_payload,
    log,
    serialize_payload=None,
):
    payload = build_payload(
        reason,
        timestamp(),
        identity_payload(),
    )
    serialize_payload = serialize_payload or (lambda payload: json.dumps(payload, ensure_ascii=False))
    try:
        await ws.send(serialize_payload(payload))
        return "sent"
    except Exception as exc:
        log(f"撤销云端设备会话失败：{exc}")
        return "error"


def send_cloud_serial_command_runtime(
    command,
    *,
    command_meta=None,
    serial_lock,
    get_serial,
    write_command_result,
    push_serial_debug,
    port_ui=None,
    log,
    allow_sensitive_commands=False,
    send_sms_transaction=None,
    set_own_number_transaction=None,
):
    raw_command = str(command or "")
    batch_error = cloud_command_batch_error(raw_command)
    if batch_error:
        return False, batch_error
    control_error = cloud_command_control_char_error(raw_command)
    if control_error:
        return False, control_error

    cmd = raw_command.strip()
    if not cmd:
        return False, "AT 指令不能为空"

    sensitive_decision = sensitive_cloud_command_decision(cmd, command_meta)
    sensitive_reason = sensitive_decision.reason
    try:
        display_cmd = _cloud_command_display_text(cmd, command_meta)
        if sensitive_reason and not is_sensitive_cloud_command_allowed(
            sensitive_decision,
            allow_sensitive_commands,
        ):
            info = cloud_sensitive_command_block_message(sensitive_reason)
            try:
                log(info, show_main=True)
            except TypeError:
                log(info)
            return False, info

        transaction_parameters, transaction_error = _cloud_transaction_parameters(
            cmd,
            command_meta,
        )
        if transaction_error:
            return False, transaction_error
        normalized_command = cmd.upper()
        if normalized_command == CLOUD_SEND_SMS_TRANSACTION_COMMAND:
            if not callable(send_sms_transaction):
                return False, "客户端未配置云端短信事务执行器"
            ok, error = _cloud_transaction_result(
                send_sms_transaction(*transaction_parameters),
                "短信发送失败，Modem 未确认发送成功",
            )
            if not ok:
                return False, error
            log("云端短信事务执行成功")
            return True, "短信发送成功"
        if normalized_command == CLOUD_SET_OWN_NUMBER_TRANSACTION_COMMAND:
            if not callable(set_own_number_transaction):
                return False, "客户端未配置本机号码修改事务执行器"
            ok, error = _cloud_transaction_result(
                set_own_number_transaction(*transaction_parameters),
                "修改本机号码失败，事务已停止",
            )
            if not ok:
                return False, error
            log("云端本机号码修改事务执行成功")
            return True, "本机号码修改成功"

        with serial_lock:
            serial_obj = get_serial()
        result = write_command_result(serial_obj, cmd)
        if not result.ok:
            if sensitive_reason:
                return False, "敏感指令执行失败（指令内容已隐藏，Modem 未确认成功）"
            return False, result.error

        try:
            if push_serial_debug:
                push_serial_debug(f">>> 云端发送: {display_cmd}\\r\\n")
        except Exception:
            pass

        try:
            if port_ui:
                _log_cloud_command_to_port(port_ui, cmd, command_meta)
        except Exception:
            pass

        log(f"云端指令执行成功：{display_cmd}")
        return True, f"执行成功：{display_cmd}"
    except Exception as exc:
        if sensitive_reason:
            return False, f"敏感指令执行失败（指令内容已隐藏）：{type(exc).__name__}"
        return False, f"执行失败：{exc}"


def _cloud_command_display_text(cmd, command_meta):
    if cloud_command_batch_error(cmd):
        return "云端 AT 指令（已拒绝：只允许单条指令）"
    meta = command_meta if isinstance(command_meta, dict) else {}
    decision = sensitive_cloud_command_decision(cmd, meta)
    if (
        decision.category == "sms"
        and str(meta.get("sms_log") or "").strip().lower() == "suppress"
    ):
        return HIDDEN_SMS_COMMAND
    reason = decision.reason
    if reason:
        return f"敏感指令（{reason}，内容已隐藏）"
    return str(cmd or "").strip()


def cloud_incoming_preview(incoming):
    data = getattr(incoming, "data", None)
    if not isinstance(data, dict):
        return safe_preview(getattr(incoming, "raw", ""))

    masked = dict(data)
    sms_log = str(masked.get("sms_log") or "").strip().lower()
    command = next(
        (
            masked.get(key)
            for key in ("command", "data", "cmd")
            if isinstance(masked.get(key), (str, bytes)) and str(masked.get(key)).strip()
        ),
        "",
    )
    sensitive_decision = sensitive_cloud_command_decision(command, masked)
    sensitive_reason = sensitive_decision.reason
    if cloud_command_batch_error(command):
        for key in ("cmd", "command", "data"):
            if key in masked:
                masked[key] = "云端 AT 指令（已隐藏：包含多条指令）"
    elif cloud_command_control_char_error(command):
        for key in ("cmd", "command", "data"):
            if key in masked:
                masked[key] = "云端 AT 指令（已隐藏：包含不支持的控制字符）"
    elif sensitive_decision.category == "sms" and sms_log == "suppress":
        for key in ("cmd", "command", "data"):
            if key in masked:
                masked[key] = HIDDEN_SMS_COMMAND
    elif sensitive_reason:
        for key in ("cmd", "command", "data"):
            if key in masked:
                masked[key] = HIDDEN_SENSITIVE_COMMAND

    if sensitive_decision.category == "sms" and (
        sms_log in ("summary", "suppress")
        or str(masked.get("command_kind") or "") == "send_sms"
    ):
        for key in ("sms_phone", "sms_message"):
            if key in masked:
                masked[key] = HIDDEN_SMS_META
    if sensitive_decision.category == "phone_number" and "own_number" in masked:
        masked["own_number"] = HIDDEN_SENSITIVE_COMMAND

    return safe_preview(json.dumps(masked, ensure_ascii=False))


def _log_cloud_command_to_port(port_ui, cmd, command_meta):
    meta = command_meta if isinstance(command_meta, dict) else {}
    if sensitive_cloud_command_decision(cmd, meta).category != "sms":
        return
    sms_log = str(meta.get("sms_log") or "").strip().lower()
    if sms_log == "suppress":
        return
    if sms_log == "summary":
        phone = str(meta.get("sms_phone") or "").strip()
        message = str(meta.get("sms_message") or "")
        if phone:
            port_ui(f"云端发送短信至 {phone}：", "normal")
        else:
            port_ui("云端发送短信：", "normal")
        if message:
            port_ui(message, "sms")
        return


def send_cloud_sms_event_runtime(
    callback_head,
    full_msg,
    metadata=None,
    *,
    authorized,
    get_loop,
    get_ws,
    is_connected,
    runtime_imei,
    build_payload,
    send_payload,
    timestamp,
    identity_payload,
    run_coroutine_threadsafe,
    enabled=True,
    enqueue_payload=None,
):
    body = str(full_msg or "").strip()
    if not body:
        return "empty"
    if not enabled:
        return "disabled"
    if not authorized and enqueue_payload is None:
        return "unauthorized"
    send_coro = None
    try:
        loop = get_loop()
        ws = get_ws()
        can_send = bool(
            authorized
            and loop is not None
            and loop.is_running()
            and ws is not None
            and is_connected()
        )
        if not can_send and enqueue_payload is None:
            return "not_connected"
        if not runtime_imei():
            return "missing_imei"
        try:
            payload = build_payload(
                callback_head,
                body,
                timestamp(),
                identity_payload(),
                metadata=metadata,
            )
        except TypeError:
            payload = build_payload(callback_head, body, timestamp(), identity_payload())
        if payload is None:
            return "empty_payload"
        if enqueue_payload is not None:
            return enqueue_payload(payload, loop, ws, can_send)
        send_coro = send_payload(ws, payload)
        run_coroutine_threadsafe(send_coro, loop)
        return "scheduled"
    except Exception:
        close = getattr(send_coro, "close", None)
        if close is not None:
            try:
                close()
            except Exception:
                pass
        return "error"


def send_cloud_call_event_runtime(
    caller,
    message,
    *,
    blocked=False,
    block_reason="",
    authorized,
    get_loop,
    get_ws,
    is_connected,
    runtime_imei,
    build_payload,
    send_payload,
    timestamp,
    identity_payload,
    run_coroutine_threadsafe,
    enabled=True,
    enqueue_payload=None,
):
    caller_text = str(caller or "").strip()
    message_text = str(message or "").strip()
    if not caller_text and not message_text:
        return "empty"
    if not enabled:
        return "disabled"
    if not authorized and enqueue_payload is None:
        return "unauthorized"
    send_coro = None
    try:
        loop = get_loop()
        ws = get_ws()
        can_send = bool(
            authorized
            and loop is not None
            and loop.is_running()
            and ws is not None
            and is_connected()
        )
        if not can_send and enqueue_payload is None:
            return "not_connected"
        if not runtime_imei():
            return "missing_imei"
        try:
            payload = build_payload(
                caller_text,
                message_text,
                timestamp(),
                identity_payload(),
                blocked=blocked,
                block_reason=block_reason,
            )
        except TypeError:
            # Keep compatibility with injected/legacy payload builders that
            # only accept the common four positional arguments.
            payload = build_payload(
                caller_text,
                message_text,
                timestamp(),
                identity_payload(),
            )
        if payload is None:
            return "empty_payload"
        if enqueue_payload is not None:
            return enqueue_payload(payload, loop, ws, can_send)
        send_coro = send_payload(ws, payload)
        run_coroutine_threadsafe(send_coro, loop)
        return "scheduled"
    except Exception:
        close = getattr(send_coro, "close", None)
        if close is not None:
            try:
                close()
            except Exception:
                pass
        return "error"


def cloud_status_payload_runtime(
    *,
    serial_lock,
    get_serial,
    build_payload,
    timestamp,
    identity_payload,
    cloud_connected,
    serial_port,
    serial_baud,
    serial_mode,
):
    serial_connected = False
    try:
        with serial_lock:
            serial_obj = get_serial()
            serial_connected = bool(serial_obj is not None and serial_obj.is_open)
    except Exception:
        serial_connected = False

    return build_payload(
        timestamp(),
        identity_payload(),
        cloud_connected,
        serial_connected,
        serial_port,
        serial_baud,
        serial_mode,
    )


async def handle_cloud_message_runtime(
    message,
    *,
    is_authorized,
    set_authorized,
    reply,
    log,
    set_auth_status_from_ack,
    set_cloud_status,
    check_replay_window,
    auth_matches,
    time_text,
    timestamp,
    status_payload,
    send_serial_command,
    show_window,
    hide_window,
    handle_call_recording_message=None,
):
    incoming, error_payload = parse_cloud_message(message)
    log(f"收到：{cloud_incoming_preview(incoming)}")
    if error_payload is not None:
        await reply(error_payload)
        return

    data = incoming.data
    msg_type = incoming.msg_type

    async def task_reply(payload):
        await reply(attach_cloud_task_ids(payload, incoming.task_id))

    if is_cloud_auth_ack_type(msg_type):
        auth_status = auth_status_from_ack(data)
        if auth_status == "authorized":
            set_authorized(True)
            set_auth_status_from_ack(data)
            log(str(data.get("message") or "服务端已确认设备密码"), show_main=True)
            return

        set_authorized(False)
        set_auth_status_from_ack(data)
        if auth_status == "waiting":
            set_cloud_status("🌐 等待授权", "#b26a00")
            log(
                str(data.get("message") or "设备正在等待网页端绑定"),
                show_main=True,
            )
            return "waiting"

        set_cloud_status("🌐 授权失败", "#cc0000")
        log(str(data.get("message") or "服务端未授权设备登录，请先在网页端添加正确 IMEI 和控制密码"), show_main=True)
        return "auth_failed"

    if (
        handle_call_recording_message is not None
        and handle_call_recording_message(data)
    ):
        return "call_recording"

    if not is_authorized():
        log("已拒绝云端指令：设备尚未获得服务端授权")
        await task_reply(cloud_unauthorized_payload())
        return

    if not await check_replay_window(data, mark_seen=False):
        return

    if not auth_matches(data):
        await task_reply(cloud_auth_failed_payload())
        set_cloud_status("🌐 授权失败", "#cc0000")
        return

    if is_authorized():
        set_cloud_status("🌐 已授权", "#008000")

    if not await check_replay_window(data, mark_seen=True):
        return

    payload = await dispatch_cloud_action(
        incoming.action,
        data,
        time_text=time_text,
        timestamp=timestamp,
        status_payload=status_payload,
        send_serial_command=send_serial_command,
        show_window=show_window,
        hide_window=hide_window,
        log=log,
        command_started=lambda: task_reply(cloud_command_started_payload()),
    )
    await task_reply(payload)
