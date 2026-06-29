import json

from sms_core.cloud_messages import (
    attach_cloud_task_ids,
    cloud_auth_failed_payload,
    cloud_unauthorized_payload,
    dispatch_cloud_action,
    is_cloud_auth_ack_type,
    parse_cloud_message,
)
from sms_core.cloud_protocol import auth_status_from_ack
from sms_core.cloud_security import HIDDEN_SMS_COMMAND, HIDDEN_SMS_META, safe_preview


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
):
    cmd = str(command or "").strip()
    if not cmd:
        return False, "AT 指令不能为空"

    try:
        display_cmd = _cloud_command_display_text(cmd, command_meta)
        with serial_lock:
            serial_obj = get_serial()
        result = write_command_result(serial_obj, cmd)
        if not result.ok:
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

        log(f"已向串口发送：{display_cmd}")
        return True, f"已发送：{display_cmd}"
    except Exception as exc:
        return False, f"发送失败：{exc}"


def _cloud_command_display_text(cmd, command_meta):
    meta = command_meta if isinstance(command_meta, dict) else {}
    if str(meta.get("sms_log") or "").strip().lower() == "suppress":
        return HIDDEN_SMS_COMMAND
    return str(cmd or "").strip()


def cloud_incoming_preview(incoming):
    data = getattr(incoming, "data", None)
    if not isinstance(data, dict):
        return safe_preview(getattr(incoming, "raw", ""))

    masked = dict(data)
    sms_log = str(masked.get("sms_log") or "").strip().lower()
    if sms_log == "suppress":
        for key in ("cmd", "command", "data"):
            if key in masked:
                masked[key] = HIDDEN_SMS_COMMAND

    if sms_log in ("summary", "suppress") or str(masked.get("command_kind") or "") == "send_sms":
        for key in ("sms_phone", "sms_message"):
            if key in masked:
                masked[key] = HIDDEN_SMS_META

    return safe_preview(json.dumps(masked, ensure_ascii=False))


def _log_cloud_command_to_port(port_ui, cmd, command_meta):
    meta = command_meta if isinstance(command_meta, dict) else {}
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
):
    body = str(full_msg or "").strip()
    if not body:
        return "empty"
    if not authorized:
        return "unauthorized"
    try:
        loop = get_loop()
        ws = get_ws()
        if loop is None or not loop.is_running() or ws is None or not is_connected():
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
        run_coroutine_threadsafe(send_payload(ws, payload), loop)
        return "scheduled"
    except Exception:
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
        log(str(data.get("message") or "服务端未授权设备登录，请先在网页端添加正确 IMEI 和控制密码"), show_main=True)
        return

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
    )
    await task_reply(payload)
