from dataclasses import dataclass
import json

from sms_core.cloud_auth import command_action, command_text


JSON_OBJECT_REQUIRED_MESSAGE = "仅支持 JSON 对象消息，且必须携带 target_imei 和 secret/password"
CLOUD_UNAUTHORIZED_MESSAGE = "设备尚未获得服务端授权，已拒绝执行云端指令"
CLOUD_AUTH_FAILED_MESSAGE = "IMEI 或密码校验失败"


@dataclass(frozen=True)
class CloudIncomingMessage:
    raw: str
    data: dict | None
    msg_type: str = ""
    task_id: str = ""
    action: str = ""


def decode_cloud_message(message) -> str:
    if isinstance(message, bytes):
        return message.decode("utf-8", "ignore")
    return str(message)


def parse_cloud_message(message):
    raw = decode_cloud_message(message)
    try:
        data = json.loads(raw)
    except Exception:
        return CloudIncomingMessage(raw, None), cloud_protocol_error_payload()

    if not isinstance(data, dict):
        return CloudIncomingMessage(raw, None), cloud_protocol_error_payload()

    task_id = str(data.get("task_id") or data.get("command_task_id") or "").strip()
    return (
        CloudIncomingMessage(
            raw=raw,
            data=data,
            msg_type=str(data.get("type") or "").strip().lower(),
            task_id=task_id,
            action=command_action(data),
        ),
        None,
    )


def cloud_protocol_error_payload():
    return {
        "type": "error",
        "ok": False,
        "message": JSON_OBJECT_REQUIRED_MESSAGE,
    }


def attach_cloud_task_ids(payload, task_id):
    if not task_id or not isinstance(payload, dict):
        return payload
    return {
        **payload,
        "task_id": payload.get("task_id") or task_id,
        "command_task_id": payload.get("command_task_id") or task_id,
    }


def is_cloud_auth_ack_type(msg_type):
    return str(msg_type or "").strip().lower() in (
        "device_login_ack",
        "device_auth",
        "device_auth_result",
    )


def cloud_action_kind(action):
    action = str(action or "").strip().lower()
    if action in ("ping", "heartbeat"):
        return "ping"
    if action in ("status", "get_status"):
        return "status"
    if action in ("send_at", "at", "cmd", "command"):
        return "send_at"
    if action == "show_window":
        return "show_window"
    if action == "hide_window":
        return "hide_window"
    return "unknown"


def cloud_unauthorized_payload():
    return {
        "type": "auth_failed",
        "ok": False,
        "message": CLOUD_UNAUTHORIZED_MESSAGE,
    }


def cloud_auth_failed_payload():
    return {
        "type": "auth_failed",
        "ok": False,
        "message": CLOUD_AUTH_FAILED_MESSAGE,
    }


def cloud_pong_payload(time_text, timestamp):
    return {
        "type": "pong",
        "ok": True,
        "time": time_text,
        "timestamp": timestamp,
    }


def cloud_send_at_result_payload(ok, message):
    return {
        "type": "send_at_result",
        "ok": bool(ok),
        "message": message,
    }


def cloud_window_result_payload(action):
    kind = cloud_action_kind(action)
    if kind == "show_window":
        return {"type": "show_window_result", "ok": True}
    if kind == "hide_window":
        return {"type": "hide_window_result", "ok": True}
    return cloud_unknown_command_payload(action)


def cloud_unknown_command_payload(action):
    return {
        "type": "error",
        "ok": False,
        "message": f"未知云端指令：{action or '(empty)'}",
    }


async def dispatch_cloud_action(
    action,
    data,
    *,
    time_text,
    timestamp,
    status_payload,
    send_serial_command,
    show_window,
    hide_window,
    log,
):
    action_kind = cloud_action_kind(action)

    if action_kind == "ping":
        return cloud_pong_payload(time_text(), timestamp())

    if action_kind == "status":
        return status_payload()

    if action_kind == "send_at":
        command = command_text(data)
        log(f"云端下发指令：{command}")
        ok, info = await send_serial_command(command, data)
        return cloud_send_at_result_payload(ok, info)

    if action_kind == "show_window":
        show_window()
        return cloud_window_result_payload(action)

    if action_kind == "hide_window":
        hide_window()
        return cloud_window_result_payload(action)

    return cloud_unknown_command_payload(action)
