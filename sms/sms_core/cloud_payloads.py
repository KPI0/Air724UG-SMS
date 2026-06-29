import socket
from datetime import datetime

from sms_core.cloud_protocol import parse_sms_callback_head


def current_time_text(now=None):
    current = now or datetime.now()
    return current.strftime("%Y-%m-%d %H:%M:%S")


def identity_payload(imei, app_version, device_name=None):
    normalized = str(imei or "").strip()
    return {
        "imei": normalized,
        "device_imei": normalized,
        "device_name": device_name or socket.gethostname(),
        "app_version": app_version,
    }


def build_register_payload(
    auto_upload,
    timestamp,
    identity,
    secret,
    serial_port,
    serial_baud,
    serial_mode,
    time_text=None,
):
    return {
        "type": "device_login",
        "event": "register" if auto_upload else "hidden",
        "public": bool(auto_upload),
        "auto_upload": bool(auto_upload),
        "hidden": not bool(auto_upload),
        "time": time_text or current_time_text(),
        "timestamp": timestamp,
        **identity,
        "secret": secret,
        "serial_port": serial_port,
        "serial_baud": serial_baud,
        "serial_mode": serial_mode,
    }


def build_unregister_payload(
    reason,
    timestamp,
    identity,
    secret,
    serial_port,
    serial_baud,
    serial_mode,
    time_text=None,
):
    return {
        "type": "device_login",
        "event": "offline",
        "action": "offline",
        "status": "offline",
        "online": False,
        "time": time_text or current_time_text(),
        "timestamp": timestamp,
        **identity,
        "secret": secret,
        "serial_port": serial_port,
        "serial_baud": serial_baud,
        "serial_mode": serial_mode,
        "reason": reason,
    }


def truncate_serial_log_text(line: str, limit: int = 2000):
    text = str(line or "").strip()
    if len(text) > limit:
        return text[:limit] + "..."
    return text


def build_serial_log_payload(
    line,
    timestamp,
    identity,
    serial_port,
    serial_baud,
    time_text=None,
):
    text = truncate_serial_log_text(line)
    if not text:
        return None
    return {
        "type": "log",
        "tag": "debug",
        "time": time_text or current_time_text(),
        "timestamp": timestamp,
        **identity,
        "serial_port": serial_port,
        "serial_baud": serial_baud,
        "data": f"[串口] {text}",
        "raw": text,
    }


def build_sms_event_payload(
    callback_head,
    full_msg,
    timestamp,
    identity,
    time_text=None,
    metadata=None,
):
    body = str(full_msg or "").strip()
    if not body:
        return None
    sender, first_body = parse_sms_callback_head(callback_head)
    payload = {
        "type": "sms_event",
        "event_type": "sms",
        "tag": "sms",
        "time": time_text or current_time_text(),
        "timestamp": timestamp,
        **identity,
        "from": sender,
        "phone": sender,
        "content": body,
        "body": body,
        "message": f"收到短信：来自 {sender or '未知号码'}，内容：{body}",
        "raw": (callback_head + "\n" + body).strip() if callback_head and first_body != body else body,
    }
    meta = metadata if isinstance(metadata, dict) else {}
    message_trace_id = str(meta.get("message_trace_id") or "").strip()
    if message_trace_id:
        payload["message_trace_id"] = message_trace_id
        payload["trace_id"] = message_trace_id
    return payload


def build_status_payload(
    timestamp,
    identity,
    cloud_connected,
    serial_connected,
    serial_port,
    serial_baud,
    serial_mode,
    time_text=None,
):
    return {
        "type": "status",
        "ok": True,
        "time": time_text or current_time_text(),
        "timestamp": timestamp,
        **identity,
        "cloud_connected": bool(cloud_connected),
        "serial_connected": bool(serial_connected),
        "serial_port": serial_port,
        "serial_baud": serial_baud,
        "serial_mode": serial_mode,
    }
