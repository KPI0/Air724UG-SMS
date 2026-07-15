import re
import urllib.parse


CLOUD_WS_DEFAULT_PATH = "/ws/device"
SMS_CALLBACK_HEAD_REGEX = re.compile(
    r"^\s*(\+?\d+)\s+\d{2}/\d{2}/\d{2},\d{2}:\d{2}:\d{2}\+\d+\s*(.*)$",
    re.DOTALL,
)


def normalize_cloud_ws_url(url: str) -> str:
    """Allow ws://host:port and auto-append the default WebSocket path."""
    text = str(url or "").strip()
    if not text:
        return ""
    lower = text.lower()
    if not (lower.startswith("ws://") or lower.startswith("wss://")):
        return text
    try:
        parsed = urllib.parse.urlsplit(text)
        if parsed.scheme.lower() not in ("ws", "wss") or not parsed.netloc:
            return text
        if parsed.path in ("", "/"):
            parsed = parsed._replace(path=CLOUD_WS_DEFAULT_PATH)
            return urllib.parse.urlunsplit(parsed)
    except Exception:
        return text
    return text


def normalize_imei(value: str) -> str:
    return re.sub(r"\D", "", str(value or "").strip())


def parse_sms_callback_head(text: str):
    match = SMS_CALLBACK_HEAD_REGEX.search(str(text or "").strip())
    if not match:
        return "", str(text or "").strip()
    return match.group(1), (match.group(2) or "").strip()


def auth_status_from_ack(data: dict):
    data = data or {}
    auth_status = str(data.get("auth_status") or "").strip().lower()
    status = str(data.get("status") or "").strip().lower()
    label = str(data.get("auth_label") or "").strip()
    message = str((data or {}).get("message") or "")
    if auth_status in ("authorized", "ok") or status in ("authorized", "auth_ok") or label == "已授权":
        return "authorized"
    if (
        auth_status in ("failed", "auth_failed", "unauthorized")
        or status in ("failed", "auth_failed", "unauthorized")
        or "不一致" in message
        or "错误" in message
    ):
        return "failed"
    return "waiting"


def version_tuple(v: str):
    text = (v or "").strip().lstrip("vV")
    parts = []
    for item in text.split("."):
        try:
            parts.append(int(item))
        except Exception:
            parts.append(0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])
