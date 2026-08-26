import re
import urllib.parse


CLOUD_WS_DEFAULT_PATH = "/ws/device"
SMS_CALLBACK_HEAD_REGEX = re.compile(
    r"^\s*(?P<sender>[^\r\n]+?)\s+"
    r"(?P<timestamp>\d{2}/\d{2}/\d{2},\d{2}:\d{2}:\d{2}\+\d+)\s*"
    r"(?P<body>.*)$",
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
        if "#" in text:
            return text
        if parsed.path in ("", "/"):
            parsed = parsed._replace(path=CLOUD_WS_DEFAULT_PATH)
            return urllib.parse.urlunsplit(parsed)
    except Exception:
        return text
    return text


def cloud_ws_url_has_host(url: str) -> bool:
    """Return whether a WebSocket URL is syntactically safe to connect."""
    try:
        text = str(url or "").strip()
        parsed = urllib.parse.urlsplit(text)
        host = parsed.hostname
    except (TypeError, ValueError):
        return False
    if any(char.isspace() for char in text):
        return False
    authority = parsed.netloc.rsplit("@", 1)[-1]
    port_text = None
    if authority.startswith("["):
        closing = authority.find("]")
        if closing <= 1:
            return False
        suffix = authority[closing + 1 :]
        if suffix:
            if not suffix.startswith(":"):
                return False
            port_text = suffix[1:]
    else:
        if authority.count(":") > 1:
            return False
        if ":" in authority:
            _authority_host, port_text = authority.rsplit(":", 1)
            if not _authority_host:
                return False
    if port_text is not None:
        if not port_text.isdigit():
            return False
        try:
            if not 1 <= int(port_text) <= 65535:
                return False
        except (TypeError, ValueError, OverflowError):
            return False
    return (
        parsed.scheme.lower() in ("ws", "wss")
        and bool(host)
        and bool(parsed.netloc)
        and not parsed.fragment
        and "#" not in text
        and not any(ord(char) < 32 or ord(char) == 127 for char in text)
        and not any(char.isspace() for char in host)
        and not any(char in "/\\?#" for char in host)
    )


def normalize_imei(value: str) -> str:
    return re.sub(r"\D", "", str(value or "").strip())


def parse_sms_callback_head(text: str):
    match = SMS_CALLBACK_HEAD_REGEX.search(str(text or "").strip())
    if not match:
        return "", str(text or "").strip()
    return (match.group("sender") or "").strip(), (match.group("body") or "").strip()


def auth_status_from_ack(data: dict):
    data = data or {}
    # When the server includes an explicit result flag, it is authoritative.
    # Do not let contradictory status/auth_status fields turn an explicit
    # rejection (or a non-boolean value) into an authorized session.
    if "ok" in data and data.get("ok") is not True:
        return "failed"
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
