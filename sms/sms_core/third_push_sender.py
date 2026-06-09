import base64
import hashlib
import hmac
import json
import time
import urllib.error
import urllib.parse
import urllib.request

from sms_core.third_push_format import apply_vars, template_vars


def http_request(url, method="POST", headers=None, data=None, timeout=15, user_agent="Air724UG-SMS"):
    headers = dict(headers or {})
    headers.setdefault("User-Agent", user_agent)
    if isinstance(data, str):
        data = data.encode("utf-8")
    try:
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = resp.read(4096).decode("utf-8", "replace")
            return True, resp.getcode(), body
    except urllib.error.HTTPError as exc:
        try:
            body = exc.read(4096).decode("utf-8", "replace")
        except Exception:
            body = ""
        return False, exc.code, body
    except Exception as exc:
        return False, None, str(exc)


def api_ok(channel: str, http_ok: bool, code, body: str):
    if not http_ok or code is None or not (200 <= int(code) < 300):
        return False, f"HTTP {code or '-'} {body}".strip()

    text = (body or "").strip()
    if not text:
        return True, f"HTTP {code}"

    try:
        data = json.loads(text)
    except Exception:
        return True, f"HTTP {code}"

    if channel in ("dingtalk", "wecom"):
        errcode = data.get("errcode", 0)
        if str(errcode) not in ("0", ""):
            return False, data.get("errmsg") or text
    elif channel == "feishu":
        errcode = data.get("code", data.get("StatusCode", 0))
        if str(errcode) not in ("0", ""):
            return False, data.get("msg") or data.get("StatusMessage") or text
    elif channel in ("pushdeer", "serverchan"):
        errcode = data.get("code", 0)
        if str(errcode) not in ("0", ""):
            return False, data.get("message") or data.get("msg") or text

    return True, f"HTTP {code}"


def required(settings, key, label):
    value = str(settings.get(key, "")).strip()
    if not value:
        return None, f"未配置 {label}"
    return value, None


def send_custom_post(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    url, err = required(settings, "custom_post_url", "CUSTOM_POST_URL")
    if err:
        return False, err
    content_type = str(settings.get("custom_post_content_type") or "application/json").strip()
    body_raw = str(settings.get("custom_post_body") or "").strip()
    variables = template_vars(message, port)

    headers = {"Content-Type": content_type or "application/json"}
    try:
        body_obj = json.loads(body_raw) if body_raw else {}
        body_obj = apply_vars(body_obj, variables)
        if "json" in content_type.lower():
            data = json.dumps(body_obj, ensure_ascii=False)
        elif isinstance(body_obj, dict):
            data = urllib.parse.urlencode(body_obj)
        else:
            data = urllib.parse.urlencode({"msg": message})
    except Exception:
        data = apply_vars(body_raw or "{msg}", variables)

    http_ok, code, body = http_request(url, "POST", headers, data, user_agent=user_agent)
    return api_ok("custom_post", http_ok, code, body)


def send_telegram(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    url, err = required(settings, "telegram_api", "TELEGRAM_API")
    if err:
        return False, err
    chat_id, err = required(settings, "telegram_chat_id", "TELEGRAM_CHAT_ID")
    if err:
        return False, err
    data = json.dumps({
        "chat_id": chat_id,
        "disable_web_page_preview": True,
        "text": message,
    }, ensure_ascii=False)
    http_ok, code, body = http_request(url, "POST", {"Content-Type": "application/json"}, data, user_agent=user_agent)
    return api_ok("telegram", http_ok, code, body)


def send_pushdeer(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    url, err = required(settings, "pushdeer_api", "PUSHDEER_API")
    if err:
        return False, err
    push_key, err = required(settings, "pushdeer_key", "PUSHDEER_KEY")
    if err:
        return False, err
    data = urllib.parse.urlencode({"pushkey": push_key, "type": "text", "text": message})
    http_ok, code, body = http_request(
        url,
        "POST",
        {"Content-Type": "application/x-www-form-urlencoded"},
        data,
        user_agent=user_agent,
    )
    return api_ok("pushdeer", http_ok, code, body)


def send_bark(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    api, err = required(settings, "bark_api", "BARK_API")
    if err:
        return False, err
    key, err = required(settings, "bark_key", "BARK_KEY")
    if err:
        return False, err
    url = api.rstrip("/") + "/" + key.strip("/")
    data = urllib.parse.urlencode({"body": message})
    http_ok, code, body = http_request(
        url,
        "POST",
        {"Content-Type": "application/x-www-form-urlencoded"},
        data,
        user_agent=user_agent,
    )
    return api_ok("bark", http_ok, code, body)


def send_dingtalk(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    url, err = required(settings, "dingtalk_webhook", "DINGTALK_WEBHOOK")
    if err:
        return False, err
    keyword = str(settings.get("dingtalk_keyword", "")).strip()
    if keyword and keyword not in message:
        message = f"{keyword}\n{message}"
    secret = str(settings.get("dingtalk_secret", "")).strip()
    if secret:
        timestamp = str(int(time.time() * 1000))
        sign_raw = f"{timestamp}\n{secret}".encode("utf-8")
        digest = hmac.new(secret.encode("utf-8"), sign_raw, hashlib.sha256).digest()
        sign = urllib.parse.quote_plus(base64.b64encode(digest).decode("utf-8"))
        sep = "&" if "?" in url else "?"
        url = f"{url}{sep}timestamp={timestamp}&sign={sign}"
    data = json.dumps({"msgtype": "text", "text": {"content": message}}, ensure_ascii=False)
    http_ok, code, body = http_request(
        url,
        "POST",
        {"Content-Type": "application/json; charset=utf-8"},
        data,
        user_agent=user_agent,
    )
    return api_ok("dingtalk", http_ok, code, body)


def send_feishu(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    url, err = required(settings, "feishu_webhook", "FEISHU_WEBHOOK")
    if err:
        return False, err
    data = json.dumps({"msg_type": "text", "content": {"text": message}}, ensure_ascii=False)
    http_ok, code, body = http_request(
        url,
        "POST",
        {"Content-Type": "application/json; charset=utf-8"},
        data,
        user_agent=user_agent,
    )
    return api_ok("feishu", http_ok, code, body)


def send_wecom(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    url, err = required(settings, "wecom_webhook", "WECOM_WEBHOOK")
    if err:
        return False, err
    data = json.dumps({"msgtype": "text", "text": {"content": message}}, ensure_ascii=False)
    http_ok, code, body = http_request(
        url,
        "POST",
        {"Content-Type": "application/json; charset=utf-8"},
        data,
        user_agent=user_agent,
    )
    return api_ok("wecom", http_ok, code, body)


def send_pushover(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    token, err = required(settings, "pushover_api_token", "PUSHOVER_API_TOKEN")
    if err:
        return False, err
    user_key, err = required(settings, "pushover_user_key", "PUSHOVER_USER_KEY")
    if err:
        return False, err
    data = json.dumps({"token": token, "user": user_key, "message": message}, ensure_ascii=False)
    http_ok, code, body = http_request(
        "https://api.pushover.net/1/messages.json",
        "POST",
        {"Content-Type": "application/json; charset=utf-8"},
        data,
        user_agent=user_agent,
    )
    return api_ok("pushover", http_ok, code, body)


def send_inotify(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    api, err = required(settings, "inotify_api", "INOTIFY_API")
    if err:
        return False, err
    url = api.rstrip("/") + "/" + urllib.parse.quote(message, safe="")
    http_ok, code, body = http_request(url, "GET", user_agent=user_agent)
    return api_ok("inotify", http_ok, code, body)


def send_next_smtp_proxy(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    url, err = required(settings, "next_smtp_proxy_api", "NEXT_SMTP_PROXY_API")
    if err:
        return False, err
    required_fields = (
        ("next_smtp_proxy_user", "NEXT_SMTP_PROXY_USER"),
        ("next_smtp_proxy_password", "NEXT_SMTP_PROXY_PASSWORD"),
        ("next_smtp_proxy_host", "NEXT_SMTP_PROXY_HOST"),
        ("next_smtp_proxy_port", "NEXT_SMTP_PROXY_PORT"),
        ("next_smtp_proxy_to_email", "NEXT_SMTP_PROXY_TO_EMAIL"),
    )
    values = {}
    for key, label in required_fields:
        values[key], err = required(settings, key, label)
        if err:
            return False, err
    data = urllib.parse.urlencode({
        "user": values["next_smtp_proxy_user"],
        "password": values["next_smtp_proxy_password"],
        "host": values["next_smtp_proxy_host"],
        "port": values["next_smtp_proxy_port"],
        "form_name": settings.get("next_smtp_proxy_form_name", ""),
        "to_email": values["next_smtp_proxy_to_email"],
        "subject": settings.get("next_smtp_proxy_subject", ""),
        "text": message,
    })
    http_ok, code, body = http_request(
        url,
        "POST",
        {"Content-Type": "application/x-www-form-urlencoded"},
        data,
        user_agent=user_agent,
    )
    return api_ok("next-smtp-proxy", http_ok, code, body)


def send_gotify(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    api, err = required(settings, "gotify_api", "GOTIFY_API")
    if err:
        return False, err
    token, err = required(settings, "gotify_token", "GOTIFY_TOKEN")
    if err:
        return False, err
    try:
        priority = int(str(settings.get("gotify_priority", "8")).strip() or "8")
    except Exception:
        priority = 8
    url = api.rstrip("/") + "/message?token=" + urllib.parse.quote(token, safe="")
    data = json.dumps({
        "title": settings.get("gotify_title", "Air724UG"),
        "message": message,
        "priority": priority,
    }, ensure_ascii=False)
    http_ok, code, body = http_request(
        url,
        "POST",
        {"Content-Type": "application/json; charset=utf-8"},
        data,
        user_agent=user_agent,
    )
    return api_ok("gotify", http_ok, code, body)


def send_serverchan(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    url, err = required(settings, "serverchan_api", "SERVERCHAN_API")
    if err:
        return False, err
    title, err = required(settings, "serverchan_title", "SERVERCHAN_TITLE")
    if err:
        return False, err
    data = urllib.parse.urlencode({"title": title, "desp": message})
    http_ok, code, body = http_request(
        url,
        "POST",
        {"Content-Type": "application/x-www-form-urlencoded"},
        data,
        user_agent=user_agent,
    )
    return api_ok("serverchan", http_ok, code, body)


CHANNEL_HANDLERS = {
    "custom_post": send_custom_post,
    "telegram": send_telegram,
    "pushdeer": send_pushdeer,
    "bark": send_bark,
    "dingtalk": send_dingtalk,
    "feishu": send_feishu,
    "wecom": send_wecom,
    "pushover": send_pushover,
    "inotify": send_inotify,
    "next-smtp-proxy": send_next_smtp_proxy,
    "gotify": send_gotify,
    "serverchan": send_serverchan,
}


def send_channel(channel: str, message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    handler = CHANNEL_HANDLERS.get(channel)
    if handler is None:
        return False, "未知通知通道"
    return handler(message, settings, user_agent=user_agent, port=port)
