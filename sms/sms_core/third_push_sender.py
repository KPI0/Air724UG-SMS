import base64
import hashlib
import hmac
import json
import re
import smtplib
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request

from sms_core.http_deadline import urlopen_with_deadline
from email.message import EmailMessage

from sms_core.third_push_format import apply_vars, template_vars


ALLOWED_PUSH_SCHEMES = ("http", "https")
MAX_PUSH_RESPONSE_BYTES = 256 * 1024
PUSH_RESPONSE_TOO_LARGE = "渠道响应超过 256 KiB，无法确认推送结果"
WXPUSHER_API_URL = "https://wxpusher.zjiecode.com/api/send/message"
JSON_RESULT_CHANNELS = frozenset((
    "dingtalk", "wecom", "feishu", "pushdeer", "serverchan",
    "telegram", "bark", "pushover", "gotify", "wxpusher",
))
WXPUSHER_UID_RE = re.compile(r"^UID_[A-Za-z0-9_-]{1,120}$")
EMAIL_ADDRESS_RE = re.compile(r"^[^@\s,;<>]+@[^@\s,;<>]+$")
SENSITIVE_TEXT_RE = re.compile(
    r"(?i)"
    r"([\"']?(?:token|secret|password|passwd|pwd|key|sign|access_token|chat_id)[\"']?"
    r"\s*(?:=|:)\s*)"
    r"([\"']?)"
    r"([^&\s,;\"'}]+)"
    r"([\"']?)"
)


def redact_sensitive_text(text):
    return SENSITIVE_TEXT_RE.sub(r"\1\2***\4", str(text or ""))


def http_request(url, method="POST", headers=None, data=None, timeout=15, user_agent="Air724UG-SMS"):
    headers = dict(headers or {})
    headers.setdefault("User-Agent", user_agent)
    if isinstance(data, str):
        data = data.encode("utf-8")
    # Only allow http/https. urllib.urlopen otherwise honours file://, ftp://
    # and other schemes, so a webhook URL set to e.g. file:///etc/passwd would
    # read a local file and leak its contents back in the response body.
    try:
        scheme = urllib.parse.urlparse(str(url)).scheme.lower()
    except Exception:
        scheme = ""
    if scheme not in ALLOWED_PUSH_SCHEMES:
        return False, None, f"不支持的 URL 协议：{scheme or '(空)'}（仅允许 http/https）"
    try:
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        with urlopen_with_deadline(req, timeout=timeout) as resp:
            raw_body = resp.read(MAX_PUSH_RESPONSE_BYTES + 1)
            if len(raw_body) > MAX_PUSH_RESPONSE_BYTES:
                return False, resp.getcode(), PUSH_RESPONSE_TOO_LARGE
            body = raw_body.decode("utf-8", "replace")
            return True, resp.getcode(), body
    except urllib.error.HTTPError as exc:
        try:
            exc.close()
        except Exception:
            pass
        return False, exc.code, ""
    except TimeoutError:
        return False, None, "推送请求超时，无法确认是否送达"
    except urllib.error.URLError as exc:
        if isinstance(exc.reason, TimeoutError):
            return False, None, "推送请求超时，无法确认是否送达"
        return False, None, redact_sensitive_text(str(exc))
    except Exception as exc:
        return False, None, redact_sensitive_text(str(exc))


def api_ok(channel: str, http_ok: bool, code, body: str, *, expected_uids=None):
    try:
        status_code = int(code) if code is not None else None
    except (TypeError, ValueError):
        status_code = None
    status_text = str(status_code) if status_code is not None else "-"
    if not http_ok or status_code is None or not (200 <= status_code < 300):
        if not http_ok and body == PUSH_RESPONSE_TOO_LARGE:
            return False, PUSH_RESPONSE_TOO_LARGE
        return False, f"HTTP {status_text} 请求失败"

    if channel not in JSON_RESULT_CHANNELS:
        return True, f"HTTP {status_text}"
    invalid_result = (False, f"HTTP {status_text} API 响应格式无效，无法确认推送结果")
    try:
        data = json.loads((body or "").strip())
    except (ValueError, TypeError, RecursionError):
        return invalid_result
    if not isinstance(data, dict):
        return invalid_result

    if channel == "wxpusher":
        if str(data.get("code", "")) != "1000":
            return False, f"HTTP {status_text} 渠道返回业务错误"
        results = data.get("data")
        if not isinstance(results, list) or not results or any(
            not isinstance(item, dict) or "code" not in item for item in results
        ):
            return False, f"HTTP {status_text} 接收方结果缺失或格式无效，无法确认推送结果"
        if expected_uids is not None:
            returned_uids = [item.get("uid") for item in results]
            if len(returned_uids) != len(expected_uids) or any(
                returned_uids.count(uid) != 1 for uid in expected_uids
            ):
                return False, f"HTTP {status_text} 接收方结果不完整或不匹配，无法确认推送结果"
        failed = sum(str(item["code"]) != "1000" for item in results)
        if failed:
            return False, f"HTTP {status_text} 接收方任务受理 {len(results) - failed}，失败 {failed}"

    if channel in ("dingtalk", "wecom"):
        fields, expected = ("errcode",), "0"
    elif channel == "feishu":
        fields, expected = ("code", "StatusCode"), "0"
    elif channel in ("pushdeer", "serverchan"):
        fields, expected = ("code",), "0"
    elif channel == "bark":
        fields, expected = ("code",), "200"
    elif channel == "pushover":
        fields, expected = ("status",), "1"
    elif channel == "telegram":
        if type(data.get("ok")) is not bool:
            return invalid_result
        if not data["ok"]:
            return False, f"HTTP {status_text} 渠道返回业务错误"
        fields, expected = (), ""
    elif channel == "gotify":
        if type(data.get("id")) is not int or data["id"] <= 0:
            return invalid_result
        fields, expected = (), ""
    else:
        fields, expected = (), ""  # WxPusher was validated per recipient above.

    if fields:
        present = [data[field] for field in fields if field in data]
        if not present or any(type(value) not in (int, str) for value in present):
            return invalid_result
        if any(str(value) != expected for value in present):
            return False, f"HTTP {status_text} 渠道返回业务错误"

    return True, f"HTTP {status_text}"


def required(settings, key, label):
    value = str(settings.get(key, "")).strip()
    if not value:
        return None, f"未配置 {label}"
    return value, None


def parse_wxpusher_uids(value, maximum=100):
    items = []
    seen = set()
    for raw_item in re.split(r"[\s,，;；]+", str(value or "").strip()):
        item = str(raw_item or "").strip()
        if not item or item in seen:
            continue
        if not WXPUSHER_UID_RE.fullmatch(item):
            raise ValueError("WXPusher UID 必须以 UID_ 开头且只能包含字母、数字、下划线或短横线")
        items.append(item)
        seen.add(item)
        if len(items) > maximum:
            raise ValueError(f"WXPusher UID 最多支持 {maximum} 个")
    if not items:
        raise ValueError("WXPusher UID 不能为空")
    return items


def parse_email_addresses(value, label="邮箱", maximum=100):
    items = []
    seen = set()
    for raw_item in re.split(r"[,，;；\r\n]+", str(value or "").strip()):
        item = str(raw_item or "").strip()
        if not item:
            continue
        if len(item) > 254 or not EMAIL_ADDRESS_RE.fullmatch(item):
            raise ValueError(f"{label}格式无效")
        key = item.casefold()
        if key not in seen:
            items.append(item)
            seen.add(key)
        if len(items) > maximum:
            raise ValueError(f"{label}最多支持 {maximum} 个")
    if not items:
        raise ValueError(f"{label}不能为空")
    return items


def parse_smtp_port(value):
    try:
        port = int(str(value or "").strip())
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError("SMTP 端口必须是 1-65535 的整数") from exc
    if port < 1 or port > 65535:
        raise ValueError("SMTP 端口必须是 1-65535 的整数")
    return port


def _smtp_host(settings):
    host = str(settings.get("email_smtp_host") or "").strip().strip("[]")
    if (
        not host
        or any(char.isspace() for char in host)
        or any(char in host for char in "/?#@")
        or "://" in host
    ):
        raise ValueError("SMTP 服务器格式无效，请只填写域名或 IP 地址")
    return host


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
    secret = str(settings.get("dingtalk_secret", "")).strip()
    keyword = str(settings.get("dingtalk_keyword", "")).strip()
    if keyword and not secret and keyword not in message:
        message = f"{keyword}\n{message}"
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


def send_wxpusher(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    app_token, err = required(settings, "wxpusher_app_token", "WXPUSHER_APP_TOKEN")
    if err:
        return False, err
    try:
        uids = parse_wxpusher_uids(settings.get("wxpusher_uids"))
    except ValueError as exc:
        return False, str(exc)
    data = json.dumps(
        {
            "appToken": app_token,
            "content": message,
            "summary": "Air724UG 通知",
            "contentType": 1,
            "uids": uids,
        },
        ensure_ascii=False,
    )
    http_ok, code, body = http_request(
        WXPUSHER_API_URL,
        "POST",
        {"Content-Type": "application/json; charset=utf-8"},
        data,
        user_agent=user_agent,
    )
    return api_ok("wxpusher", http_ok, code, body, expected_uids=uids)


def send_email(message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    try:
        host = _smtp_host(settings)
        smtp_port = parse_smtp_port(settings.get("email_smtp_port"))
        encryption = str(settings.get("email_encryption") or "").strip().lower()
        if encryption not in ("ssl", "starttls"):
            return False, "EMAIL_ENCRYPTION 只能是 ssl 或 starttls"
        username, err = required(settings, "email_username", "EMAIL_USERNAME")
        if err:
            return False, err
        password, err = required(settings, "email_password", "EMAIL_PASSWORD")
        if err:
            return False, err
        from_address = parse_email_addresses(
            settings.get("email_from_address"), "发件邮箱", maximum=1
        )[0]
        to_addresses = parse_email_addresses(settings.get("email_to_addresses"), "收件邮箱")
        subject = str(settings.get("email_subject") or "Air724UG 通知").strip()[:200]
        if "\r" in subject or "\n" in subject:
            return False, "邮件主题格式无效"
        email = EmailMessage()
        email["Subject"] = subject or "Air724UG 通知"
        email["From"] = from_address
        email["To"] = ", ".join(to_addresses)
        email.set_content(str(message or ""))

        context = ssl.create_default_context()
        client = None
        try:
            if encryption == "ssl":
                client = smtplib.SMTP_SSL(host, smtp_port, timeout=15, context=context)
            else:
                client = smtplib.SMTP(host, smtp_port, timeout=15)
                client.ehlo()
                client.starttls(context=context)
                client.ehlo()
            client.login(username, password)
            refused = client.send_message(email)
        finally:
            if client is not None:
                try:
                    client.quit()
                except Exception:
                    try:
                        client.close()
                    except Exception:
                        pass
        if refused:
            return False, f"SMTP 接收方已接受 {len(to_addresses) - len(refused)}，拒收 {len(refused)}"
        return True, "SMTP 邮件已发送"
    except ValueError as exc:
        return False, str(exc)
    except Exception as exc:
        if isinstance(exc, smtplib.SMTPRecipientsRefused):
            return False, f"SMTP 接收方已接受 0，拒收 {len(exc.recipients)}"
        if isinstance(exc, (smtplib.SMTPAuthenticationError, smtplib.SMTPNotSupportedError)):
            return False, "SMTP 认证或加密协商失败"
        if isinstance(exc, smtplib.SMTPResponseException):
            return False, "SMTP 服务器拒绝发送邮件"
        if isinstance(exc, (TimeoutError, smtplib.SMTPConnectError, OSError)):
            return False, "SMTP 连接失败"
        return False, "SMTP 邮件发送失败"


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
    "wxpusher": send_wxpusher,
    "email": send_email,
}


def send_channel(channel: str, message: str, settings: dict, user_agent="Air724UG-SMS", port: str = ""):
    handler = CHANNEL_HANDLERS.get(channel)
    if handler is None:
        return False, "未知通知通道"
    return handler(message, settings, user_agent=user_agent, port=port)
