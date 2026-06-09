from dataclasses import dataclass
import json
import re

from .config_schema import THIRD_PUSH_DEFAULTS
from sms_core.third_push_format import apply_vars, format_message, template_vars
from sms_core.third_push_sender import api_ok, http_request, required, send_channel


THIRD_PUSH_TEST_MESSAGE = "这是一条三方推送测试短信。"
THIRD_PUSH_CHANNELS = [
    ("dingtalk", "钉钉"),
    ("wecom", "企业微信"),
    ("feishu", "飞书"),
    ("custom_post", "自定义POST"),
    ("telegram", "Telegram"),
    ("pushdeer", "PushDeer"),
    ("bark", "Bark"),
    ("pushover", "Pushover"),
    ("inotify", "Inotify"),
    ("next-smtp-proxy", "next-smtp-proxy"),
    ("gotify", "Gotify"),
    ("serverchan", "Server酱"),
]
THIRD_PUSH_CHANNEL_LABELS = dict(THIRD_PUSH_CHANNELS)
THIRD_PUSH_SETTINGS_KEYS = [
    k for k in THIRD_PUSH_DEFAULTS
    if k not in ("enabled", "sms_enabled", "call_enabled", "notify_type")
]
THIRD_PUSH_REQUIRED_FIELDS = {
    "dingtalk": (("dingtalk_webhook", "DINGTALK_WEBHOOK"),),
    "wecom": (("wecom_webhook", "WECOM_WEBHOOK"),),
    "feishu": (("feishu_webhook", "FEISHU_WEBHOOK"),),
    "custom_post": (("custom_post_url", "CUSTOM_POST_URL"),),
    "telegram": (("telegram_api", "TELEGRAM_API"), ("telegram_chat_id", "TELEGRAM_CHAT_ID")),
    "pushdeer": (("pushdeer_api", "PUSHDEER_API"), ("pushdeer_key", "PUSHDEER_KEY")),
    "bark": (("bark_api", "BARK_API"), ("bark_key", "BARK_KEY")),
    "pushover": (("pushover_api_token", "PUSHOVER_API_TOKEN"), ("pushover_user_key", "PUSHOVER_USER_KEY")),
    "inotify": (("inotify_api", "INOTIFY_API"),),
    "next-smtp-proxy": (
        ("next_smtp_proxy_api", "NEXT_SMTP_PROXY_API"),
        ("next_smtp_proxy_user", "NEXT_SMTP_PROXY_USER"),
        ("next_smtp_proxy_password", "NEXT_SMTP_PROXY_PASSWORD"),
        ("next_smtp_proxy_host", "NEXT_SMTP_PROXY_HOST"),
        ("next_smtp_proxy_port", "NEXT_SMTP_PROXY_PORT"),
        ("next_smtp_proxy_to_email", "NEXT_SMTP_PROXY_TO_EMAIL"),
    ),
    "gotify": (("gotify_api", "GOTIFY_API"), ("gotify_token", "GOTIFY_TOKEN")),
    "serverchan": (("serverchan_api", "SERVERCHAN_API"), ("serverchan_title", "SERVERCHAN_TITLE")),
}


def push_label(channel: str) -> str:
    return THIRD_PUSH_CHANNEL_LABELS.get(channel, channel)


def third_push_state(enabled, sms_enabled, call_enabled, channels, settings):
    return {
        "enabled": bool(enabled),
        "sms_enabled": bool(sms_enabled),
        "call_enabled": bool(call_enabled),
        "channels": list(channels or []),
        "settings": dict(settings or {}),
    }


def third_push_save_kwargs(values):
    enabled, sms_enabled, call_enabled, selected, settings = values
    return {
        "enabled": bool(enabled),
        "sms_enabled": bool(sms_enabled),
        "call_enabled": bool(call_enabled),
        "notify_type": list(selected or []),
        "settings": dict(settings or {}),
    }


def third_push_saved_status(enabled, selected):
    state_text = "已开启" if enabled else "已关闭"
    channels_text = ", ".join(selected or []) or "未选择"
    return f"📡 三方推送：{state_text}，通道：{channels_text}"


def parse_push_channels(raw: str):
    channels = []
    raw = (raw or "").strip()
    if raw:
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, str):
                parsed = [parsed]
        except Exception:
            parsed = [x.strip() for x in re.split(r"[|,，\s]+", raw) if x.strip()]
        if isinstance(parsed, (list, tuple)):
            for item in parsed:
                ch = str(item).strip()
                if ch in THIRD_PUSH_CHANNEL_LABELS and ch not in channels:
                    channels.append(ch)
    return channels


def coerce_text_list(value):
    if isinstance(value, str):
        value = [value]
    if not isinstance(value, (list, tuple)):
        return []
    result = []
    for item in value:
        text = str(item).strip()
        if text:
            result.append(text)
    return result


def validate_push_settings(channels, settings):
    missing = []
    for channel in channels:
        for key, label in THIRD_PUSH_REQUIRED_FIELDS.get(channel, ()):
            if not str(settings.get(key, "")).strip():
                missing.append(f"{push_label(channel)}: {label}")
    return missing


@dataclass
class PushDispatchResult:
    ok_channels: list
    fail_infos: list
    show_success: bool = False
    show_result: bool = False


def dispatch_push_item(item, send_channel_func, format_message_func=None, label_func=push_label):
    raw_msg = item.get("message", "")
    channels = item.get("channels") or []
    settings = item.get("settings") or {}
    template = item.get("template")
    show_success = bool(item.get("show_success"))
    show_result = bool(item.get("show_result"))
    formatter = format_message if format_message_func is None else format_message_func
    message = formatter(raw_msg, template)

    ok_channels = []
    fail_infos = []
    for channel in channels:
        try:
            ok, info = send_channel_func(channel, message, settings)
        except Exception as exc:
            ok, info = False, str(exc)
        label = label_func(channel)
        if ok:
            ok_channels.append(label)
        else:
            fail_infos.append(f"{label}: {info}")

    return PushDispatchResult(
        ok_channels=ok_channels,
        fail_infos=fail_infos,
        show_success=show_success,
        show_result=show_result,
    )


def push_result_status_message(result):
    ok_channels = result.ok_channels
    fail_infos = result.fail_infos
    if ok_channels and fail_infos:
        return (
            "📡 三方推送部分成功：成功="
            + "、".join(ok_channels)
            + "；失败="
            + "；".join(fail_infos)
        )
    if fail_infos:
        return "📡 三方推送失败：" + "；".join(fail_infos)
    if result.show_success and ok_channels:
        return "📡 三方推送测试成功：" + "、".join(ok_channels)
    return ""
