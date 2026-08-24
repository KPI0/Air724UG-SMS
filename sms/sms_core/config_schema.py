from sms_core.cloud_command_security import CLOUD_COMMAND_PERMISSION_DEFAULTS


DEFAULT_VOICE_TEXT = "注意！四川安播中心预警短信，请及时查看。"
DEFAULT_LOG_RETENTION_DAYS = 30
DEFAULT_SMS_FONT_SIZE = 30
SMS_FONT_SIZE_MIN = 8
SMS_FONT_SIZE_MAX = 72


def normalize_log_retention_days(value, default=DEFAULT_LOG_RETENTION_DAYS):
    try:
        days = int(value)
    except (TypeError, ValueError, OverflowError):
        return DEFAULT_LOG_RETENTION_DAYS
    if days < 0:
        try:
            fallback = int(default)
        except (TypeError, ValueError, OverflowError):
            fallback = DEFAULT_LOG_RETENTION_DAYS
        return fallback if fallback >= 0 else DEFAULT_LOG_RETENTION_DAYS
    return days


def normalize_sms_font_size(value, default=DEFAULT_SMS_FONT_SIZE):
    try:
        size = int(value)
    except (TypeError, ValueError, OverflowError):
        return DEFAULT_SMS_FONT_SIZE
    if SMS_FONT_SIZE_MIN <= size <= SMS_FONT_SIZE_MAX:
        return size
    try:
        fallback = int(default)
    except (TypeError, ValueError, OverflowError):
        fallback = DEFAULT_SMS_FONT_SIZE
    if SMS_FONT_SIZE_MIN <= fallback <= SMS_FONT_SIZE_MAX:
        return fallback
    return DEFAULT_SMS_FONT_SIZE

DEFAULT_SERIAL_CONFIG = {
    "port": "",
    "baud": "115200",
    "mode": "Auto",
}

DEFAULT_UI_CONFIG = {
    "voice_enabled": "1",
    "popup_enabled": "1",
    "call_popup_enabled": "1",
    "voice_text": DEFAULT_VOICE_TEXT,
    "allow_multi_instance": "0",
    "auto_log_cleanup": "1",
    "log_unmatched_sms": "1",
    "log_retention_days": str(DEFAULT_LOG_RETENTION_DAYS),
    "desktop_shortcut_name": "短信监听系统",
    "keywords": '["【四川安播中心】"]',
    "sms_font_size": str(DEFAULT_SMS_FONT_SIZE),
    "sms_font_color": "#ff0000",
    "call_filter_mode": "Disabled",
    "call_whitelist": "[]",
    "call_blacklist": "[]",
}

DEFAULT_UPDATE_CONFIG = {
    "api_proxy_base": "https://github-api.daybyday.top/",
    "proxy_base": "https://gh-proxy.com/",
}

DEFAULT_CLOUD_CONTROL_CONFIG = {
    "enabled": "0",
    "url": "",
    "device_secret": "",
    "reconnect_interval": "5",
    "auto_upload": "0",
    "allow_sensitive_commands": "0",
    **CLOUD_COMMAND_PERMISSION_DEFAULTS,
}

THIRD_PUSH_SMS_TEMPLATE = "{msg}\n\n发件号码：{sender}\n本机号码：{local_number}\n时间：{sms_time}"
THIRD_PUSH_CALL_TEMPLATE = "收到来电：{caller}\n\n本机号码：{local_number}\n时间：{call_time}"
THIRD_PUSH_DEFAULTS = {
    "enabled": "0",
    "sms_enabled": "1",
    "call_enabled": "1",
    "notify_type": "[]",
    "custom_post_url": "",
    "custom_post_content_type": "application/json",
    "custom_post_body": '{"title":"短信提醒","desp":"{msg}"}',
    "telegram_api": "https://api.telegram.org/bot<BOT_TOKEN>/sendMessage",
    "telegram_chat_id": "",
    "pushdeer_api": "https://api2.pushdeer.com/message/push",
    "pushdeer_key": "",
    "bark_api": "https://api.day.app",
    "bark_key": "",
    "dingtalk_webhook": "",
    "dingtalk_secret": "",
    "dingtalk_keyword": "",
    "feishu_webhook": "",
    "wecom_webhook": "",
    "pushover_api_token": "",
    "pushover_user_key": "",
    "inotify_api": "",
    "next_smtp_proxy_api": "",
    "next_smtp_proxy_user": "",
    "next_smtp_proxy_password": "",
    "next_smtp_proxy_host": "smtp-mail.outlook.com",
    "next_smtp_proxy_port": "587",
    "next_smtp_proxy_form_name": "Air724UG",
    "next_smtp_proxy_to_email": "",
    "next_smtp_proxy_subject": "来自 Air724UG 的通知",
    "gotify_api": "",
    "gotify_title": "Air724UG",
    "gotify_priority": "8",
    "gotify_token": "",
    "serverchan_title": "来自 Air724UG 的通知",
    "serverchan_api": "",
    "wxpusher_app_token": "",
    "wxpusher_uids": "",
    "email_smtp_host": "",
    "email_smtp_port": "465",
    "email_encryption": "ssl",
    "email_username": "",
    "email_password": "",
    "email_from_address": "",
    "email_to_addresses": "",
    "email_subject": "Air724UG 通知",
}
