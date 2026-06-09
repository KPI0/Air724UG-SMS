DEFAULT_VOICE_TEXT = "注意！四川安播中心预警短信，请及时查看。"

DEFAULT_SERIAL_CONFIG = {
    "port": "",
    "baud": "115200",
    "mode": "Auto",
}

DEFAULT_UI_CONFIG = {
    "voice_enabled": "1",
    "popup_enabled": "1",
    "voice_text": DEFAULT_VOICE_TEXT,
    "allow_multi_instance": "0",
    "auto_log_cleanup": "1",
    "log_unmatched_sms": "1",
    "log_retention_days": "30",
    "desktop_shortcut_name": "短信监听系统",
    "keywords": '["【四川安播中心】"]',
    "sms_font_size": "30",
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
}

THIRD_PUSH_SMS_TEMPLATE = "收到短信：\n{msg}"
THIRD_PUSH_CALL_TEMPLATE = "{msg}"
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
}

