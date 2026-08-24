THIRD_PUSH_CHANNEL_PARAM_DEFS = {
    "dingtalk": {
        "tip": "如果机器人用了关键词安全设置，请填写 DINGTALK_KEYWORD；加签才需要 Secret。",
        "fields": [
            ("DINGTALK_WEBHOOK：", "dingtalk_webhook", "entry", None),
            ("DINGTALK_SECRET：", "dingtalk_secret", "entry", None),
            ("DINGTALK_KEYWORD：", "dingtalk_keyword", "entry", None),
        ],
    },
    "wecom": {
        "fields": [("WECOM_WEBHOOK：", "wecom_webhook", "entry", None)],
    },
    "feishu": {
        "fields": [("FEISHU_WEBHOOK：", "feishu_webhook", "entry", None)],
    },
    "custom_post": {
        "tip": "Body 里的 {msg} 会替换成推送内容。",
        "fields": [
            ("CUSTOM_POST_URL：", "custom_post_url", "entry", None),
            ("CUSTOM_POST_CONTENT_TYPE：", "custom_post_content_type", "entry", None),
            ("CUSTOM_POST_BODY：", "custom_post_body", "text", None),
        ],
    },
    "telegram": {
        "tip": "TELEGRAM_API 必须填写完整 URL，例如 https://api.telegram.org/bot真实TOKEN/sendMessage。",
        "fields": [
            ("TELEGRAM_API：", "telegram_api", "entry", None),
            ("TELEGRAM_CHAT_ID：", "telegram_chat_id", "entry", None),
        ],
    },
    "pushdeer": {
        "fields": [
            ("PUSHDEER_API：", "pushdeer_api", "entry", None),
            ("PUSHDEER_KEY：", "pushdeer_key", "entry", None),
        ],
    },
    "bark": {
        "fields": [
            ("BARK_API：", "bark_api", "entry", None),
            ("BARK_KEY：", "bark_key", "entry", None),
        ],
    },
    "inotify": {
        "fields": [("INOTIFY_API：", "inotify_api", "entry", None)],
    },
    "pushover": {
        "fields": [
            ("PUSHOVER_API_TOKEN：", "pushover_api_token", "entry", None),
            ("PUSHOVER_USER_KEY：", "pushover_user_key", "entry", None),
        ],
    },
    "gotify": {
        "fields": [
            ("GOTIFY_API：", "gotify_api", "entry", None),
            ("GOTIFY_TOKEN：", "gotify_token", "entry", None),
            ("GOTIFY_TITLE：", "gotify_title", "entry", None),
            ("GOTIFY_PRIORITY：", "gotify_priority", "entry", None),
        ],
    },
    "serverchan": {
        "fields": [
            ("SERVERCHAN_API：", "serverchan_api", "entry", None),
            ("SERVERCHAN_TITLE：", "serverchan_title", "entry", None),
        ],
    },
    "wxpusher": {
        "tip": "WxPusher UID 可填写多个，使用逗号、分号或空格分隔。",
        "fields": [
            ("WXPUSHER_APP_TOKEN：", "wxpusher_app_token", "entry", "*"),
            ("WXPUSHER_UIDS：", "wxpusher_uids", "entry", None),
        ],
    },
    "email": {
        "tip": "支持 SSL（通常端口 465）或 STARTTLS（通常端口 587）。密码建议使用邮箱授权码。",
        "fields": [
            ("EMAIL_SMTP_HOST：", "email_smtp_host", "entry", None),
            ("EMAIL_SMTP_PORT：", "email_smtp_port", "entry", None),
            ("EMAIL_ENCRYPTION：", "email_encryption", "select", ("ssl", "starttls")),
            ("EMAIL_USERNAME：", "email_username", "entry", None),
            ("EMAIL_PASSWORD：", "email_password", "entry", "*"),
            ("EMAIL_FROM_ADDRESS：", "email_from_address", "entry", None),
            ("EMAIL_TO_ADDRESSES：", "email_to_addresses", "entry", None),
            ("EMAIL_SUBJECT：", "email_subject", "entry", None),
        ],
    },
    "next-smtp-proxy": {
        "fields": [
            ("NEXT_SMTP_PROXY_API：", "next_smtp_proxy_api", "entry", None),
            ("NEXT_SMTP_PROXY_USER：", "next_smtp_proxy_user", "entry", None),
            ("NEXT_SMTP_PROXY_PASSWORD：", "next_smtp_proxy_password", "entry", "*"),
            ("NEXT_SMTP_PROXY_HOST：", "next_smtp_proxy_host", "entry", None),
            ("NEXT_SMTP_PROXY_PORT：", "next_smtp_proxy_port", "entry", None),
            ("NEXT_SMTP_PROXY_FORM_NAME：", "next_smtp_proxy_form_name", "entry", None),
            ("NEXT_SMTP_PROXY_TO_EMAIL：", "next_smtp_proxy_to_email", "entry", None),
            ("NEXT_SMTP_PROXY_SUBJECT：", "next_smtp_proxy_subject", "entry", None),
        ],
    },
}
