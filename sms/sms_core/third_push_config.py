import json
from dataclasses import dataclass

from sms_core.config_schema import THIRD_PUSH_DEFAULTS
from sms_core.third_push import THIRD_PUSH_CHANNEL_LABELS, THIRD_PUSH_SETTINGS_KEYS, parse_push_channels


@dataclass(frozen=True)
class ThirdPushSettings:
    enabled: bool
    sms_enabled: bool
    call_enabled: bool
    channels: list
    settings: dict


def default_bool(key):
    return THIRD_PUSH_DEFAULTS[key] == "1"


def ensure_third_push_config_values(config):
    changed = False
    if not config.has_section("third_push"):
        config["third_push"] = {}
        changed = True
    if config.has_option("third_push", "message_template"):
        config.remove_option("third_push", "message_template")
        changed = True
    for key, value in THIRD_PUSH_DEFAULTS.items():
        if not config.has_option("third_push", key):
            config.set("third_push", key, value)
            changed = True
    return changed


def read_third_push_settings(config):
    ensure_third_push_config_values(config)

    try:
        enabled = config.getboolean("third_push", "enabled", fallback=default_bool("enabled"))
    except Exception:
        enabled = default_bool("enabled")

    try:
        sms_enabled = config.getboolean("third_push", "sms_enabled", fallback=default_bool("sms_enabled"))
    except Exception:
        sms_enabled = default_bool("sms_enabled")

    try:
        call_enabled = config.getboolean("third_push", "call_enabled", fallback=default_bool("call_enabled"))
    except Exception:
        call_enabled = default_bool("call_enabled")

    channels = parse_push_channels(config.get("third_push", "notify_type", fallback="[]"))
    settings = {
        key: config.get("third_push", key, fallback=THIRD_PUSH_DEFAULTS.get(key, ""))
        for key in THIRD_PUSH_SETTINGS_KEYS
    }
    return ThirdPushSettings(enabled, sms_enabled, call_enabled, channels, settings)


def update_third_push_settings(
    current,
    *,
    enabled=None,
    sms_enabled=None,
    call_enabled=None,
    notify_type=None,
    settings=None,
):
    channels = current.channels
    if notify_type is not None:
        channels = [ch for ch in notify_type if ch in THIRD_PUSH_CHANNEL_LABELS]

    next_settings = current.settings
    if settings is not None:
        next_settings = {key: str(settings.get(key, "")) for key in THIRD_PUSH_SETTINGS_KEYS}

    return ThirdPushSettings(
        enabled=current.enabled if enabled is None else bool(enabled),
        sms_enabled=current.sms_enabled if sms_enabled is None else bool(sms_enabled),
        call_enabled=current.call_enabled if call_enabled is None else bool(call_enabled),
        channels=list(channels or []),
        settings=dict(next_settings or {}),
    )


def write_third_push_settings(config, settings):
    ensure_third_push_config_values(config)
    config.set("third_push", "enabled", "1" if settings.enabled else "0")
    config.set("third_push", "sms_enabled", "1" if settings.sms_enabled else "0")
    config.set("third_push", "call_enabled", "1" if settings.call_enabled else "0")
    config.set("third_push", "notify_type", json.dumps(settings.channels, ensure_ascii=False))
    for key in THIRD_PUSH_SETTINGS_KEYS:
        config.set("third_push", key, settings.settings.get(key, ""))
