import os
import threading
import json
from dataclasses import dataclass


@dataclass(frozen=True)
class StartupConfigValues:
    voice_text: str
    popup_enabled: bool
    auto_log_cleanup: bool
    log_retention_days: int
    allow_multi_instance: bool
    log_unmatched_sms: bool
    voice_enabled: bool
    sms_font_size: int
    sms_font_color: str
    keywords: list
    call_filter_mode: str
    call_whitelist: list
    call_blacklist: list
    port: str
    baud: int
    mode: str


def initialize_config_runtime(
    *,
    config,
    config_file,
    defaults_by_section,
    save_config,
    path_exists=os.path.exists,
    encoding="utf-8-sig",
):
    created = False
    if not path_exists(config_file):
        for section, values in defaults_by_section.items():
            config[section] = dict(values)
        save_config()
        created = True

    config.read(config_file, encoding=encoding)
    return created


def _coerce_text_list_default(value):
    if isinstance(value, str):
        value = [value]
    if not isinstance(value, (list, tuple)):
        return []
    return [str(item).strip() for item in value if str(item).strip()]


def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def _read_config_value(label, fallback, reader, log_error=None):
    try:
        return reader()
    except Exception as exc:
        _safe_log(log_error, f"Config value {label} invalid; using default {fallback!r}: {exc!r}")
        return fallback


def _read_keywords(config, coerce_text_list, log_error=None):
    try:
        raw = config.get("ui", "keywords", fallback="").strip()
        if not raw:
            return []
        if raw.startswith("[") and raw.endswith("]"):
            try:
                return coerce_text_list(json.loads(raw))
            except Exception as exc:
                _safe_log(log_error, f"Config value ui.keywords JSON invalid; using pipe fallback: {exc!r}")
                return coerce_text_list([x.strip() for x in raw.split("|") if x.strip()])
        return coerce_text_list([x.strip() for x in raw.split("|") if x.strip()])
    except Exception as exc:
        _safe_log(log_error, f"Config value ui.keywords invalid; using default []: {exc!r}")
        return []


def _read_json_text_list(config, key, coerce_text_list, log_error=None):
    try:
        raw = config.get("ui", key, fallback="").strip()
        if raw:
            return coerce_text_list(json.loads(raw))
    except Exception as exc:
        _safe_log(log_error, f"Config value ui.{key} JSON invalid; using default []: {exc!r}")
    return []


def read_startup_config_values(
    config,
    *,
    default_voice_text,
    coerce_text_list=_coerce_text_list_default,
    log_error=None,
):
    voice_text = _read_config_value(
        "ui.voice_text",
        default_voice_text,
        lambda: config.get("ui", "voice_text", fallback=default_voice_text).strip(),
        log_error,
    )
    if not voice_text:
        voice_text = default_voice_text

    popup_enabled = _read_config_value(
        "ui.popup_enabled",
        True,
        lambda: config.getboolean("ui", "popup_enabled", fallback=True),
        log_error,
    )
    auto_log_cleanup = _read_config_value(
        "ui.auto_log_cleanup",
        True,
        lambda: config.getboolean("ui", "auto_log_cleanup", fallback=True),
        log_error,
    )
    log_retention_days = _read_config_value(
        "ui.log_retention_days",
        30,
        lambda: config.getint("ui", "log_retention_days", fallback=30),
        log_error,
    )
    allow_multi_instance = _read_config_value(
        "ui.allow_multi_instance",
        False,
        lambda: config.getboolean("ui", "allow_multi_instance", fallback=False),
        log_error,
    )
    log_unmatched_sms = _read_config_value(
        "ui.log_unmatched_sms",
        False,
        lambda: config.getboolean("ui", "log_unmatched_sms", fallback=False),
        log_error,
    )
    voice_enabled = _read_config_value(
        "ui.voice_enabled",
        True,
        lambda: config.getboolean("ui", "voice_enabled", fallback=True),
        log_error,
    )
    sms_font_size = _read_config_value(
        "ui.sms_font_size",
        30,
        lambda: config.getint("ui", "sms_font_size", fallback=30),
        log_error,
    )
    sms_font_color = _read_config_value(
        "ui.sms_font_color",
        "#ff0000",
        lambda: config.get("ui", "sms_font_color", fallback="#ff0000").strip() or "#ff0000",
        log_error,
    )

    keywords = _read_keywords(config, coerce_text_list, log_error=log_error)

    call_filter_mode_raw = _read_config_value(
        "ui.call_filter_mode",
        "Disabled",
        lambda: config.get("ui", "call_filter_mode", fallback="Disabled").strip(),
        log_error,
    )
    call_filter_mode = {
        "disabled": "Disabled",
        "whitelist": "Whitelist",
        "blacklist": "Blacklist",
    }.get(call_filter_mode_raw.lower(), "Disabled")

    call_whitelist = _read_json_text_list(config, "call_whitelist", coerce_text_list, log_error=log_error)
    call_blacklist = _read_json_text_list(config, "call_blacklist", coerce_text_list, log_error=log_error)

    port = _read_config_value(
        "serial.port",
        "",
        lambda: config.get("serial", "port", fallback="").strip(),
        log_error,
    )
    baud = _read_config_value(
        "serial.baud",
        115200,
        lambda: config.getint("serial", "baud", fallback=115200),
        log_error,
    )
    if baud <= 0:
        _safe_log(log_error, f"Config value serial.baud must be positive; using default 115200: {baud!r}")
        baud = 115200
    mode = _read_config_value(
        "serial.mode",
        "Auto",
        lambda: config.get("serial", "mode", fallback="Auto").strip(),
        log_error,
    ).lower()
    if mode not in ("auto", "manual"):
        _safe_log(log_error, f"Config value serial.mode unknown; using default 'Auto': {mode!r}")
        mode = "auto"
    mode = "Auto" if mode == "auto" else "Manual"

    return StartupConfigValues(
        voice_text=voice_text,
        popup_enabled=popup_enabled,
        auto_log_cleanup=auto_log_cleanup,
        log_retention_days=log_retention_days,
        allow_multi_instance=allow_multi_instance,
        log_unmatched_sms=log_unmatched_sms,
        voice_enabled=voice_enabled,
        sms_font_size=sms_font_size,
        sms_font_color=sms_font_color,
        keywords=keywords,
        call_filter_mode=call_filter_mode,
        call_whitelist=call_whitelist,
        call_blacklist=call_blacklist,
        port=port,
        baud=baud,
        mode=mode,
    )


def safe_save_config_runtime(
    *,
    config,
    config_file,
    config_lock,
    log_error=None,
    getpid=os.getpid,
    get_thread_id=threading.get_ident,
    open_file=open,
    replace_file=os.replace,
    path_exists=os.path.exists,
    remove_file=os.remove,
):
    tmp_file = f"{config_file}.{getpid()}.{get_thread_id()}.tmp"
    try:
        with config_lock:
            with open_file(tmp_file, "w", encoding="utf-8") as file_obj:
                config.write(file_obj)
            replace_file(tmp_file, config_file)
        return True
    except Exception as exc:
        try:
            if path_exists(tmp_file):
                remove_file(tmp_file)
        except Exception:
            pass
        try:
            if log_error is not None:
                log_error(f"配置保存失败: {exc}")
        except Exception:
            pass
        return False
