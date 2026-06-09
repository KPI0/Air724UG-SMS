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


def _read_keywords(config, coerce_text_list):
    try:
        raw = config.get("ui", "keywords", fallback="").strip()
        if not raw:
            return []
        if raw.startswith("[") and raw.endswith("]"):
            try:
                return coerce_text_list(json.loads(raw))
            except Exception:
                return coerce_text_list([x.strip() for x in raw.split("|") if x.strip()])
        return coerce_text_list([x.strip() for x in raw.split("|") if x.strip()])
    except Exception:
        return []


def _read_json_text_list(config, key, coerce_text_list):
    try:
        raw = config.get("ui", key, fallback="").strip()
        if raw:
            return coerce_text_list(json.loads(raw))
    except Exception:
        pass
    return []


def read_startup_config_values(config, *, default_voice_text, coerce_text_list=_coerce_text_list_default):
    try:
        voice_text = config.get("ui", "voice_text", fallback=default_voice_text).strip()
        if not voice_text:
            voice_text = default_voice_text
    except Exception:
        voice_text = default_voice_text

    try:
        popup_enabled = config.getboolean("ui", "popup_enabled", fallback=True)
    except Exception:
        popup_enabled = True

    try:
        auto_log_cleanup = config.getboolean("ui", "auto_log_cleanup", fallback=True)
    except Exception:
        auto_log_cleanup = True

    try:
        log_retention_days = config.getint("ui", "log_retention_days", fallback=30)
    except Exception:
        log_retention_days = 30

    try:
        allow_multi_instance = config.getboolean("ui", "allow_multi_instance", fallback=False)
    except Exception:
        allow_multi_instance = False

    try:
        log_unmatched_sms = config.getboolean("ui", "log_unmatched_sms", fallback=False)
    except Exception:
        log_unmatched_sms = False

    try:
        voice_enabled = config.getboolean("ui", "voice_enabled", fallback=True)
    except Exception:
        voice_enabled = True

    try:
        sms_font_size = config.getint("ui", "sms_font_size", fallback=30)
    except Exception:
        sms_font_size = 30

    try:
        sms_font_color = config.get("ui", "sms_font_color", fallback="#ff0000").strip() or "#ff0000"
    except Exception:
        sms_font_color = "#ff0000"

    keywords = _read_keywords(config, coerce_text_list)

    try:
        call_filter_mode_raw = config.get("ui", "call_filter_mode", fallback="Disabled").strip()
    except Exception:
        call_filter_mode_raw = "Disabled"
    call_filter_mode = {
        "disabled": "Disabled",
        "whitelist": "Whitelist",
        "blacklist": "Blacklist",
    }.get(call_filter_mode_raw.lower(), "Disabled")

    call_whitelist = _read_json_text_list(config, "call_whitelist", coerce_text_list)
    call_blacklist = _read_json_text_list(config, "call_blacklist", coerce_text_list)

    port = config.get("serial", "port", fallback="").strip()
    baud = config.getint("serial", "baud", fallback=115200)
    mode = config.get("serial", "mode", fallback="Auto").strip().lower()
    if mode not in ("auto", "manual"):
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
