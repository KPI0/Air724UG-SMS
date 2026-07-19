import configparser
import hashlib
import os
import threading
import json
import time
from dataclasses import dataclass

from sms_core.windows_runtime import acquire_named_mutex_lock, release_named_mutex_lock


CONFIG_SNAPSHOT_ATTR = "_sms_last_disk_snapshot"
CONFIG_MUTEX_PREFIX = "Air724UG_SMS_Config_Write_V1"


class ConfigInitializationError(RuntimeError):
    """Raised when startup defaults cannot be persisted safely."""


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


def snapshot_config_section(config, section):
    section = str(section)
    if not config.has_section(section):
        return None
    return {
        option: config.get(section, option, raw=True)
        for option in config.options(section)
    }


def restore_config_section(config, section, snapshot):
    section = str(section)
    if config.has_section(section):
        config.remove_section(section)
    if snapshot is None:
        return
    config.add_section(section)
    for option, value in snapshot.items():
        config.set(section, str(option), str(value))


def snapshot_config_runtime(config):
    return {
        section: {
            option: config.get(section, option, raw=True)
            for option in config.options(section)
        }
        for section in config.sections()
    }


def restore_config_runtime(config, snapshot):
    config.clear()
    for section, values in dict(snapshot or {}).items():
        config.add_section(str(section))
        for option, value in dict(values or {}).items():
            config.set(str(section), str(option), str(value))


def remember_config_snapshot(config, snapshot=None):
    try:
        setattr(
            config,
            CONFIG_SNAPSHOT_ATTR,
            snapshot_config_runtime(config) if snapshot is None else dict(snapshot),
        )
    except Exception:
        pass


def config_mutex_name(config_file):
    normalized = os.path.normcase(os.path.abspath(str(config_file or "config.ini")))
    digest = hashlib.sha256(normalized.encode("utf-8", "surrogatepass")).hexdigest()[:16]
    return f"{CONFIG_MUTEX_PREFIX}_{digest}"


def merge_config_changes(disk_snapshot, baseline_snapshot, current_snapshot):
    merged = {
        section: dict(values)
        for section, values in dict(disk_snapshot or {}).items()
    }
    baseline = dict(baseline_snapshot or {})
    current = dict(current_snapshot or {})

    for section in set(baseline) | set(current):
        if section not in current:
            merged.pop(section, None)
            continue
        if section not in baseline:
            merged[section] = dict(current[section])
            continue

        target = merged.setdefault(section, {})
        old_values = dict(baseline[section])
        new_values = dict(current[section])
        for option in set(old_values) | set(new_values):
            if option not in new_values:
                target.pop(option, None)
            elif option not in old_values or new_values[option] != old_values[option]:
                target[option] = new_values[option]

    return merged


def load_config_snapshot(config_file, *, encoding="utf-8-sig"):
    disk_config = configparser.ConfigParser(interpolation=None)
    disk_config.read(config_file, encoding=encoding)
    return snapshot_config_runtime(disk_config)


def initialize_config_runtime(
    *,
    config,
    config_file,
    defaults_by_section,
    save_config,
    path_exists=os.path.exists,
    encoding="utf-8-sig",
    backup_file=os.replace,
    time_func=time.time,
    log_error=None,
):
    created = False

    def load_defaults():
        config.clear()
        for section, values in defaults_by_section.items():
            config[section] = dict(values)

    def save_startup_defaults(action):
        try:
            result = save_config()
        except Exception as exc:
            _safe_log(log_error, f"Config {action} save raised an exception: {exc!r}")
            raise ConfigInitializationError(
                f"配置文件{action}失败，无法写入：{config_file}"
            ) from exc
        if result is False:
            _safe_log(log_error, f"Config {action} save returned False: {config_file!r}")
            raise ConfigInitializationError(
                f"配置文件{action}失败，无法写入：{config_file}"
            )

    if not path_exists(config_file):
        load_defaults()
        save_startup_defaults("创建")
        created = True

    try:
        config.read(config_file, encoding=encoding)
    except configparser.Error as exc:
        backup_path = f"{config_file}.broken.{int(time_func())}.bak"
        try:
            if path_exists(config_file):
                backup_file(config_file, backup_path)
                _safe_log(log_error, f"Config file invalid; moved to {backup_path!r}: {exc!r}")
        except Exception as backup_exc:
            _safe_log(log_error, f"Config file invalid and backup failed; recreating defaults: {backup_exc!r}")
        load_defaults()
        save_startup_defaults("修复")
        created = True
    remember_config_snapshot(config)
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
    acquire_process_lock=acquire_named_mutex_lock,
    release_process_lock=release_named_mutex_lock,
    load_snapshot=load_config_snapshot,
    lock_timeout_ms=10000,
):
    tmp_file = f"{config_file}.{getpid()}.{get_thread_id()}.tmp"
    process_lock = None
    try:
        with config_lock:
            process_lock, lock_result = acquire_process_lock(
                config_mutex_name(config_file),
                timeout_ms=lock_timeout_ms,
            )
            if not process_lock:
                raise RuntimeError(f"配置文件跨进程锁获取失败，结果码：{lock_result}")

            current_snapshot = snapshot_config_runtime(config)
            baseline_snapshot = getattr(config, CONFIG_SNAPSHOT_ATTR, None)
            if baseline_snapshot is not None and path_exists(config_file):
                disk_snapshot = load_snapshot(config_file)
                output_snapshot = merge_config_changes(
                    disk_snapshot,
                    baseline_snapshot,
                    current_snapshot,
                )
            else:
                output_snapshot = current_snapshot

            output_config = configparser.ConfigParser(interpolation=None)
            restore_config_runtime(output_config, output_snapshot)
            with open_file(tmp_file, "w", encoding="utf-8") as file_obj:
                output_config.write(file_obj)
            replace_file(tmp_file, config_file)
            restore_config_runtime(config, output_snapshot)
            remember_config_snapshot(config, output_snapshot)
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
    finally:
        if process_lock:
            try:
                release_process_lock(process_lock)
            except Exception as exc:
                try:
                    if log_error is not None:
                        log_error(f"配置文件跨进程锁释放失败: {exc}")
                except Exception:
                    pass
