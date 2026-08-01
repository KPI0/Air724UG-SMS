import time

from sms_core.cloud_command_security import (
    normalize_cloud_command_permissions,
    read_cloud_command_permissions,
)
from sms_core.config_runtime import (
    load_config_snapshot,
    read_startup_config_values,
    reload_config_runtime,
)
from sms_ui.config_sync_runtime import (
    ConfigReloadFailureLogState,
    clear_config_reload_failure_runtime,
    report_config_reload_failure_runtime,
    schedule_config_file_watch_runtime,
)
from sms_ui.call_popup_namespace_runtime import close_phone_popups_namespace_runtime


def _safe_log(namespace, message):
    log_error = namespace.get("log_file_only")
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def _replace_runtime_list(namespace, name, values):
    current = namespace.get(name)
    next_values = list(values)
    if isinstance(current, list):
        changed = current != next_values
        current[:] = next_values
        return changed
    namespace[name] = next_values
    return current != next_values


def register_config_sync_refresher_namespace_runtime(namespace, group, callback):
    registry = namespace.setdefault("_CONFIG_SYNC_REFRESHERS", {})
    callbacks = registry.setdefault(str(group), {})
    token = object()
    callbacks[token] = callback

    def unregister():
        group_callbacks = registry.get(str(group))
        if not group_callbacks:
            return
        group_callbacks.pop(token, None)
        if not group_callbacks:
            registry.pop(str(group), None)

    return unregister


def _notify_config_sync_refreshers(namespace, group):
    registry = namespace.get("_CONFIG_SYNC_REFRESHERS", {})
    callbacks = tuple(registry.get(str(group), {}).values())
    refreshed = 0
    for callback in callbacks:
        try:
            callback()
            refreshed += 1
        except Exception as exc:
            _safe_log(namespace, f"Refresh synced {group} settings window failed: {exc!r}")
    return refreshed


def reload_shared_ui_config_namespace_runtime(
    namespace,
    *,
    load_snapshot=load_config_snapshot,
    read_values=read_startup_config_values,
):
    try:
        values = reload_config_runtime(
            config=namespace["config"],
            config_file=namespace["CONFIG_FILE"],
            config_lock=namespace["CONFIG_LOCK"],
            load_snapshot=load_snapshot,
            read_values=lambda merged_config: read_values(
                merged_config,
                default_voice_text=namespace["DEFAULT_VOICE_TEXT"],
                log_error=namespace.get("log_file_only"),
            ),
        )
    except Exception as exc:
        state = namespace.setdefault(
            "_CONFIG_SYNC_RELOAD_FAILURE_STATE",
            ConfigReloadFailureLogState(),
        )
        time_source = namespace.get("time", time)
        monotonic = getattr(time_source, "monotonic", time.monotonic)
        report_config_reload_failure_runtime(
            state,
            type(exc).__name__,
            log_error=namespace.get("log_file_only"),
            monotonic=monotonic,
        )
        return False

    failure_state = namespace.get("_CONFIG_SYNC_RELOAD_FAILURE_STATE")
    if failure_state is not None:
        clear_config_reload_failure_runtime(
            failure_state,
            log_error=namespace.get("log_file_only"),
        )

    changed_groups = []

    popup_changed = bool(namespace["POPUP_ENABLED"]) != values.popup_enabled
    namespace["POPUP_ENABLED"] = values.popup_enabled
    if popup_changed:
        changed_groups.append("短信弹窗")
        try:
            namespace["popup_var"].set(values.popup_enabled)
        except Exception as exc:
            _safe_log(namespace, f"Sync popup menu state failed: {exc!r}")

    call_popup_changed = (
        bool(namespace.get("CALL_POPUP_ENABLED", True))
        != values.call_popup_enabled
    )
    namespace["CALL_POPUP_ENABLED"] = values.call_popup_enabled
    if call_popup_changed:
        changed_groups.append("电话弹窗")
        try:
            namespace["call_popup_var"].set(values.call_popup_enabled)
        except Exception as exc:
            _safe_log(namespace, f"Sync call popup menu state failed: {exc!r}")
        if not values.call_popup_enabled:
            close_phone_popups_namespace_runtime(namespace)

    voice_enabled_changed = bool(namespace["VOICE_ENABLED"]) != values.voice_enabled
    namespace["VOICE_ENABLED"] = values.voice_enabled
    if voice_enabled_changed:
        changed_groups.append("语音播报")
        try:
            namespace["update_voice_menu_label"]()
        except Exception as exc:
            _safe_log(namespace, f"Sync voice menu label failed: {exc!r}")

    voice_text_changed = namespace["VOICE_TEXT"] != values.voice_text
    namespace["VOICE_TEXT"] = values.voice_text
    if voice_text_changed:
        changed_groups.append("语音内容")
        try:
            namespace["generate_alert_voice"](force=True)
        except Exception as exc:
            _safe_log(namespace, f"Regenerate synced voice alert failed: {exc!r}")

    font_changed = (
        namespace["SMS_FONT_SIZE"] != values.sms_font_size
        or namespace["SMS_FONT_COLOR"] != values.sms_font_color
    )
    namespace["SMS_FONT_SIZE"] = values.sms_font_size
    namespace["SMS_FONT_COLOR"] = values.sms_font_color
    if font_changed:
        changed_groups.append("短信字体")
        try:
            namespace["apply_sms_font_style"]()
        except Exception as exc:
            _safe_log(namespace, f"Apply synced SMS font failed: {exc!r}")

    keywords_changed = _replace_runtime_list(namespace, "KEYWORDS", values.keywords)
    log_unmatched_changed = bool(namespace["LOG_UNMATCHED_SMS"]) != values.log_unmatched_sms
    namespace["LOG_UNMATCHED_SMS"] = values.log_unmatched_sms
    if keywords_changed or log_unmatched_changed:
        changed_groups.append("关键词")
        _notify_config_sync_refreshers(namespace, "keywords")

    call_mode_changed = namespace["CALL_FILTER_MODE"] != values.call_filter_mode
    namespace["CALL_FILTER_MODE"] = values.call_filter_mode
    whitelist_changed = _replace_runtime_list(namespace, "CALL_WHITELIST", values.call_whitelist)
    blacklist_changed = _replace_runtime_list(namespace, "CALL_BLACKLIST", values.call_blacklist)
    if call_mode_changed or whitelist_changed or blacklist_changed:
        changed_groups.append("防骚扰")
        _notify_config_sync_refreshers(namespace, "call_filter")

    cleanup_changed = (
        bool(namespace["AUTO_LOG_CLEANUP"]) != values.auto_log_cleanup
        or namespace["LOG_RETENTION_DAYS"] != values.log_retention_days
    )
    namespace["AUTO_LOG_CLEANUP"] = values.auto_log_cleanup
    namespace["LOG_RETENTION_DAYS"] = values.log_retention_days
    if cleanup_changed:
        changed_groups.append("日志清理")
        try:
            namespace["schedule_auto_log_cleanup"](restart=True, first_delay_sec=60)
        except Exception as exc:
            _safe_log(namespace, f"Reschedule synced log cleanup failed: {exc!r}")

    multi_instance_changed = bool(namespace["ALLOW_MULTI_INSTANCE"]) != values.allow_multi_instance
    namespace["ALLOW_MULTI_INSTANCE"] = values.allow_multi_instance
    if multi_instance_changed:
        changed_groups.append("程序多开")
        try:
            namespace["multi_instance_var"].set(values.allow_multi_instance)
        except Exception as exc:
            _safe_log(namespace, f"Sync multi-instance menu state failed: {exc!r}")

    command_permissions = read_cloud_command_permissions(namespace["config"])
    current_permissions = normalize_cloud_command_permissions(
        namespace.get(
            "CLOUD_SENSITIVE_COMMAND_PERMISSIONS",
            namespace.get("CLOUD_SENSITIVE_COMMANDS_ENABLED", False),
        )
    )
    security_changed = current_permissions != command_permissions
    namespace["CLOUD_SENSITIVE_COMMAND_PERMISSIONS"] = command_permissions
    if security_changed:
        changed_groups.append("安全设置")

    if changed_groups:
        try:
            namespace["system_ui"](
                "⚙️ 已同步其他实例更新的配置：" + "、".join(changed_groups),
                "normal",
            )
        except Exception as exc:
            _safe_log(namespace, f"Report synced config changes failed: {exc!r}")
    return tuple(changed_groups)


def start_config_file_watch_namespace_runtime(
    namespace,
    *,
    interval_ms=1000,
    schedule_runtime=schedule_config_file_watch_runtime,
):
    state = namespace["CONFIG_FILE_WATCH_STATE"]
    return schedule_runtime(
        state=state,
        config_file=namespace["CONFIG_FILE"],
        interval_ms=interval_ms,
        root_after=namespace["root"].after,
        root_after_cancel=namespace["root"].after_cancel,
        tk_alive=namespace["tk_alive"],
        is_stopping=lambda: namespace["TK_SHUTDOWN"].is_set() or bool(namespace.get("is_exiting")),
        on_change=namespace["reload_shared_ui_config"],
        log_error=namespace.get("log_file_only"),
    )
