from sms_core.third_push_config import ensure_third_push_config_values, read_third_push_settings
from sms_core.config_runtime import (
    reload_config_runtime,
    restore_config_section,
    snapshot_config_section,
)
from sms_ui.third_push_app_runtime import (
    enqueue_third_push_app_runtime,
    open_third_push_values_app_runtime,
    save_third_push_setting_runtime,
    show_third_push_test_result_app_runtime,
    third_push_settings_from_values,
    third_push_worker_app_runtime,
)


def ensure_third_push_config_namespace_runtime(namespace, *, save=False):
    config_snapshot = snapshot_config_section(namespace["config"], "third_push")
    changed = ensure_third_push_config_values(namespace["config"])
    if changed and save:
        try:
            if namespace["safe_save_config"]() is False:
                raise RuntimeError("配置保存失败")
        except Exception:
            restore_config_section(namespace["config"], "third_push", config_snapshot)
            return False
    return changed


def apply_third_push_settings_namespace_runtime(namespace, settings):
    namespace.__setitem__("THIRD_PUSH_ENABLED", settings.enabled)
    namespace.__setitem__("THIRD_PUSH_SMS_ENABLED", settings.sms_enabled)
    namespace.__setitem__("THIRD_PUSH_CALL_ENABLED", settings.call_enabled)
    namespace.__setitem__("THIRD_PUSH_TYPES", list(settings.channels or []))
    namespace.__setitem__("THIRD_PUSH_SETTINGS", dict(settings.settings or {}))


def _log_config_reload_failure(namespace, exc):
    log_error = namespace.get("log_file_only")
    if log_error is None:
        return
    try:
        log_error(f"Reload third-push config failed ({type(exc).__name__})")
    except Exception:
        pass


def refresh_third_push_settings_namespace_runtime(
    namespace,
    *,
    reload_config=reload_config_runtime,
):
    defaults_changed = False

    def prepare_config(staged_config):
        nonlocal defaults_changed
        defaults_changed = ensure_third_push_config_values(staged_config)

    def commit_config():
        if not defaults_changed:
            return True
        return namespace["safe_save_config"]()

    try:
        settings = reload_config(
            config=namespace["config"],
            config_file=namespace["CONFIG_FILE"],
            config_lock=namespace["CONFIG_LOCK"],
            prepare_config=prepare_config,
            commit_config=commit_config,
            read_values=read_third_push_settings,
        )
    except Exception as exc:
        _log_config_reload_failure(namespace, exc)
        return False

    apply_third_push_settings_namespace_runtime(namespace, settings)
    return True


def save_third_push_setting_namespace_runtime(
    namespace,
    *,
    enabled=None,
    sms_enabled=None,
    call_enabled=None,
    notify_type=None,
    settings=None,
):
    return save_third_push_setting_runtime(
        current_settings=lambda: third_push_settings_from_values(
            namespace["THIRD_PUSH_ENABLED"],
            namespace["THIRD_PUSH_SMS_ENABLED"],
            namespace["THIRD_PUSH_CALL_ENABLED"],
            namespace["THIRD_PUSH_TYPES"],
            namespace["THIRD_PUSH_SETTINGS"],
        ),
        apply_settings=lambda next_settings: apply_third_push_settings_namespace_runtime(namespace, next_settings),
        config=namespace["config"],
        save_config=namespace["safe_save_config"],
        enabled=enabled,
        sms_enabled=sms_enabled,
        call_enabled=call_enabled,
        notify_type=notify_type,
        settings=settings,
    )


def third_push_worker_namespace_runtime(namespace):
    return third_push_worker_app_runtime(
        stop_event=namespace["third_push_stop"],
        push_queue=namespace["THIRD_PUSH_Q"],
        get_log_prefix=lambda: namespace["LOG_PREFIX"],
        app_version=namespace["APP_VERSION"],
        system_ui=namespace["system_ui"],
        show_result=namespace["show_third_push_test_result"],
        shutdown_event=namespace.get("TK_SHUTDOWN"),
    )


def show_third_push_test_result_namespace_runtime(namespace, ok_channels, fail_infos):
    return show_third_push_test_result_app_runtime(
        root=namespace["root"],
        get_current_window=lambda: namespace["third_push_win"],
        messagebox=namespace["messagebox"],
        ui_post=namespace["ui_post"],
        ok_channels=ok_channels,
        fail_infos=fail_infos,
    )


def enqueue_third_push_namespace_runtime(
    namespace,
    raw_msg,
    *,
    show_success=False,
    show_result=False,
    channels=None,
    settings=None,
    template=None,
    variables=None,
    event_type="sms",
):
    merged_variables = dict(variables or {})
    local_number = str(namespace.get("LOCAL_NUMBER") or "").strip()
    if local_number:
        if not str(merged_variables.get("local_number") or "").strip():
            merged_variables["local_number"] = local_number
        if not str(merged_variables.get("self_number") or "").strip():
            merged_variables["self_number"] = local_number
    return enqueue_third_push_app_runtime(
        raw_msg,
        push_queue=namespace["THIRD_PUSH_Q"],
        enabled=lambda: namespace["THIRD_PUSH_ENABLED"],
        sms_enabled=lambda: namespace["THIRD_PUSH_SMS_ENABLED"],
        call_enabled=lambda: namespace["THIRD_PUSH_CALL_ENABLED"],
        configured_channels=lambda: namespace["THIRD_PUSH_TYPES"],
        current_settings=lambda: namespace["THIRD_PUSH_SETTINGS"],
        channels=channels,
        settings=settings,
        template=template,
        variables=merged_variables,
        event_type=event_type,
        show_success=show_success,
        show_result=show_result,
        system_ui=namespace["system_ui"],
    )


def open_third_push_window_namespace_runtime(namespace):
    def sync_existing_window(win, attr):
        return namespace["sync_and_focus_existing_window"](
            win,
            attr,
            log_error=namespace.get("log_file_only"),
        )

    return open_third_push_values_app_runtime(
        parent=namespace["root"],
        current_window=lambda: namespace["third_push_win"],
        enabled=lambda: namespace["THIRD_PUSH_ENABLED"],
        sms_enabled=lambda: namespace["THIRD_PUSH_SMS_ENABLED"],
        call_enabled=lambda: namespace["THIRD_PUSH_CALL_ENABLED"],
        channels=lambda: namespace["THIRD_PUSH_TYPES"],
        settings=lambda: namespace["THIRD_PUSH_SETTINGS"],
        refresh_settings=namespace["refresh_third_push_settings_from_config"],
        save_setting=namespace["save_third_push_setting"],
        enqueue_push=namespace["enqueue_third_push"],
        system_ui=namespace["system_ui"],
        sync_existing_window=sync_existing_window,
        set_window=lambda win: namespace.__setitem__("third_push_win", win),
        center_window=namespace["center_window"],
    )
