import json

from sms_core.cloud_command_security import (
    CLOUD_COMMAND_PERMISSION_SPECS,
    LEGACY_PERMISSION_OPTIONS,
    cloud_sensitive_commands_status,
    normalize_cloud_command_permissions,
)
from sms_core.config_runtime import restore_config_section, snapshot_config_section
from sms_ui.sms_font_dialog import validated_tk_color
from sms_ui.settings_dialogs import (
    open_call_filter_setting_dialog,
    open_keywords_setting_dialog,
)


def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def _save_or_raise(save_func):
    result = save_func()
    if result is False:
        raise RuntimeError("配置保存失败")
    return True


def _save_failed_status():
    return "❌ 配置保存失败，已保留原设置"


def _save_ui_values(config, values, safe_save, error_label, log_error=None):
    snapshot = snapshot_config_section(config, "ui")
    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        for key, value in values.items():
            config.set("ui", str(key), str(value))
        return _save_or_raise(safe_save)
    except Exception as exc:
        restore_config_section(config, "ui", snapshot)
        _safe_log(log_error, f"{error_label}: {exc!r}")
        return False


def save_keywords_config(config, keywords, safe_save, *, log_error=None):
    return _save_ui_values(
        config,
        {"keywords": json.dumps(keywords, ensure_ascii=False)},
        safe_save,
        "Save keywords config failed",
        log_error,
    )


def save_ui_config_values(config, values, safe_save, *, log_error=None):
    return _save_ui_values(
        config,
        values,
        safe_save,
        "Save UI config values failed",
        log_error,
    )


def save_cloud_sensitive_commands_config(config, permissions, safe_save, *, log_error=None):
    snapshot = snapshot_config_section(config, "cloud_control")
    try:
        if not config.has_section("cloud_control"):
            config["cloud_control"] = {}
        normalized = normalize_cloud_command_permissions(permissions)
        config.set("cloud_control", "allow_sensitive_commands", "0")
        for option in LEGACY_PERMISSION_OPTIONS:
            config.set("cloud_control", option, "0")
        for spec in CLOUD_COMMAND_PERMISSION_SPECS:
            config.set(
                "cloud_control",
                spec.option,
                "1" if normalized[spec.category] else "0",
            )
        return _save_or_raise(safe_save)
    except Exception as exc:
        restore_config_section(config, "cloud_control", snapshot)
        _safe_log(log_error, f"Save cloud command security config failed: {exc!r}")
        return False


def open_security_settings_runtime(
    parent,
    current_permissions,
    *,
    config,
    safe_save,
    set_permissions,
    system_ui,
    center_window,
    open_dialog,
    log_error=None,
):
    def change(permissions):
        normalized = normalize_cloud_command_permissions(permissions)
        if not save_cloud_sensitive_commands_config(
            config,
            normalized,
            safe_save,
            log_error=log_error,
        ):
            system_ui(_save_failed_status(), "normal")
            return False
        set_permissions(normalized)
        system_ui("🔐 " + cloud_sensitive_commands_status(normalized), "normal")
        return True

    return open_dialog(
        parent,
        normalize_cloud_command_permissions(current_permissions),
        change,
        center_window,
    )


def keyword_change_message(action, value=None, old_value=None):
    if action == "add":
        return f"💬 关键词 增加：{value}"
    if action == "delete":
        return f"💬 关键词 删除：{value}"
    if action == "edit":
        return f"💬 关键词 修改：{old_value} -> {value}"
    return ""


def save_log_unmatched_config(config, enabled, safe_save, *, log_error=None):
    return _save_ui_values(
        config,
        {"log_unmatched_sms": "1" if enabled else "0"},
        safe_save,
        "Save unmatched SMS config failed",
        log_error,
    )


def log_unmatched_status(enabled):
    return f"⚙️ 未匹配短信写入COM日志：{'已开启' if enabled else '已关闭'}"


def save_multi_instance_config(config, enabled, safe_save, *, log_error=None):
    return _save_ui_values(
        config,
        {"allow_multi_instance": "1" if enabled else "0"},
        safe_save,
        "Save multi-instance config failed",
        log_error,
    )


def multi_instance_status(enabled):
    return "✅️ 程序多开：已开启" if enabled else "❌ 程序多开：已关闭"


MULTI_INSTANCE_NOTICE_TITLE = "程序多开提醒"
MULTI_INSTANCE_NOTICE_MESSAGE = (
    "开启程序多开后，通过当前程序目录启动的多个程序实例会读取同一个 config.ini 配置文件。\n\n"
    "短信关键词、云端控制、三方推送等配置都会共享；如需独立配置，请复制到不同目录分别使用。"
)


def show_multi_instance_notice(show_notice, *, log_error=None):
    if show_notice is None:
        return
    try:
        show_notice(MULTI_INSTANCE_NOTICE_TITLE, MULTI_INSTANCE_NOTICE_MESSAGE)
    except Exception as exc:
        _safe_log(log_error, f"Show multi-instance notice failed: {exc!r}")


def voice_broadcast_status(enabled):
    return "🔊 语音播报：已开启" if enabled else "🔇 语音播报：已关闭"


def popup_status(enabled):
    return "✅️ 短信弹窗：已开启" if enabled else "❌ 短信弹窗：已关闭"


def call_popup_status(enabled):
    return "✅️ 电话弹窗：已开启" if enabled else "❌ 电话弹窗：已关闭"


def toggle_voice_broadcast_runtime(current_enabled, config, safe_save, set_enabled, update_label, system_ui, *, log_error=None):
    enabled = not bool(current_enabled)
    if save_ui_config_values(config, {"voice_enabled": "1" if enabled else "0"}, safe_save, log_error=log_error):
        set_enabled(enabled)
        update_label()
        system_ui(voice_broadcast_status(enabled), "normal")
        return enabled
    else:
        system_ui(_save_failed_status(), "normal")
        return bool(current_enabled)


def toggle_popup_runtime(enabled, config, safe_save, set_enabled, system_ui, *, log_error=None):
    enabled = bool(enabled)
    if save_ui_config_values(config, {"popup_enabled": "1" if enabled else "0"}, safe_save, log_error=log_error):
        set_enabled(enabled)
        system_ui(popup_status(enabled), "normal")
        return enabled
    else:
        system_ui(_save_failed_status(), "normal")
        return None


def toggle_call_popup_runtime(
    enabled,
    config,
    safe_save,
    set_enabled,
    system_ui,
    *,
    log_error=None,
):
    enabled = bool(enabled)
    if save_ui_config_values(
        config,
        {"call_popup_enabled": "1" if enabled else "0"},
        safe_save,
        log_error=log_error,
    ):
        set_enabled(enabled)
        system_ui(call_popup_status(enabled), "normal")
        return enabled
    system_ui(_save_failed_status(), "normal")
    return None


def toggle_multi_instance_runtime(
    enabled,
    config,
    safe_save,
    set_multi_instance,
    system_ui,
    *,
    show_notice=None,
    log_error=None,
):
    enabled = bool(enabled)
    if save_multi_instance_config(config, enabled, safe_save, log_error=log_error):
        set_multi_instance(enabled)
        system_ui(multi_instance_status(enabled), "normal")
        if enabled:
            show_multi_instance_notice(show_notice, log_error=log_error)
        return enabled
    else:
        system_ui(_save_failed_status(), "normal")
        return None


def open_voice_text_dialog_runtime(
    parent,
    current_text,
    *,
    config,
    safe_save,
    set_voice_text,
    generate_voice,
    system_ui,
    center_window,
    open_dialog,
    log_error=None,
):
    def preview(text):
        generate_voice(force=True, text=text, play_after=True)

    def save(text):
        if save_ui_config_values(config, {"voice_text": text}, safe_save, log_error=log_error):
            set_voice_text(text)
            generate_voice(force=True)
            system_ui("🔊 已更新语音播报内容：" + text, "normal")
            return True
        else:
            system_ui(_save_failed_status(), "normal")
            return False

    open_dialog(parent, current_text, preview, save, center_window)


def open_sms_font_dialog_runtime(
    parent,
    current_size,
    current_color,
    *,
    config,
    safe_save,
    set_font,
    apply_font_style,
    system_ui,
    center_window,
    open_dialog,
    log_error=None,
):
    def save_font(size, color):
        color = validated_tk_color(parent, color)
        if color is None:
            system_ui("❌ 短信字体颜色无效，请输入有效颜色", "normal")
            return False

        saved = save_ui_config_values(
            config,
            {
                "sms_font_size": size,
                "sms_font_color": color,
            },
            safe_save,
            log_error=log_error,
        )
        if saved:
            set_font(size, color)
            apply_font_style()
            system_ui(f"🎨 已更新短信字体：字号 {size}，颜色 {color}", "normal")
        else:
            system_ui(_save_failed_status(), "normal")
        return saved

    open_dialog(parent, current_size, current_color, save_font, center_window)


def open_keywords_setting_runtime(
    parent,
    keywords,
    log_unmatched,
    config,
    safe_save,
    system_ui,
    set_log_unmatched,
    center_window,
    log_error=None,
    *,
    register_external_refresh=None,
    get_log_unmatched=None,
):
    def on_keywords_changed(action, value=None, old_value=None):
        if not save_keywords_config(config, keywords, safe_save, log_error=log_error):
            system_ui(_save_failed_status(), "normal")
            return False
        message = keyword_change_message(action, value=value, old_value=old_value)
        if message:
            system_ui(message)
        return True

    def on_log_unmatched_changed(enabled):
        enabled = bool(enabled)
        if save_log_unmatched_config(config, enabled, safe_save, log_error=log_error):
            set_log_unmatched(enabled)
            system_ui(log_unmatched_status(enabled), "normal")
            return True
        else:
            system_ui(_save_failed_status(), "normal")
            return False

    open_keywords_setting_dialog(
        parent,
        keywords,
        log_unmatched,
        on_keywords_changed,
        on_log_unmatched_changed,
        center_window,
        register_external_refresh=register_external_refresh,
        get_log_unmatched=get_log_unmatched,
    )


def save_call_filter_list(config, config_key, target_list, safe_save, *, log_error=None):
    return _save_ui_values(
        config,
        {config_key: json.dumps(target_list, ensure_ascii=False)},
        safe_save,
        "Save call filter list failed",
        log_error,
    )


def save_call_filter_mode(config, mode, safe_save, *, log_error=None):
    return _save_ui_values(
        config,
        {"call_filter_mode": mode},
        safe_save,
        "Save call filter mode failed",
        log_error,
    )


def call_filter_mode_label(mode):
    return {
        "Disabled": "关闭过滤",
        "Whitelist": "白名单模式",
        "Blacklist": "黑名单模式",
    }.get(mode, mode)


def call_filter_mode_status(mode):
    return f"📞 防骚扰模式已切换为：{call_filter_mode_label(mode)}"


def call_filter_list_status(list_kind, action, value=None, old_value=None):
    list_name = "白名单" if list_kind == "whitelist" else "黑名单"
    if action == "add":
        return f"📞 {list_name} 增加：{value}"
    if action == "delete":
        return f"📞 {list_name} 删除：{value}"
    if action == "edit":
        return f"📞 {list_name} 修改：{old_value} -> {value}"
    return ""


def open_call_filter_setting_runtime(
    parent,
    mode,
    whitelist,
    blacklist,
    config,
    safe_save,
    system_ui,
    set_mode,
    center_window,
    log_error=None,
    *,
    register_external_refresh=None,
    get_mode=None,
):
    def on_mode_changed(next_mode):
        if save_call_filter_mode(config, next_mode, safe_save, log_error=log_error):
            set_mode(next_mode)
            system_ui(call_filter_mode_status(next_mode))
            return True
        else:
            system_ui(_save_failed_status(), "normal")
            return False

    def on_list_changed(list_kind, action, value=None, old_value=None):
        if list_kind == "whitelist":
            saved = save_call_filter_list(config, "call_whitelist", whitelist, safe_save, log_error=log_error)
        else:
            saved = save_call_filter_list(config, "call_blacklist", blacklist, safe_save, log_error=log_error)
        if not saved:
            system_ui(_save_failed_status(), "normal")
            return False

        message = call_filter_list_status(list_kind, action, value=value, old_value=old_value)
        if message:
            system_ui(message)
        return True

    open_call_filter_setting_dialog(
        parent,
        mode,
        whitelist,
        blacklist,
        on_mode_changed,
        on_list_changed,
        center_window,
        register_external_refresh=register_external_refresh,
        get_mode=get_mode,
    )
