import json

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
    return "❌ 配置保存失败，设置可能不会在重启后保留"


def save_keywords_config(config, keywords, safe_save, *, log_error=None):
    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        config.set("ui", "keywords", json.dumps(keywords, ensure_ascii=False))
        return _save_or_raise(safe_save)
    except Exception as exc:
        _safe_log(log_error, f"Save keywords config failed: {exc!r}")
        return False


def save_ui_config_values(config, values, safe_save, *, log_error=None):
    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        for key, value in values.items():
            config.set("ui", str(key), str(value))
        return _save_or_raise(safe_save)
    except Exception as exc:
        _safe_log(log_error, f"Save UI config values failed: {exc!r}")
        return False


def keyword_change_message(action, value=None, old_value=None):
    if action == "add":
        return f"💬 关键词 增加：{value}"
    if action == "delete":
        return f"💬 关键词 删除：{value}"
    if action == "edit":
        return f"💬 关键词 修改：{old_value} -> {value}"
    return ""


def save_log_unmatched_config(config, enabled, safe_save, *, log_error=None):
    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        config.set("ui", "log_unmatched_sms", "1" if enabled else "0")
        return _save_or_raise(safe_save)
    except Exception as exc:
        _safe_log(log_error, f"Save unmatched SMS config failed: {exc!r}")
        return False


def log_unmatched_status(enabled):
    return f"⚙️ 未匹配短信写入COM日志：{'已开启' if enabled else '已关闭'}"


def save_multi_instance_config(config, enabled, safe_save, *, log_error=None):
    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        config.set("ui", "allow_multi_instance", "1" if enabled else "0")
        return _save_or_raise(safe_save)
    except Exception as exc:
        _safe_log(log_error, f"Save multi-instance config failed: {exc!r}")
        return False


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


def toggle_voice_broadcast_runtime(current_enabled, config, safe_save, set_enabled, update_label, system_ui, *, log_error=None):
    enabled = not bool(current_enabled)
    set_enabled(enabled)
    update_label()
    if save_ui_config_values(config, {"voice_enabled": "1" if enabled else "0"}, safe_save, log_error=log_error):
        system_ui(voice_broadcast_status(enabled), "normal")
    else:
        system_ui(_save_failed_status(), "normal")
    return enabled


def toggle_popup_runtime(enabled, config, safe_save, set_enabled, system_ui, *, log_error=None):
    enabled = bool(enabled)
    set_enabled(enabled)
    if save_ui_config_values(config, {"popup_enabled": "1" if enabled else "0"}, safe_save, log_error=log_error):
        system_ui(popup_status(enabled), "normal")
    else:
        system_ui(_save_failed_status(), "normal")
    return enabled


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
    set_multi_instance(enabled)
    if save_multi_instance_config(config, enabled, safe_save, log_error=log_error):
        system_ui(multi_instance_status(enabled), "normal")
    else:
        system_ui(_save_failed_status(), "normal")
    if enabled:
        show_multi_instance_notice(show_notice, log_error=log_error)


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
        set_voice_text(text)
        if save_ui_config_values(config, {"voice_text": text}, safe_save, log_error=log_error):
            generate_voice(force=True)
            system_ui("🔊 已更新语音播报内容：" + text, "normal")
        else:
            system_ui(_save_failed_status(), "normal")

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
        set_font(size, color)
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
            apply_font_style()
            system_ui(f"🎨 已更新短信字体：字号 {size}，颜色 {color}", "normal")
        else:
            system_ui(_save_failed_status(), "normal")

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
):
    def on_keywords_changed(action, value=None, old_value=None):
        if not save_keywords_config(config, keywords, safe_save, log_error=log_error):
            system_ui(_save_failed_status(), "normal")
            return
        message = keyword_change_message(action, value=value, old_value=old_value)
        if message:
            system_ui(message)

    def on_log_unmatched_changed(enabled):
        enabled = bool(enabled)
        set_log_unmatched(enabled)
        if save_log_unmatched_config(config, enabled, safe_save, log_error=log_error):
            system_ui(log_unmatched_status(enabled), "normal")
        else:
            system_ui(_save_failed_status(), "normal")

    open_keywords_setting_dialog(
        parent,
        keywords,
        log_unmatched,
        on_keywords_changed,
        on_log_unmatched_changed,
        center_window,
    )


def save_call_filter_list(config, config_key, target_list, safe_save, *, log_error=None):
    try:
        if "ui" not in config:
            config["ui"] = {}
        config.set("ui", config_key, json.dumps(target_list, ensure_ascii=False))
        return _save_or_raise(safe_save)
    except Exception as exc:
        _safe_log(log_error, f"Save call filter list failed: {exc!r}")
        return False


def save_call_filter_mode(config, mode, safe_save, *, log_error=None):
    try:
        if "ui" not in config:
            config["ui"] = {}
        config.set("ui", "call_filter_mode", mode)
        return _save_or_raise(safe_save)
    except Exception as exc:
        _safe_log(log_error, f"Save call filter mode failed: {exc!r}")
        return False


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
):
    def on_mode_changed(next_mode):
        set_mode(next_mode)
        if save_call_filter_mode(config, next_mode, safe_save, log_error=log_error):
            system_ui(call_filter_mode_status(next_mode))
        else:
            system_ui(_save_failed_status(), "normal")

    def on_list_changed(list_kind, action, value=None, old_value=None):
        if list_kind == "whitelist":
            saved = save_call_filter_list(config, "call_whitelist", whitelist, safe_save, log_error=log_error)
        else:
            saved = save_call_filter_list(config, "call_blacklist", blacklist, safe_save, log_error=log_error)
        if not saved:
            system_ui(_save_failed_status(), "normal")
            return

        message = call_filter_list_status(list_kind, action, value=value, old_value=old_value)
        if message:
            system_ui(message)

    open_call_filter_setting_dialog(
        parent,
        mode,
        whitelist,
        blacklist,
        on_mode_changed,
        on_list_changed,
        center_window,
    )
