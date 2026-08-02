from sms_core.namespace_binding import make_namespace_runtime_binder
from sms_ui.settings_namespace_runtime import (
    open_call_filter_setting_namespace_runtime,
    open_desktop_shortcut_dialog_namespace_runtime,
    open_keywords_setting_namespace_runtime,
    open_security_settings_namespace_runtime,
    open_serial_setting_namespace_runtime,
    open_sms_font_dialog_namespace_runtime,
    open_voice_text_dialog_namespace_runtime,
)
from sms_ui.settings_runtime import save_ui_config_values


def install_settings_namespace_bindings(namespace):
    bind = make_namespace_runtime_binder(namespace, globals())

    def save_voice_text_setting():
        return save_ui_config_values(
            namespace["config"],
            {"voice_text": namespace["VOICE_TEXT"]},
            namespace["safe_save_config"],
            log_error=namespace.get("log_file_only"),
        )

    def update_voice_menu_label():
        try:
            label = "🔊 语音播报" if namespace["VOICE_ENABLED"] else "🔇 语音播报"
            namespace["menu_bar"].entryconfig(namespace["voice_menu_index"], label=label)
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if log_error is not None:
                try:
                    log_error(f"Update voice menu label failed: {exc!r}")
                except Exception:
                    pass

    def save_voice_setting():
        return save_ui_config_values(
            namespace["config"],
            {"voice_enabled": "1" if namespace["VOICE_ENABLED"] else "0"},
            namespace["safe_save_config"],
            log_error=namespace.get("log_file_only"),
        )

    def open_security_settings():
        return open_security_settings_namespace_runtime(namespace)

    namespace.update({
        "save_voice_text_setting": save_voice_text_setting,
        "open_sms_font_dialog": bind("open_sms_font_dialog_namespace_runtime"),
        "open_voice_text_dialog": bind("open_voice_text_dialog_namespace_runtime"),
        "open_serial_setting": bind("open_serial_setting_namespace_runtime"),
        "open_desktop_shortcut_dialog": bind("open_desktop_shortcut_dialog_namespace_runtime"),
        "open_keywords_setting": bind("open_keywords_setting_namespace_runtime"),
        "open_call_filter_setting": bind("open_call_filter_setting_namespace_runtime"),
        "open_security_settings": open_security_settings,
        "update_voice_menu_label": update_voice_menu_label,
        "save_voice_setting": save_voice_setting,
    })
    return namespace
