from sms_ui.settings_namespace_runtime import (
    open_call_filter_setting_namespace_runtime,
    open_desktop_shortcut_dialog_namespace_runtime,
    open_keywords_setting_namespace_runtime,
    open_serial_setting_namespace_runtime,
    open_sms_font_dialog_namespace_runtime,
    open_voice_text_dialog_namespace_runtime,
)
from sms_ui.settings_runtime import save_ui_config_values


def install_settings_namespace_bindings(namespace):
    def save_voice_text_setting():
        return save_ui_config_values(
            namespace["config"],
            {"voice_text": namespace["VOICE_TEXT"]},
            namespace["safe_save_config"],
        )

    def open_sms_font_dialog():
        return open_sms_font_dialog_namespace_runtime(namespace)

    def open_voice_text_dialog():
        return open_voice_text_dialog_namespace_runtime(namespace)

    def open_serial_setting():
        return open_serial_setting_namespace_runtime(namespace)

    def open_desktop_shortcut_dialog():
        return open_desktop_shortcut_dialog_namespace_runtime(namespace)

    def open_keywords_setting():
        return open_keywords_setting_namespace_runtime(namespace)

    def open_call_filter_setting():
        return open_call_filter_setting_namespace_runtime(namespace)

    def update_voice_menu_label():
        try:
            label = "🔊 语音播报" if namespace["VOICE_ENABLED"] else "🔇 语音播报"
            namespace["menu_bar"].entryconfig(namespace["voice_menu_index"], label=label)
        except Exception:
            pass

    def save_voice_setting():
        return save_ui_config_values(
            namespace["config"],
            {"voice_enabled": "1" if namespace["VOICE_ENABLED"] else "0"},
            namespace["safe_save_config"],
        )

    namespace.update({
        "save_voice_text_setting": save_voice_text_setting,
        "open_sms_font_dialog": open_sms_font_dialog,
        "open_voice_text_dialog": open_voice_text_dialog,
        "open_serial_setting": open_serial_setting,
        "open_desktop_shortcut_dialog": open_desktop_shortcut_dialog,
        "open_keywords_setting": open_keywords_setting,
        "open_call_filter_setting": open_call_filter_setting,
        "update_voice_menu_label": update_voice_menu_label,
        "save_voice_setting": save_voice_setting,
    })
    return namespace
