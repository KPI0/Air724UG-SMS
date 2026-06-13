from sms_ui.desktop_shortcut_runtime import open_desktop_shortcut_dialog_runtime
from sms_ui.serial_settings_runtime import open_serial_setting_runtime
from sms_ui.settings_runtime import (
    open_call_filter_setting_runtime,
    open_keywords_setting_runtime,
    open_sms_font_dialog_runtime,
    open_voice_text_dialog_runtime,
)


def open_sms_font_dialog_namespace_runtime(
    namespace,
    *,
    open_dialog_runtime=open_sms_font_dialog_runtime,
):
    return open_dialog_runtime(
        namespace["root"],
        namespace["SMS_FONT_SIZE"],
        namespace["SMS_FONT_COLOR"],
        config=namespace["config"],
        safe_save=namespace["safe_save_config"],
        set_font=lambda size, color: (
            namespace.__setitem__("SMS_FONT_SIZE", size),
            namespace.__setitem__("SMS_FONT_COLOR", color),
        ),
        apply_font_style=namespace["apply_sms_font_style"],
        system_ui=namespace["system_ui"],
        center_window=namespace["center_window"],
        open_dialog=namespace["_ui_open_sms_font_dialog"],
        log_error=namespace.get("log_file_only"),
    )


def open_voice_text_dialog_namespace_runtime(
    namespace,
    *,
    open_dialog_runtime=open_voice_text_dialog_runtime,
):
    return open_dialog_runtime(
        namespace["root"],
        namespace["VOICE_TEXT"],
        config=namespace["config"],
        safe_save=namespace["safe_save_config"],
        set_voice_text=lambda text: namespace.__setitem__("VOICE_TEXT", text),
        generate_voice=namespace["generate_alert_voice"],
        system_ui=namespace["system_ui"],
        center_window=namespace["center_window"],
        open_dialog=namespace["_ui_open_voice_text_dialog"],
        log_error=namespace.get("log_file_only"),
    )


def open_serial_setting_namespace_runtime(
    namespace,
    *,
    open_setting_runtime=open_serial_setting_runtime,
):
    def set_serial_state(mode, port, baud):
        namespace.__setitem__("MODE", mode)
        namespace.__setitem__("PORT", port)
        namespace.__setitem__("BAUD", baud)

    return open_setting_runtime(
        namespace["root"],
        current_mode=namespace["MODE"],
        current_port=namespace["PORT"],
        current_baud=namespace["BAUD"],
        scan_ports=namespace["scan_com_ports_all"],
        center_window=namespace["center_window"],
        config=namespace["config"],
        save_config=namespace["safe_save_config"],
        set_serial_state=set_serial_state,
        set_status=namespace["set_status"],
        safe_close_serial=namespace["safe_close_serial"],
        wake_serial=namespace["serial_wakeup_event"].set,
        system_ui=namespace["system_ui"],
    )


def open_desktop_shortcut_dialog_namespace_runtime(
    namespace,
    *,
    open_dialog_runtime=open_desktop_shortcut_dialog_runtime,
):
    return open_dialog_runtime(
        namespace["root"],
        config=namespace["config"],
        save_config=namespace["safe_save_config"],
        create_shortcut=namespace["create_desktop_shortcut"],
        system_ui=namespace["system_ui"],
        center_window=namespace["center_window"],
        open_dialog=namespace["_ui_open_desktop_shortcut_dialog"],
    )


def open_keywords_setting_namespace_runtime(
    namespace,
    *,
    open_setting_runtime=open_keywords_setting_runtime,
):
    return open_setting_runtime(
        namespace["root"],
        namespace["KEYWORDS"],
        namespace["LOG_UNMATCHED_SMS"],
        namespace["config"],
        namespace["safe_save_config"],
        namespace["system_ui"],
        lambda enabled: namespace.__setitem__("LOG_UNMATCHED_SMS", bool(enabled)),
        namespace["center_window"],
        namespace.get("log_file_only"),
    )


def open_call_filter_setting_namespace_runtime(
    namespace,
    *,
    open_setting_runtime=open_call_filter_setting_runtime,
):
    return open_setting_runtime(
        namespace["root"],
        namespace["CALL_FILTER_MODE"],
        namespace["CALL_WHITELIST"],
        namespace["CALL_BLACKLIST"],
        namespace["config"],
        namespace["safe_save_config"],
        namespace["system_ui"],
        lambda mode: namespace.__setitem__("CALL_FILTER_MODE", mode),
        namespace["center_window"],
        namespace.get("log_file_only"),
    )
