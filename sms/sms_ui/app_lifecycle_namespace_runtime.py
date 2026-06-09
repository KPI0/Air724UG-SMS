import os

from sms_ui.app_autostart_runtime import set_autostart_runtime
from sms_ui.app_restart_runtime import restart_software_app_runtime
from sms_ui.app_shutdown_runtime import cleanup_and_exit_app_runtime
from sms_ui.settings_runtime import (
    toggle_multi_instance_runtime,
    toggle_popup_runtime,
    toggle_voice_broadcast_runtime,
)


def set_autostart_namespace_runtime(namespace, enable, *, set_autostart_app_runtime=set_autostart_runtime):
    return set_autostart_app_runtime(
        enable,
        autostart_flag=namespace["AUTOSTART_FLAG"],
        create_startup_shortcut=namespace["create_startup_shortcut"],
        remove_startup_shortcut=namespace["remove_startup_shortcut"],
        system_ui=namespace["system_ui"],
        show_error=lambda title, message: namespace["ui_messagebox"]("error", title, message),
    )


def cleanup_and_exit_namespace_runtime(namespace, *, cleanup_app_runtime=cleanup_and_exit_app_runtime):
    return cleanup_app_runtime(
        root=namespace["root"],
        messagebox=namespace["messagebox"],
        is_exiting=namespace["is_exiting"],
        set_exiting=lambda value: namespace.__setitem__("is_exiting", bool(value)),
        set_serial_running=lambda value: namespace.__setitem__("serial_running", bool(value)),
        shutdown_events=(namespace["TK_SHUTDOWN"],),
        worker_stop_events=(
            namespace["file_log_stop"],
            namespace["third_push_stop"],
            namespace["serial_stop_event"],
            namespace["serial_wakeup_event"],
        ),
        tts_stop_event=namespace["TTS_STOP"],
        safe_set_events=namespace["safe_set_events"],
        stop_cloud_control=namespace["stop_cloud_control"],
        safe_close_serial=namespace["safe_close_serial"],
        stop_tray_icon=namespace["stop_tray_icon"],
        flush_log_queue=namespace["flush_log_queue"],
        file_log_queue=namespace["FILE_LOG_Q"],
        destroy_root=namespace["root"].destroy,
    )


def toggle_voice_broadcast_namespace_runtime(
    namespace,
    *,
    toggle_runtime=toggle_voice_broadcast_runtime,
):
    return toggle_runtime(
        namespace["VOICE_ENABLED"],
        namespace["config"],
        namespace["safe_save_config"],
        lambda value: namespace.__setitem__("VOICE_ENABLED", bool(value)),
        namespace["update_voice_menu_label"],
        namespace["system_ui"],
    )


def toggle_multi_instance_namespace_runtime(
    namespace,
    *,
    toggle_runtime=toggle_multi_instance_runtime,
):
    return toggle_runtime(
        namespace["multi_instance_var"].get(),
        namespace["config"],
        namespace["safe_save_config"],
        lambda value: namespace.__setitem__("ALLOW_MULTI_INSTANCE", bool(value)),
        namespace["system_ui"],
    )


def toggle_popup_namespace_runtime(namespace, *, toggle_runtime=toggle_popup_runtime):
    return toggle_runtime(
        namespace["popup_var"].get(),
        namespace["config"],
        namespace["safe_save_config"],
        lambda value: namespace.__setitem__("POPUP_ENABLED", bool(value)),
        namespace["system_ui"],
    )


def restart_software_namespace_runtime(
    namespace,
    *,
    restart_app_runtime=restart_software_app_runtime,
):
    os_module = namespace.get("os", os)
    return restart_app_runtime(
        root=namespace["root"],
        messagebox=namespace["messagebox"],
        is_exiting=namespace["is_exiting"],
        set_exiting=lambda value: namespace.__setitem__("is_exiting", bool(value)),
        set_serial_running=lambda value: namespace.__setitem__("serial_running", bool(value)),
        autostart_flag=namespace["AUTOSTART_FLAG"],
        restart_helper_flag=namespace["RESTART_HELPER_FLAG"],
        log_error=namespace["log_file_only"],
        system_ui=namespace["system_ui"],
        stop_tray_icon=namespace["stop_tray_icon"],
        safe_set_events=namespace["safe_set_events"],
        stop_events=(
            namespace["third_push_stop"],
            namespace["serial_stop_event"],
            namespace["serial_wakeup_event"],
            namespace["file_log_stop"],
            namespace["TTS_STOP"],
        ),
        stop_cloud_control=namespace["stop_cloud_control"],
        safe_close_serial=namespace["safe_close_serial"],
        app_mutex=namespace["app_mutex"],
        release_mutex=namespace["release_mutex_handle"],
        flush_log_queue=namespace["flush_log_queue"],
        file_log_queue=namespace["FILE_LOG_Q"],
        exit_process=os_module._exit,
    )
