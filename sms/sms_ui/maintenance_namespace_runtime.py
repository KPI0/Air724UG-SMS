import os
import webbrowser

from sms_ui.maintenance_runtime import (
    open_log_cleanup_dialog_app_runtime,
    open_log_dir_runtime,
    open_update_proxy_dialog_runtime,
    run_auto_log_cleanup_tick_app_runtime,
    schedule_auto_log_cleanup_app_runtime,
)
from sms_ui.update_check_runtime import check_update_and_prompt_runtime, read_update_config_runtime


def open_log_dir_namespace_runtime(namespace, *, open_dir_runtime=open_log_dir_runtime):
    os_module = namespace.get("os", os)
    return open_dir_runtime(
        namespace["LOG_DIR"],
        path_abspath=os_module.path.abspath,
        path_exists=os_module.path.exists,
        open_path=os_module.startfile,
        show_message=namespace["ui_messagebox"],
    )


def open_log_cleanup_dialog_namespace_runtime(
    namespace,
    *,
    open_dialog_runtime=open_log_cleanup_dialog_app_runtime,
):
    return open_dialog_runtime(
        namespace["root"],
        get_retention_days=lambda: namespace["LOG_RETENTION_DAYS"],
        get_interval_hours=lambda: namespace["AUTO_CLEANUP_INTERVAL_HOURS"],
        config=namespace["config"],
        save_config=namespace["safe_save_config"],
        set_cleanup_state=lambda days, enabled: (
            namespace.__setitem__("LOG_RETENTION_DAYS", days),
            namespace.__setitem__("AUTO_LOG_CLEANUP", enabled),
        ),
        system_ui=namespace["system_ui"],
        schedule_cleanup=namespace["schedule_auto_log_cleanup"],
        center_window=namespace["center_window"],
    )


def open_update_proxy_dialog_namespace_runtime(
    namespace,
    *,
    open_dialog_runtime=open_update_proxy_dialog_runtime,
):
    return open_dialog_runtime(
        namespace["root"],
        config=namespace["config"],
        owner=namespace["GITHUB_OWNER"],
        repo=namespace["GITHUB_REPO"],
        save_config=namespace["safe_save_config"],
        ui_post=namespace["ui_post"],
        center_window=namespace["center_window"],
        log_error=namespace.get("log_file_only"),
    )


def run_auto_log_cleanup_tick_namespace_runtime(
    namespace,
    *,
    run_tick_runtime=run_auto_log_cleanup_tick_app_runtime,
):
    return run_tick_runtime(
        state=namespace["AUTO_LOG_CLEANUP_STATE"],
        is_enabled=lambda: namespace["AUTO_LOG_CLEANUP"],
        retention_days=lambda: namespace["LOG_RETENTION_DAYS"],
        interval_hours=lambda: namespace["AUTO_CLEANUP_INTERVAL_HOURS"],
        cleanup_old_logs=namespace["cleanup_old_logs"],
        system_ui=namespace["system_ui"],
        tk_alive=namespace["tk_alive"],
        root_after=namespace["root"].after,
        tick_callback=namespace["_auto_log_cleanup_tick"],
        ui_post=namespace["ui_post"],
    )


def schedule_auto_log_cleanup_namespace_runtime(
    namespace,
    *,
    restart=True,
    first_delay_sec=60,
    schedule_runtime=schedule_auto_log_cleanup_app_runtime,
):
    return schedule_runtime(
        state=namespace["AUTO_LOG_CLEANUP_STATE"],
        restart=restart,
        first_delay_sec=first_delay_sec,
        is_enabled=lambda: namespace["AUTO_LOG_CLEANUP"],
        tk_alive=namespace["tk_alive"],
        root_after=namespace["root"].after,
        root_after_cancel=namespace["root"].after_cancel,
        tick_callback=namespace["_auto_log_cleanup_tick"],
        ui_post=namespace["ui_post"],
    )


def check_update_and_prompt_namespace_runtime(
    namespace,
    *,
    check_runtime=check_update_and_prompt_runtime,
    read_config_runtime=read_update_config_runtime,
):
    webbrowser_module = namespace.get("webbrowser", webbrowser)
    return check_runtime(
        owner=namespace["GITHUB_OWNER"],
        repo=namespace["GITHUB_REPO"],
        current_version=namespace["APP_VERSION"],
        get_update_config=lambda: read_config_runtime(namespace["config"]),
        ui_post=namespace["ui_post"],
        show_info=namespace["messagebox"].showinfo,
        show_warning=namespace["messagebox"].showwarning,
        show_error=namespace["messagebox"].showerror,
        ask_open_download=namespace["messagebox"].askyesno,
        open_url=webbrowser_module.open,
        log_error=namespace.get("log_file_only"),
    )
