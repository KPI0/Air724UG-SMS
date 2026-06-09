from sms_core.log_cleanup import cleanup_old_logs_in_dir
from sms_ui.app_infrastructure_namespace_runtime import schedule_next_midnight_clear_namespace_runtime
from sms_ui.maintenance_namespace_runtime import (
    check_update_and_prompt_namespace_runtime,
    open_log_cleanup_dialog_namespace_runtime,
    open_log_dir_namespace_runtime,
    open_update_proxy_dialog_namespace_runtime,
    run_auto_log_cleanup_tick_namespace_runtime,
    schedule_auto_log_cleanup_namespace_runtime,
)


def install_maintenance_namespace_bindings(namespace):
    def open_log_dir():
        return open_log_dir_namespace_runtime(namespace)

    def cleanup_old_logs(days):
        return cleanup_old_logs_in_dir(namespace["LOG_DIR"], days)

    def open_log_cleanup_dialog():
        return open_log_cleanup_dialog_namespace_runtime(namespace)

    def open_update_proxy_dialog():
        return open_update_proxy_dialog_namespace_runtime(namespace)

    def auto_log_cleanup_tick():
        return run_auto_log_cleanup_tick_namespace_runtime(namespace)

    def schedule_auto_log_cleanup(restart=True, first_delay_sec=60):
        return schedule_auto_log_cleanup_namespace_runtime(
            namespace,
            restart=restart,
            first_delay_sec=first_delay_sec,
        )

    def check_update_and_prompt():
        return check_update_and_prompt_namespace_runtime(namespace)

    def clear_text_area_for_new_day():
        namespace["clear_window"]()
        namespace["system_ui"]("📅 新的一天，窗口已清空")
        namespace["schedule_next_midnight_clear"]()

    def schedule_next_midnight_clear():
        return schedule_next_midnight_clear_namespace_runtime(namespace)

    namespace.update({
        "open_log_dir": open_log_dir,
        "cleanup_old_logs": cleanup_old_logs,
        "open_log_cleanup_dialog": open_log_cleanup_dialog,
        "open_update_proxy_dialog": open_update_proxy_dialog,
        "_auto_log_cleanup_tick": auto_log_cleanup_tick,
        "schedule_auto_log_cleanup": schedule_auto_log_cleanup,
        "check_update_and_prompt": check_update_and_prompt,
        "clear_text_area_for_new_day": clear_text_area_for_new_day,
        "schedule_next_midnight_clear": schedule_next_midnight_clear,
    })
    return namespace
