from sms_core.log_cleanup import cleanup_old_logs_in_dir
from sms_core.namespace_binding import make_namespace_runtime_binder
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
    bind = make_namespace_runtime_binder(namespace, globals())

    def cleanup_old_logs(days):
        return cleanup_old_logs_in_dir(namespace["LOG_DIR"], days)

    def clear_text_area_for_new_day():
        namespace["clear_window"]()
        namespace["system_ui"]("📅 新的一天，窗口已清空")
        namespace["schedule_next_midnight_clear"]()

    namespace.update({
        "open_log_dir": bind("open_log_dir_namespace_runtime"),
        "cleanup_old_logs": cleanup_old_logs,
        "open_log_cleanup_dialog": bind("open_log_cleanup_dialog_namespace_runtime"),
        "open_update_proxy_dialog": bind("open_update_proxy_dialog_namespace_runtime"),
        "_auto_log_cleanup_tick": bind("run_auto_log_cleanup_tick_namespace_runtime"),
        "schedule_auto_log_cleanup": bind(
            "schedule_auto_log_cleanup_namespace_runtime",
            positional_keywords=("restart", "first_delay_sec"),
        ),
        "check_update_and_prompt": bind("check_update_and_prompt_namespace_runtime"),
        "clear_text_area_for_new_day": clear_text_area_for_new_day,
        "schedule_next_midnight_clear": bind("schedule_next_midnight_clear_namespace_runtime"),
    })
    return namespace
