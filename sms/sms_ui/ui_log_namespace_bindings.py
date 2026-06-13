import os
from datetime import datetime

from sms_core.namespace_binding import make_namespace_runtime_binder
from sms_ui.app_infrastructure_namespace_runtime import schedule_delayed_ui_namespace_runtime
from sms_ui.thread_runtime import run_on_ui_thread, tk_alive_runtime, ui_pump_runtime
from sms_ui.ui_log_namespace_runtime import (
    log_early_namespace_runtime,
    log_file_only_namespace_runtime,
    log_namespace_runtime,
    main_text_available_namespace_runtime,
    safe_insert_main_text_namespace_runtime,
    system_ui_namespace_runtime,
    ui_only_namespace_runtime,
    ui_post_namespace_runtime,
)


def install_ui_log_namespace_bindings(namespace):
    bind = make_namespace_runtime_binder(namespace, globals())

    def tk_alive():
        return tk_alive_runtime(namespace.get("root"), namespace["TK_SHUTDOWN"])

    def ui_pump(max_batch=200):
        return ui_pump_runtime(
            namespace["UI_TASK_QUEUE"],
            namespace["root"],
            namespace["tk_alive"],
            namespace["ui_pump"],
            max_batch=max_batch,
            log_error=namespace.get("log_file_only"),
        )

    def get_log_file():
        today = datetime.now().strftime("%Y-%m-%d")
        return os.path.join(namespace["LOG_DIR"], f"sms_{namespace['LOG_PREFIX']}_{today}.txt")

    def cloud_repeat_filter(notice, msg):
        return notice.next_message(msg)

    def port_ui(message, tag="normal"):
        return namespace["log"](message, tag=tag)

    namespace.update({
        "run_on_ui_thread": run_on_ui_thread,
        "tk_alive": tk_alive,
        "ui_post": bind("ui_post_namespace_runtime"),
        "ui_pump": ui_pump,
        "get_log_file": get_log_file,
        "log_file_only": bind("log_file_only_namespace_runtime"),
        "_cloud_repeat_filter": cloud_repeat_filter,
        "ui_only": bind("ui_only_namespace_runtime"),
        "log_early": bind("log_early_namespace_runtime"),
        "system_ui": bind("system_ui_namespace_runtime"),
        "main_text_available": bind("main_text_available_namespace_runtime"),
        "schedule_delayed_ui": bind("schedule_delayed_ui_namespace_runtime"),
        "port_ui": port_ui,
        "safe_insert_main_text": bind("safe_insert_main_text_namespace_runtime"),
        "log": bind("log_namespace_runtime"),
    })
    return namespace
