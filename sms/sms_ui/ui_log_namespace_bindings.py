import os
from datetime import datetime

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
    def tk_alive():
        return tk_alive_runtime(namespace.get("root"), namespace["TK_SHUTDOWN"])

    def ui_post(fn, *args, **kwargs):
        return ui_post_namespace_runtime(namespace, fn, *args, **kwargs)

    def ui_pump(max_batch=200):
        return ui_pump_runtime(
            namespace["UI_TASK_QUEUE"],
            namespace["root"],
            namespace["tk_alive"],
            namespace["ui_pump"],
            max_batch=max_batch,
        )

    def get_log_file():
        today = datetime.now().strftime("%Y-%m-%d")
        return os.path.join(namespace["LOG_DIR"], f"sms_{namespace['LOG_PREFIX']}_{today}.txt")

    def log_file_only(msg):
        return log_file_only_namespace_runtime(namespace, msg)

    def cloud_repeat_filter(notice, msg):
        return notice.next_message(msg)

    def ui_only(msg, tag="normal"):
        return ui_only_namespace_runtime(namespace, msg, tag)

    def log_early(msg, tag="normal"):
        return log_early_namespace_runtime(namespace, msg, tag)

    def system_ui(message, tag="normal"):
        return system_ui_namespace_runtime(namespace, message, tag)

    def main_text_available():
        return main_text_available_namespace_runtime(namespace)

    def schedule_delayed_ui(callback):
        return schedule_delayed_ui_namespace_runtime(namespace, callback)

    def port_ui(message, tag="normal"):
        return namespace["log"](message, tag=tag)

    def safe_insert_main_text(msg, tag="normal"):
        return safe_insert_main_text_namespace_runtime(namespace, msg, tag)

    def log(msg, tag="normal"):
        return log_namespace_runtime(namespace, msg, tag)

    namespace.update({
        "run_on_ui_thread": run_on_ui_thread,
        "tk_alive": tk_alive,
        "ui_post": ui_post,
        "ui_pump": ui_pump,
        "get_log_file": get_log_file,
        "log_file_only": log_file_only,
        "_cloud_repeat_filter": cloud_repeat_filter,
        "ui_only": ui_only,
        "log_early": log_early,
        "system_ui": system_ui,
        "main_text_available": main_text_available,
        "schedule_delayed_ui": schedule_delayed_ui,
        "port_ui": port_ui,
        "safe_insert_main_text": safe_insert_main_text,
        "log": log,
    })
    return namespace
