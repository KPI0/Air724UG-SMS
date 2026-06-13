from sms_ui.thread_runtime import ui_post_runtime
from sms_ui.ui_log_runtime import (
    flush_pending_ui_logs_runtime,
    insert_main_text_runtime,
    log_file_only_runtime,
    run_log_runtime,
    system_ui_runtime,
    ui_only_runtime,
)


def ui_post_namespace_runtime(namespace, fn, *args, ui_post_runtime_func=ui_post_runtime, **kwargs):
    log_error = namespace.get("log_file_only")
    return ui_post_runtime_func(
        namespace["UI_TASK_QUEUE"],
        fn,
        args,
        kwargs,
        on_full=lambda: log_error("⚠️ UI_TASK_QUEUE 已满：丢弃一次 UI 任务") if log_error else None,
        log_error=log_error,
    )


def log_file_only_namespace_runtime(namespace, msg):
    return log_file_only_runtime(
        msg,
        log_dir=namespace["LOG_DIR"],
        file_log=namespace["FILE_LOG_Q"].put_nowait,
    )


def ui_only_namespace_runtime(namespace, msg, tag="normal"):
    return ui_only_runtime(
        msg,
        tag,
        pending_logs=namespace["PENDING_UI_LOGS"],
        has_text=namespace["main_text_available"],
        insert_text=namespace["safe_insert_main_text"],
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
    )


def log_early_namespace_runtime(namespace, msg, tag="normal"):
    namespace["log_file_only"](msg)
    try:
        namespace["PENDING_UI_LOGS"].put_nowait((msg, tag))
    except Exception:
        pass


def system_ui_namespace_runtime(namespace, message, tag="normal"):
    return system_ui_runtime(
        message,
        tag,
        tk_alive=namespace["tk_alive"],
        log_file_only=namespace["log_file_only"],
        pending_logs=namespace["PENDING_UI_LOGS"],
        has_text=namespace["main_text_available"],
        insert_text=namespace["safe_insert_main_text"],
        schedule_ui=namespace["schedule_delayed_ui"],
    )


def main_text_available_namespace_runtime(namespace):
    try:
        return (
            "text_area" in namespace
            and namespace["text_area"] is not None
            and namespace["text_area"].winfo_exists()
        )
    except Exception:
        return False


def safe_insert_main_text_namespace_runtime(namespace, msg, tag="normal"):
    return insert_main_text_runtime(
        namespace.get("text_area"),
        msg,
        tag,
        end_marker=namespace["tk"].END,
    )


def flush_pending_ui_logs_namespace_runtime(namespace):
    return flush_pending_ui_logs_runtime(
        namespace["PENDING_UI_LOGS"],
        namespace["safe_insert_main_text"],
    )


def log_namespace_runtime(namespace, msg, tag="normal"):
    def ui_and_file():
        return run_log_runtime(
            msg,
            tag,
            has_text=namespace["main_text_available"],
            insert_text=namespace["safe_insert_main_text"],
            log_early=namespace["log_early"],
            log_dir=namespace["LOG_DIR"],
            log_prefix=namespace["LOG_PREFIX"],
            file_log=namespace["FILE_LOG_Q"].put_nowait,
        )

    return namespace["run_on_ui_thread"](ui_and_file, namespace["ui_post"])
