import os
import sys

from sms_core.app_launch import restart_software_runtime
from sms_core.threading_runtime import start_daemon_thread
from sms_ui.app_shutdown_runtime import start_shutdown_task_runtime
from sms_ui.shutdown_progress import ShutdownProgress


def restart_software_app_runtime(
    *,
    root,
    messagebox,
    is_exiting,
    set_exiting,
    set_serial_running,
    autostart_flag,
    restart_helper_flag,
    log_error,
    system_ui,
    stop_tray_icon,
    safe_set_events,
    stop_events,
    stop_cloud_control,
    safe_close_serial,
    app_mutex,
    release_mutex,
    flush_log_queue,
    file_log_queue,
    exit_process=os._exit,
    argv=None,
    current_pid=None,
    restart_runtime=restart_software_runtime,
    file_log_thread=None,
    file_log_stop_event=None,
    worker_threads=(),
    deferred_stop_events=(),
    deferred_worker_threads=(),
    deferred_worker_queues=(),
    pre_cloud_worker_threads=(),
    before_stop_cloud=None,
    report_progress=None,
):
    return restart_runtime(
        is_exiting=is_exiting,
        confirm_restart=lambda: messagebox.askyesno("重启软件", "确定要重启软件吗？", parent=root),
        argv=sys.argv if argv is None else argv,
        autostart_flag=autostart_flag,
        restart_helper_flag=restart_helper_flag,
        current_pid=os.getpid() if current_pid is None else current_pid,
        log_error=log_error,
        show_launch_error=lambda error: messagebox.showerror(
            "重启失败",
            f"启动重启助手失败，当前软件将继续运行。\n\n{error}",
            parent=root,
        ),
        set_exiting=set_exiting,
        system_ui=system_ui,
        stop_tray_icon=stop_tray_icon,
        set_serial_running=set_serial_running,
        safe_set_events=safe_set_events,
        stop_events=stop_events,
        stop_cloud_control=stop_cloud_control,
        safe_close_serial=safe_close_serial,
        app_mutex=app_mutex,
        release_mutex=release_mutex,
        flush_log_queue=flush_log_queue,
        file_log_queue=file_log_queue,
        file_log_thread=file_log_thread,
        file_log_stop_event=file_log_stop_event,
        worker_threads=worker_threads,
        deferred_stop_events=deferred_stop_events,
        deferred_worker_threads=deferred_worker_threads,
        deferred_worker_queues=deferred_worker_queues,
        exit_process=exit_process,
        pre_cloud_worker_threads=pre_cloud_worker_threads,
        before_stop_cloud=before_stop_cloud,
        report_progress=report_progress,
    )


def start_restart_software_app_runtime(
    *, root, messagebox, is_exiting, set_exiting, log_error,
    exit_process=os._exit, argv=None, current_pid=None,
    restart_runtime=restart_software_runtime, start_worker=start_daemon_thread,
    progress_factory=None, notify_log_errors=None, **restart_options,
):
    def run_task(report_progress):
        return restart_runtime(
            is_exiting=False, confirm_restart=lambda: True, set_exiting=lambda _value: None,
            argv=sys.argv if argv is None else argv,
            current_pid=os.getpid() if current_pid is None else current_pid,
            log_error=log_error, show_launch_error=lambda _error: None,
            exit_process=lambda _code: None, report_progress=report_progress,
            **restart_options,
        )

    def failure_message(result):
        if getattr(result, "status", None) == "launch_failed":
            return "启动重启助手失败，当前软件将继续运行。请检查日志后重试。"
        return "后台清理未能完成，窗口已保留。请再次选择重启或退出，并检查日志中的错误。"

    return start_shutdown_task_runtime(
        root=root, messagebox=messagebox, is_exiting=is_exiting, set_exiting=set_exiting,
        run_task=run_task, on_success=lambda: exit_process(0), action="重启",
        confirmation="确定要重启软件吗？", start_worker=start_worker,
        progress_factory=progress_factory or (lambda parent: ShutdownProgress(parent, title="正在重启")),
        log_error=log_error, failure_message=failure_message,
        notify_log_errors=notify_log_errors,
    )
