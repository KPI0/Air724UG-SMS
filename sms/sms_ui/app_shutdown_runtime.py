import queue

from sms_core.app_shutdown import _safe_log, cleanup_and_exit_runtime
from sms_core.threading_runtime import start_daemon_thread
from sms_ui.thread_runtime import ui_post_runtime, ui_pump_runtime
from sms_ui.shutdown_progress import ShutdownProgress


def cleanup_and_exit_app_runtime(
    *,
    root,
    messagebox,
    is_exiting,
    set_exiting,
    set_serial_running,
    shutdown_events,
    worker_stop_events,
    tts_stop_event,
    safe_set_events,
    stop_cloud_control,
    safe_close_serial,
    stop_tray_icon,
    flush_log_queue,
    file_log_queue,
    destroy_root=None,
    cleanup_runtime=cleanup_and_exit_runtime,
    file_log_thread=None,
    file_log_stop_event=None,
    worker_threads=(),
    deferred_worker_stop_events=(),
    deferred_worker_threads=(),
    deferred_worker_queues=(),
    log_error=None,
    pre_cloud_worker_threads=(),
    before_stop_cloud=None,
    report_progress=None,
):
    destroy_root = destroy_root or root.destroy
    return cleanup_runtime(
        is_exiting=is_exiting,
        confirm_exit=lambda: messagebox.askyesno(
            "退出软件",
            "确定要完全退出软件吗？\n\n退出后将停止监听短信和来电。",
            parent=root,
        ),
        set_exiting=set_exiting,
        set_serial_running=set_serial_running,
        shutdown_events=shutdown_events,
        worker_stop_events=worker_stop_events,
        tts_stop_event=tts_stop_event,
        safe_set_events=safe_set_events,
        stop_cloud_control=stop_cloud_control,
        safe_close_serial=safe_close_serial,
        stop_tray_icon=stop_tray_icon,
        flush_log_queue=flush_log_queue,
        file_log_queue=file_log_queue,
        file_log_thread=file_log_thread,
        file_log_stop_event=file_log_stop_event,
        worker_threads=worker_threads,
        deferred_worker_stop_events=deferred_worker_stop_events,
        deferred_worker_threads=deferred_worker_threads,
        deferred_worker_queues=deferred_worker_queues,
        destroy_root=destroy_root,
        log_error=log_error,
        pre_cloud_worker_threads=pre_cloud_worker_threads,
        before_stop_cloud=before_stop_cloud,
        report_progress=report_progress,
    )


def start_cleanup_and_exit_app_runtime(
    *, root, messagebox, is_exiting, set_exiting, destroy_root=None,
    cleanup_runtime=cleanup_and_exit_runtime, start_worker=start_daemon_thread,
    progress_factory=ShutdownProgress, log_error=None, notify_log_errors=None, **cleanup_options,
):
    """Keep Tk on its owner thread while the ordered cleanup waits for I/O."""
    def run_task(report_progress):
        return cleanup_runtime(
            is_exiting=False, confirm_exit=lambda: True, set_exiting=lambda _value: None,
            destroy_root=lambda: None, log_error=log_error,
            report_progress=report_progress, **cleanup_options,
        )

    return start_shutdown_task_runtime(
        root=root, messagebox=messagebox, is_exiting=is_exiting, set_exiting=set_exiting,
        run_task=run_task, on_success=destroy_root or root.destroy,
        confirmation="确定要完全退出软件吗？\n\n退出后将停止监听短信和来电。",
        start_worker=start_worker, progress_factory=progress_factory, log_error=log_error,
        notify_log_errors=notify_log_errors,
    )


def start_shutdown_task_runtime(
    *, root, messagebox, is_exiting, set_exiting, run_task, on_success,
    confirmation, action="退出", start_worker=start_daemon_thread,
    progress_factory=ShutdownProgress, log_error=None, failure_message=None,
    notify_log_errors=None,
):
    """Run ordered stop work on a worker and deliver its result through Tk's queue."""
    if is_exiting:
        return "already_exiting"
    if not messagebox.askyesno(
        action + "软件", confirmation, parent=root,
    ):
        return "cancelled"

    try:
        progress = progress_factory(root)
    except Exception as exc:
        _safe_log(log_error, f"Create shutdown progress failed: {exc!r}")
        messagebox.showerror(action + "未开始", f"无法显示{action}进度，请稍后重试。", parent=root)
        return "start_failed"
    set_exiting(True)
    task_queue = queue.Queue()
    finished = False

    def finish(result):
        nonlocal finished
        if notify_log_errors is not None:
            notify_log_errors(force=True)
        finished = True
        progress.close()
        if getattr(result, "status", result) == "exited":
            on_success()
        else:
            set_exiting(False)
            messagebox.showerror(
                action + "未完成",
                failure_message(result) if failure_message else
                f"后台清理未能完成，窗口已保留。请再次选择{action}，并检查日志中的错误。",
                parent=root,
            )

    def pump():
        ui_pump_runtime(task_queue, root, lambda: not finished, pump, log_error=log_error)

    def worker():
        try:
            result = run_task(lambda message: ui_post_runtime(task_queue, progress.update, (message,)))
        except Exception as exc:
            result = "worker_wait_failed"
            _safe_log(log_error, f"Background shutdown failed: {exc!r}")
        ui_post_runtime(task_queue, finish, (result,))

    initial_pump = None
    try:
        initial_pump = root.after(0, pump)
        start_worker("app_shutdown", worker, log_error=log_error)
    except Exception as exc:
        finished = True
        if initial_pump is not None:
            root.after_cancel(initial_pump)
        progress.close()
        set_exiting(False)
        _safe_log(log_error, f"Start background shutdown failed: {exc!r}")
        messagebox.showerror(action + "未开始", f"无法启动{action}任务，请稍后重试。", parent=root)
        return "start_failed"
    return "started"
