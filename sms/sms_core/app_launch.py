import base64
import json
import os
import subprocess
import sys
import time
from dataclasses import dataclass

from sms_core.file_log_runtime import wait_for_file_log_worker
from sms_core.threading_runtime import queues_are_drained, wait_for_worker_threads


@dataclass(frozen=True)
class RestartRuntimeResult:
    status: str
    error: object = None


def _threads_added_after_snapshot(previous_threads, current_threads):
    previous_ids = {
        id(thread)
        for thread in tuple(previous_threads or ())
        if thread is not None
    }
    return tuple(
        thread
        for thread in tuple(current_threads or ())
        if thread is not None and id(thread) not in previous_ids
    )


def get_launch_target_and_args(frozen=None, executable=None, argv0=None):
    """
    Return (target_path, script_argument, working_dir) for launching the app.

    Frozen builds launch the executable directly. Script mode launches through
    pythonw.exe when available and passes the script path as the first argument.
    """
    if frozen is None:
        frozen = getattr(sys, "frozen", False)
    if executable is None:
        executable = sys.executable
    if argv0 is None:
        argv0 = sys.argv[0]

    if frozen:
        exe_path = os.path.abspath(executable)
        return exe_path, "", os.path.dirname(exe_path)

    pyw = os.path.join(os.path.dirname(executable), "pythonw.exe")
    if not os.path.exists(pyw):
        pyw = executable

    script_path = os.path.abspath(argv0)
    return pyw, script_path, os.path.dirname(script_path)


def get_clean_restart_env(environ=None):
    # PyInstaller 6.9+ needs the onefile runtime environment reset on restart.
    clean_env = dict(os.environ if environ is None else environ)
    clean_env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    for key in ("_MEIPASS2", "_MEIPASS", "PYINSTALLER_TEMP", "TCL_LIBRARY", "TK_LIBRARY"):
        clean_env.pop(key, None)
    return clean_env


def get_detached_creationflags():
    flags = 0
    for name in ("CREATE_NO_WINDOW", "DETACHED_PROCESS", "CREATE_NEW_PROCESS_GROUP"):
        flags |= getattr(subprocess, name, 0)
    return flags


def launch_detached_process(command, env=None, cwd=None):
    kwargs = {
        "env": env,
        "cwd": cwd,
        "close_fds": True,
    }
    creationflags = get_detached_creationflags()
    if creationflags:
        kwargs["creationflags"] = creationflags
    return subprocess.Popen(command, **kwargs)


def cancel_launched_process(process, timeout=2.0):
    if process is None:
        return False
    poll = getattr(process, "poll", None)
    if callable(poll) and poll() is not None:
        return True

    terminate = getattr(process, "terminate", None)
    if not callable(terminate):
        return False
    terminate()

    wait = getattr(process, "wait", None)
    if not callable(wait):
        return True
    try:
        wait(timeout=max(0.0, float(timeout)))
        return True
    except subprocess.TimeoutExpired:
        kill = getattr(process, "kill", None)
        if not callable(kill):
            return False
        kill()
        wait(timeout=max(0.0, float(timeout)))
        return True


def encode_restart_args(args):
    payload = json.dumps(list(args), ensure_ascii=False).encode("utf-8")
    return base64.urlsafe_b64encode(payload).decode("ascii")


def decode_restart_args(payload: str):
    if not payload:
        return []
    raw = base64.urlsafe_b64decode(payload.encode("ascii"))
    data = json.loads(raw.decode("utf-8"))
    return data if isinstance(data, list) else []


def filtered_restart_args(args, excluded_flags):
    excluded = set(excluded_flags or ())
    return [arg for arg in list(args or []) if arg not in excluded]


def build_restart_helper_command(
    target,
    script_arg,
    restart_helper_flag,
    current_pid,
    restart_args,
):
    command = [target]
    if script_arg:
        command.append(script_arg)
    command.extend([
        restart_helper_flag,
        str(current_pid),
        encode_restart_args(restart_args),
    ])
    return command


def prepare_restart_helper_launch(
    argv,
    autostart_flag,
    restart_helper_flag,
    current_pid,
    launch_target_func=get_launch_target_and_args,
):
    target, script_arg, workdir = launch_target_func()
    restart_args = filtered_restart_args(
        list(argv or [])[1:],
        (autostart_flag, restart_helper_flag),
    )
    return (
        build_restart_helper_command(
            target,
            script_arg,
            restart_helper_flag,
            current_pid,
            restart_args,
        ),
        workdir,
    )


def restart_software_runtime(
    *,
    is_exiting,
    confirm_restart,
    argv,
    autostart_flag,
    restart_helper_flag,
    current_pid,
    log_error,
    show_launch_error,
    set_exiting,
    system_ui,
    stop_tray_icon,
    set_serial_running,
    safe_set_events,
    stop_events,
    stop_cloud_control,
    safe_close_serial,
    app_mutex,
    release_mutex,
    flush_log_queue,
    file_log_queue,
    exit_process,
    prepare_launch=prepare_restart_helper_launch,
    launch_process=launch_detached_process,
    cancel_launch=cancel_launched_process,
    clean_env=get_clean_restart_env,
    file_log_thread=None,
    file_log_stop_event=None,
    worker_threads=(),
    deferred_stop_events=(),
    deferred_worker_threads=(),
    deferred_worker_queues=(),
    wait_worker_threads=wait_for_worker_threads,
    wait_file_log_worker=wait_for_file_log_worker,
):
    if is_exiting:
        return RestartRuntimeResult("already_exiting")

    if not confirm_restart():
        return RestartRuntimeResult("cancelled")

    helper_process = None
    restart_committed = False
    try:
        helper_cmd, workdir = prepare_launch(
            argv,
            autostart_flag,
            restart_helper_flag,
            current_pid,
        )
        helper_process = launch_process(
            helper_cmd,
            env=clean_env(),
            cwd=workdir,
        )
    except Exception as e:
        message = f"重启尝试失败：{e}"
        log_error(message)
        try:
            show_launch_error(e)
        except Exception:
            system_ui(message, "normal")
        return RestartRuntimeResult("launch_failed", e)

    try:
        set_exiting(True)
        system_ui("🔄 正在重启软件...", "normal")
        stop_tray_icon(wait_after=0.45)
        set_serial_running(False)
        safe_set_events(*tuple(stop_events or ()))

        try:
            stop_cloud_control(update_status=False)
        except Exception:
            pass
        safe_close_serial()

        try:
            threads_to_wait = worker_threads() if callable(worker_threads) else worker_threads
        except Exception as exc:
            try:
                log_error(f"Snapshot restart worker threads failed: {exc!r}")
            except Exception:
                pass
            return RestartRuntimeResult("worker_wait_failed", exc)
        try:
            workers_stopped = wait_worker_threads(threads_to_wait, log_error=log_error)
        except Exception as exc:
            try:
                log_error(f"Wait for producer worker threads during restart raised: {exc!r}")
            except Exception:
                pass
            return RestartRuntimeResult("worker_wait_failed", exc)
        if workers_stopped is False:
            try:
                log_error(
                    "Producer worker threads are still running; file logger stop, final flush, mutex release, and restart exit were aborted"
                )
            except Exception:
                pass
            return RestartRuntimeResult("worker_wait_failed")
        if callable(worker_threads):
            try:
                final_threads_to_wait = worker_threads()
            except Exception as exc:
                try:
                    log_error(f"Snapshot final restart worker threads failed: {exc!r}")
                except Exception:
                    pass
                return RestartRuntimeResult("worker_wait_failed", exc)
            late_threads = _threads_added_after_snapshot(
                threads_to_wait,
                final_threads_to_wait,
            )
            if late_threads:
                try:
                    late_workers_stopped = wait_worker_threads(
                        late_threads,
                        log_error=log_error,
                    )
                except Exception as exc:
                    try:
                        log_error(f"Wait for late producer worker threads during restart raised: {exc!r}")
                    except Exception:
                        pass
                    return RestartRuntimeResult("worker_wait_failed", exc)
                if late_workers_stopped is False:
                    try:
                        log_error("Late producer worker threads are still running; restart cleanup was aborted")
                    except Exception:
                        pass
                    return RestartRuntimeResult("worker_wait_failed")
        deferred_events = tuple(deferred_stop_events or ())
        has_deferred_workers = bool(
            deferred_events
            or deferred_worker_queues
            or callable(deferred_worker_threads)
            or deferred_worker_threads
        )
        if has_deferred_workers:
            safe_set_events(*deferred_events)
            try:
                deferred_threads_to_wait = (
                    deferred_worker_threads() if callable(deferred_worker_threads) else deferred_worker_threads
                )
            except Exception as exc:
                try:
                    log_error(f"Snapshot deferred restart worker threads failed: {exc!r}")
                except Exception:
                    pass
                return RestartRuntimeResult("worker_wait_failed", exc)
            try:
                deferred_workers_stopped = wait_worker_threads(
                    deferred_threads_to_wait,
                    log_error=log_error,
                )
            except Exception as exc:
                try:
                    log_error(f"Wait for deferred worker threads during restart raised: {exc!r}")
                except Exception:
                    pass
                return RestartRuntimeResult("worker_wait_failed", exc)
            if deferred_workers_stopped is False:
                try:
                    log_error("Deferred worker threads are still running; restart cleanup was aborted")
                except Exception:
                    pass
                return RestartRuntimeResult("worker_wait_failed")
            if not queues_are_drained(deferred_worker_queues, log_error=log_error):
                try:
                    log_error("Deferred worker queues were not fully drained; restart cleanup was aborted")
                except Exception:
                    pass
                return RestartRuntimeResult("worker_wait_failed")
        if file_log_stop_event is not None:
            safe_set_events(file_log_stop_event)

        try:
            file_log_stopped = wait_file_log_worker(file_log_thread, log_error=log_error)
        except Exception as exc:
            try:
                log_error(f"Wait for file log worker during restart raised: {exc!r}")
            except Exception:
                pass
            return RestartRuntimeResult("file_log_wait_failed", exc)
        if file_log_stopped is False:
            try:
                log_error(
                    "File log worker is still running; final flush, mutex release, and restart exit were aborted"
                )
            except Exception:
                pass
            return RestartRuntimeResult("file_log_wait_failed")

        flush_log_queue(file_log_queue)
        try:
            if app_mutex:
                release_mutex(app_mutex)
        except Exception:
            pass
        restart_committed = True
        exit_process(0)
        return RestartRuntimeResult("exited")
    finally:
        if helper_process is not None and not restart_committed:
            try:
                cancelled = cancel_launch(helper_process)
                if cancelled is False:
                    log_error("Cancel pending restart helper returned False")
            except Exception as exc:
                try:
                    log_error(f"Cancel pending restart helper failed: {exc!r}")
                except Exception:
                    pass


def show_early_error(title: str, message: str):
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, str(message), str(title), 0x10)
    except Exception:
        pass


def wait_for_process_exit(pid: int, sleep_func=time.sleep):
    try:
        import ctypes

        target_pid = int(pid)
        if target_pid <= 0:
            return

        synchronize = 0x00100000
        wait_object_0 = 0x00000000
        wait_timeout = 0x00000102

        handle = ctypes.windll.kernel32.OpenProcess(synchronize, False, target_pid)
        if handle:
            try:
                while True:
                    result = ctypes.windll.kernel32.WaitForSingleObject(handle, 200)
                    if result == wait_object_0:
                        break
                    if result != wait_timeout:
                        break
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
        else:
            sleep_func(2.0)
    except Exception:
        sleep_func(2.0)

    sleep_func(0.3)


def maybe_run_restart_helper_mode(
    restart_helper_flag,
    argv=None,
    wait_func=wait_for_process_exit,
    launch_func=launch_detached_process,
    error_func=show_early_error,
):
    argv = list(sys.argv if argv is None else argv)
    if restart_helper_flag not in argv:
        return False

    try:
        idx = argv.index(restart_helper_flag)
        wait_pid = int(argv[idx + 1])
        payload = argv[idx + 2] if len(argv) > idx + 2 else ""
        restart_args = decode_restart_args(payload)

        target, script_arg, workdir = get_launch_target_and_args(argv0=argv[0])
        launch_cmd = [target]
        if script_arg:
            launch_cmd.append(script_arg)
        launch_cmd.extend(arg for arg in restart_args if arg != restart_helper_flag)

        wait_func(wait_pid)
        launch_func(
            launch_cmd,
            env=get_clean_restart_env(),
            cwd=workdir,
        )
        raise SystemExit(0)
    except SystemExit:
        raise
    except Exception as e:
        error_func("重启失败", f"软件重启辅助进程启动失败：\n\n{e}")
        raise SystemExit(1) from e
