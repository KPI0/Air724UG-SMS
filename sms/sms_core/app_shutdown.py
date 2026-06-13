import inspect
import queue


def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def _call_with_optional_log_error(func, *args, log_error=None, **kwargs):
    if log_error is not None:
        try:
            signature = inspect.signature(func)
            supports_log_error = (
                "log_error" in signature.parameters
                or any(param.kind == inspect.Parameter.VAR_KEYWORD for param in signature.parameters.values())
            )
        except (TypeError, ValueError):
            supports_log_error = True
        if supports_log_error:
            kwargs["log_error"] = log_error
    return func(*args, **kwargs)


def safe_set_events(*events, log_error=None):
    count = 0
    for event in events:
        if event is None:
            continue
        try:
            event.set()
            count += 1
        except Exception as exc:
            _safe_log(log_error, f"Set shutdown event failed: {exc!r}")
    return count


def drain_log_queue(log_queue, *, log_error=None):
    batches = {}
    while True:
        try:
            path, line = log_queue.get_nowait()
        except queue.Empty:
            break
        except Exception as exc:
            _safe_log(log_error, f"Drain shutdown log queue failed: {exc!r}")
            break
        batches.setdefault(path, []).append(line)
    return batches


def flush_log_queue(log_queue, encoding="utf-8", *, log_error=None):
    batches = drain_log_queue(log_queue, log_error=log_error)
    written = 0
    for path, lines in batches.items():
        try:
            with open(path, "a", encoding=encoding) as file:
                file.writelines(lines)
            written += len(lines)
        except Exception as exc:
            _safe_log(log_error, f"Flush shutdown log queue failed for {path!r}: {exc!r}")
    return written


def cleanup_and_exit_runtime(
    *,
    is_exiting,
    confirm_exit,
    set_exiting,
    set_serial_running,
    shutdown_events,
    worker_stop_events,
    tts_stop_event,
    stop_cloud_control,
    safe_close_serial,
    stop_tray_icon,
    file_log_queue,
    destroy_root,
    safe_set_events=safe_set_events,
    flush_log_queue=flush_log_queue,
    log_error=None,
):
    if is_exiting:
        return "already_exiting"

    if not confirm_exit():
        return "cancelled"

    set_exiting(True)
    _call_with_optional_log_error(safe_set_events, *tuple(shutdown_events or ()), log_error=log_error)
    set_serial_running(False)
    _call_with_optional_log_error(safe_set_events, *tuple(worker_stop_events or ()), log_error=log_error)

    try:
        stop_cloud_control(update_status=False)
    except Exception as exc:
        _safe_log(log_error, f"Stop cloud control during shutdown failed: {exc!r}")

    try:
        safe_close_serial()
    except Exception as exc:
        _safe_log(log_error, f"Close serial during shutdown failed: {exc!r}")

    try:
        stop_tray_icon(wait_after=0.25)
    except Exception as exc:
        _safe_log(log_error, f"Stop tray icon during shutdown failed: {exc!r}")

    _call_with_optional_log_error(flush_log_queue, file_log_queue, log_error=log_error)

    try:
        destroy_root()
    except Exception as exc:
        _safe_log(log_error, f"Destroy root during shutdown failed: {exc!r}")

    _call_with_optional_log_error(safe_set_events, tts_stop_event, log_error=log_error)
    return "exited"
