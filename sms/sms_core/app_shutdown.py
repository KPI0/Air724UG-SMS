def safe_set_events(*events):
    count = 0
    for event in events:
        if event is None:
            continue
        try:
            event.set()
            count += 1
        except Exception:
            pass
    return count


def drain_log_queue(log_queue):
    batches = {}
    while True:
        try:
            path, line = log_queue.get_nowait()
        except Exception:
            break
        batches.setdefault(path, []).append(line)
    return batches


def flush_log_queue(log_queue, encoding="utf-8"):
    batches = drain_log_queue(log_queue)
    written = 0
    for path, lines in batches.items():
        try:
            with open(path, "a", encoding=encoding) as file:
                file.writelines(lines)
            written += len(lines)
        except Exception:
            pass
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
):
    if is_exiting:
        return "already_exiting"

    if not confirm_exit():
        return "cancelled"

    set_exiting(True)
    safe_set_events(*tuple(shutdown_events or ()))
    set_serial_running(False)
    safe_set_events(*tuple(worker_stop_events or ()))

    try:
        stop_cloud_control(update_status=False)
    except Exception:
        pass

    safe_close_serial()
    stop_tray_icon(wait_after=0.25)
    flush_log_queue(file_log_queue)

    try:
        destroy_root()
    except Exception:
        pass

    safe_set_events(tts_stop_event)
    return "exited"
