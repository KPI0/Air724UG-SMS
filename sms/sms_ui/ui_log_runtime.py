from datetime import datetime, timedelta
import os
import queue


def insert_main_text_runtime(text_widget, msg, tag="normal", *, max_lines=3000, end_marker=None):
    if text_widget is None:
        return False
    try:
        if not text_widget.winfo_exists():
            return False
    except Exception:
        return False

    end = end_marker or "end"
    try:
        is_at_bottom = text_widget.yview()[1] >= 0.98
    except Exception:
        is_at_bottom = True

    text_widget.insert(end, msg + "\n", tag)

    try:
        total_lines = int(text_widget.index("end-1c").split(".")[0])
        if total_lines > max_lines:
            text_widget.delete("1.0", f"{total_lines - max_lines + 1}.0")
    except Exception:
        pass

    if is_at_bottom:
        text_widget.see(end)
    return True


def write_port_log_runtime(msg, log_dir, log_prefix, file_log, now=None):
    now = now or datetime.now()
    today = now.strftime("%Y-%m-%d")
    path = os.path.join(log_dir, f"sms_{log_prefix}_{today}.txt")
    line = f"{now:%Y-%m-%d %H:%M:%S} {msg}\n"
    file_log((path, line))
    return path, line


def write_system_log_runtime(msg, log_dir, file_log, now=None):
    now = now or datetime.now()
    today = now.strftime("%Y-%m-%d")
    path = os.path.join(log_dir, f"sms_system_{today}.txt")
    line = f"{now:%Y-%m-%d %H:%M:%S} {msg}\n"
    file_log((path, line))
    return path, line


def log_file_only_runtime(msg, *, log_dir, file_log, now=None):
    try:
        return write_system_log_runtime(msg, log_dir, file_log, now=now)
    except Exception:
        return None


def flush_pending_ui_logs_runtime(pending_logs, insert_text):
    flushed = 0
    while True:
        try:
            message, tag = pending_logs.get_nowait()
        except queue.Empty:
            break
        except Exception:
            break
        try:
            insert_text(message, tag)
        except Exception:
            pass
        flushed += 1
    return flushed


def show_sms_popup_runtime(message, *, popup_enabled, show_info, show_window):
    if not popup_enabled:
        return "disabled"
    try:
        show_info("短信提醒", message)
        show_window()
        return "shown"
    except Exception:
        return "error"


def clear_text_widget_runtime(text_widget, *, start="1.0", end="end"):
    try:
        if text_widget is not None and text_widget.winfo_exists():
            text_widget.delete(start, end)
            return True
    except Exception:
        pass
    return False


def schedule_next_midnight_clear_runtime(
    *,
    tk_alive,
    schedule_after,
    clear_callback,
    now_func=datetime.now,
):
    try:
        now = now_func()
        next_midnight = now.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)
        if tk_alive():
            delay_ms = int((next_midnight - now).total_seconds() * 1000)
            schedule_after(delay_ms, clear_callback)
            return delay_ms
    except Exception:
        pass
    return None


def run_log_runtime(
    msg,
    tag,
    *,
    has_text,
    insert_text,
    log_early,
    log_dir,
    log_prefix,
    file_log,
    now=None,
):
    try:
        if has_text():
            insert_text(msg, tag)
        else:
            log_early(msg, tag)
            return "early"
    except Exception:
        try:
            log_early(msg, tag)
        except Exception:
            pass
        return "early"

    try:
        write_port_log_runtime(msg, log_dir, log_prefix, file_log, now=now)
    except Exception:
        pass
    return "logged"


def system_ui_runtime(
    message,
    tag,
    *,
    tk_alive,
    log_file_only,
    pending_logs,
    has_text,
    insert_text,
    schedule_ui,
):
    if not tk_alive():
        log_file_only(message)
        _put_pending(pending_logs, message, tag)
        return "pending"

    log_file_only(message)

    def do_ui():
        try:
            if has_text():
                insert_text(message, tag)
            else:
                _put_pending(pending_logs, message, tag)
        except Exception:
            _put_pending(pending_logs, message, tag)

    schedule_ui(do_ui)
    return "scheduled"


def ui_only_runtime(
    message,
    tag,
    *,
    pending_logs,
    has_text,
    insert_text,
    run_on_ui_thread,
    ui_post,
):
    def do_ui():
        try:
            if has_text():
                insert_text(message, tag)
                return "inserted"
            _put_pending(pending_logs, message, tag)
            return "pending"
        except Exception:
            _put_pending(pending_logs, message, tag)
            return "pending"

    return run_on_ui_thread(do_ui, ui_post)


def _put_pending(pending_logs, message, tag):
    try:
        pending_logs.put_nowait((message, tag))
    except Exception:
        pass
