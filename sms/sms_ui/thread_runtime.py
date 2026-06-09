import queue
import threading


def run_on_ui_thread(callback, ui_post, *, current_thread=threading.current_thread, main_thread=threading.main_thread):
    if current_thread() is main_thread():
        return callback()
    ui_post(callback)
    return None


def tk_alive_runtime(root, shutdown_event, *, current_thread=threading.current_thread, main_thread=threading.main_thread):
    if shutdown_event.is_set():
        return False
    if root is None:
        return False
    if current_thread() is main_thread():
        try:
            return bool(root.winfo_exists())
        except Exception:
            return False
    return True


def ui_post_runtime(task_queue, callback, args=(), kwargs=None, *, on_full=None):
    try:
        task_queue.put_nowait((callback, tuple(args), dict(kwargs or {})))
    except queue.Full:
        if on_full is not None:
            try:
                on_full()
            except Exception:
                pass
    except Exception:
        pass


def ui_pump_runtime(task_queue, root, tk_alive, schedule_self, *, max_batch=200, interval_ms=30):
    processed = 0
    while processed < max_batch:
        try:
            callback, args, kwargs = task_queue.get_nowait()
        except queue.Empty:
            break
        try:
            callback(*args, **kwargs)
        except Exception:
            pass
        processed += 1

    if tk_alive():
        try:
            root.after(interval_ms, schedule_self)
        except Exception:
            pass
    return processed


def schedule_delayed_ui_runtime(
    callback,
    *,
    app_start_mono,
    start_ui_delay,
    monotonic,
    root_after,
    run_on_ui_thread,
    ui_post,
):
    def schedule_in_main():
        try:
            elapsed = monotonic() - app_start_mono
            delay_ms = int(max(0.0, start_ui_delay - elapsed) * 1000)
        except Exception:
            delay_ms = 0

        try:
            if delay_ms > 0:
                return root_after(delay_ms, callback)
            return callback()
        except Exception:
            return callback()

    return run_on_ui_thread(schedule_in_main, ui_post)


def ui_messagebox_runtime(kind, title, message, *, messagebox, run_on_ui_thread, ui_post):
    def show_on_ui():
        if kind == "info":
            return messagebox.showinfo(title, message)
        if kind == "warning":
            return messagebox.showwarning(title, message)
        if kind == "error":
            return messagebox.showerror(title, message)
        if kind == "askyesno":
            return messagebox.askyesno(title, message)
        return None

    return run_on_ui_thread(show_on_ui, ui_post)
