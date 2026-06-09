import threading
import traceback


def format_thread_exception(name, exc):
    title = f"后台线程异常退出 [{name}]: {exc or exc.__class__.__name__}"
    return title + "\n" + traceback.format_exc()


def start_daemon_thread(
    name,
    target,
    args=(),
    kwargs=None,
    *,
    log_error=None,
    before_start=None,
    thread_factory=threading.Thread,
):
    def guarded_target():
        try:
            return target(*tuple(args), **dict(kwargs or {}))
        except Exception as exc:
            message = format_thread_exception(name, exc)
            if log_error is not None:
                try:
                    log_error(message)
                except Exception:
                    pass
            return None

    try:
        thread = thread_factory(target=guarded_target, daemon=True, name=name)
    except TypeError:
        thread = thread_factory(target=guarded_target, daemon=True)
    if before_start is not None:
        before_start(thread)
    thread.start()
    return thread
