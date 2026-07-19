import threading
import time
import traceback


class WorkerThreadRegistry:
    """Thread-safe registry for short-lived workers that must join on exit."""

    def __init__(self, lock_factory=threading.Lock):
        self._lock = lock_factory()
        self._threads = set()

    def register(self, thread):
        if thread is None:
            return False
        with self._lock:
            self._threads.add(thread)
        return True

    def unregister(self, thread):
        if thread is None:
            return False
        with self._lock:
            existed = thread in self._threads
            self._threads.discard(thread)
        return existed

    def snapshot(self):
        with self._lock:
            return tuple(self._threads)


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


def wait_for_worker_threads(
    threads,
    timeout=None,
    *,
    log_error=None,
    monotonic=time.monotonic,
    current_thread=threading.current_thread,
):
    """Wait for producer threads before the file logger is stopped.

    Production shutdown uses the default ``None`` timeout so registered
    serial and other producer threads cannot cross the final-log boundary.
    A finite timeout remains available for diagnostics and tests; callers must
    treat a ``False`` result as a hard stop and must not continue shutdown.
    """
    deadline = None
    if timeout is not None:
        try:
            deadline = monotonic() + max(0.0, float(timeout))
        except Exception:
            deadline = monotonic() + 20.0

    current = current_thread()
    all_stopped = True
    seen = set()
    for thread in tuple(threads or ()):
        if thread is None or thread is current:
            continue
        marker = id(thread)
        if marker in seen:
            continue
        seen.add(marker)
        try:
            if deadline is None:
                thread.join()
            else:
                remaining = max(0.0, deadline - monotonic())
                thread.join(timeout=remaining)
            if thread.is_alive():
                all_stopped = False
                if log_error is not None:
                    log_error(f"后台线程未在退出时限内停止: {getattr(thread, 'name', 'unknown')}")
        except Exception as exc:
            all_stopped = False
            if log_error is not None:
                try:
                    log_error(f"等待后台线程退出失败: {exc!r}")
                except Exception:
                    pass
    return all_stopped
