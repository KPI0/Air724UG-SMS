import threading
import time
import traceback


def task_done_safely(work_queue):
    """Balance one successful queue ``get`` without trusting the queue type.

    The production queues are :class:`queue.Queue` instances, but a few
    headless/test paths use light-weight queue doubles.  Consumers should call
    this exactly once for every item they remove; a missing or already-closed
    ``task_done`` implementation must not take down the worker or UI pump.
    """
    task_done = getattr(work_queue, "task_done", None)
    if not callable(task_done):
        return False
    try:
        task_done()
        return True
    except Exception:
        return False


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


class SingleFlightTaskState:
    """Thread-safe state for tasks that allow only one active invocation."""

    def __init__(self, lock_factory=threading.Lock):
        self._lock = lock_factory()
        self._active = False

    def try_acquire(self):
        with self._lock:
            if self._active:
                return False
            self._active = True
            return True

    def release(self):
        with self._lock:
            was_active = self._active
            self._active = False
            return was_active

    def is_active(self):
        with self._lock:
            return self._active


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


def start_registered_daemon_thread(
    name,
    target,
    *,
    thread_registry,
    log_error=None,
    thread_factory=threading.Thread,
):
    """Start a daemon worker whose full lifetime is visible to shutdown."""
    thread_holder = {}

    def register_thread(thread):
        thread_holder["thread"] = thread
        if thread_registry is not None:
            thread_registry.register(thread)

    def run_registered():
        try:
            return target()
        finally:
            if thread_registry is not None:
                thread_registry.unregister(thread_holder.get("thread"))

    try:
        return start_daemon_thread(
            name,
            run_registered,
            log_error=log_error,
            before_start=register_thread,
            thread_factory=thread_factory,
        )
    except Exception:
        if thread_registry is not None:
            thread_registry.unregister(thread_holder.get("thread"))
        raise


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


def queues_are_drained(queues, *, log_error=None):
    """Return whether queued work has no unfinished tasks after its worker exits."""
    for work_queue in tuple(queues or ()):
        if work_queue is None:
            continue
        try:
            unfinished = int(work_queue.unfinished_tasks)
        except (AttributeError, TypeError, ValueError):
            unfinished = None
        except Exception as exc:
            if log_error is not None:
                try:
                    log_error(f"检查退出队列状态失败: {exc!r}")
                except Exception:
                    pass
            return False
        if unfinished is not None and unfinished != 0:
            if log_error is not None:
                try:
                    log_error(f"退出时仍有 {unfinished} 个队列任务未完成")
                except Exception:
                    pass
            return False
        try:
            empty = bool(work_queue.empty())
        except Exception as exc:
            if log_error is not None:
                try:
                    log_error(f"检查退出队列是否为空失败: {exc!r}")
                except Exception:
                    pass
            return False
        if not empty:
            if log_error is not None:
                try:
                    log_error("退出时队列仍有待处理任务")
                except Exception:
                    pass
            return False
    return True
