import inspect
import queue
import sys
import threading
import time

from sms_core.threading_runtime import start_daemon_thread, task_done_safely


class FileLogErrorState:
    """A bounded, independent mailbox consumed by the UI, including at exit."""

    def __init__(self, *, monotonic=time.monotonic, interval=60.0):
        self._pending = queue.Queue(maxsize=1)
        self._monotonic = monotonic
        self._interval = interval
        self._last_notice = None

    def report(self, _detail):
        # Never retain paths, log contents or remote error text in the notice.
        try:
            self._pending.put_nowait(True)
        except queue.Full:
            pass

    def take_notice(self, *, force=False):
        now = self._monotonic()
        if not force and self._last_notice is not None and now - self._last_notice < self._interval:
            return None
        try:
            self._pending.get_nowait()
        except queue.Empty:
            return None
        self._pending.task_done()
        self._last_notice = now
        return "⚠️ 日志保存失败，部分记录可能未写入文件。请检查磁盘空间、日志目录权限和文件占用；失败记录不会自动补写。"


class FileLogQueue(queue.Queue):
    """Retain Queue semantics while reporting rejected log entries independently."""

    def __init__(self, maxsize, *, log_error):
        super().__init__(maxsize)
        self._log_error = log_error

    def put(self, item, block=True, timeout=None):
        try:
            return super().put(item, block=block, timeout=timeout)
        except queue.Full:
            self._log_error("File log queue is full")
            raise


def drain_available_log_lines(log_queue, first_item):
    path, line = first_item
    batches = {path: [line]}
    while True:
        try:
            next_item = log_queue.get_nowait()
        except queue.Empty:
            break
        # Mark the item complete before unpacking it so malformed entries do
        # not leave Queue.unfinished_tasks permanently elevated.
        task_done_safely(log_queue)
        next_path, next_line = next_item
        batches.setdefault(next_path, []).append(next_line)
    return batches


def write_log_batches(batches, encoding="utf-8", open_file=open, on_error=None):
    written = 0
    for path, lines in batches.items():
        try:
            with open_file(path, "a", encoding=encoding) as file:
                file.writelines(lines)
            written += len(lines)
        except Exception as exc:
            # A single failing path (disk full, permission denied, bad path)
            # must not silently drop logs nor abort the remaining batches.
            # Report it so the failure is diagnosable, then continue.
            if on_error is not None:
                try:
                    on_error(
                        f"file_log_worker failed to write {len(lines)} line(s) "
                        f"to {path!r}: {exc!r}"
                    )
                except Exception:
                    pass
    return written


def _default_diagnostic(message):
    try:
        sys.stderr.write(message + "\n")
        sys.stderr.flush()
    except Exception:
        pass


def _write_batches_with_error_reporting(write_batches, batches, on_error):
    # The default write_batches (write_log_batches) accepts on_error so per-path
    # write failures are reported. Injected test doubles often take only the
    # batches arg, so forward on_error only when the callable can accept it.
    try:
        supports_on_error = "on_error" in inspect.signature(write_batches).parameters
    except (TypeError, ValueError):
        supports_on_error = False
    if supports_on_error:
        return write_batches(batches, on_error=on_error)
    return write_batches(batches)


def run_file_log_worker(
    *,
    log_queue,
    stop_event,
    poll_timeout=0.5,
    queue_error_sleep=0.2,
    drain_batches=drain_available_log_lines,
    write_batches=write_log_batches,
    on_error=_default_diagnostic,
    sleep=time.sleep,
):
    while not stop_event.is_set():
        try:
            first_item = log_queue.get(timeout=poll_timeout)
        except queue.Empty:
            continue
        except Exception as exc:
            # Unexpected queue failure: report it, avoid a hot spin, keep alive.
            if on_error is not None:
                try:
                    on_error(f"file_log_worker queue read failed: {exc!r}")
                except Exception:
                    pass
            try:
                sleep(queue_error_sleep)
            except Exception:
                pass
            continue

        try:
            try:
                batches = drain_batches(log_queue, first_item)
                _write_batches_with_error_reporting(write_batches, batches, on_error)
            except Exception as exc:
                # A malformed queue item (e.g. not a (path, line) tuple) must
                # not kill the only thread that flushes logs to disk. Report
                # and skip.
                if on_error is not None:
                    try:
                        on_error(f"file_log_worker skipped a bad log item: {exc!r}")
                    except Exception:
                        pass
        finally:
            # ``first_item`` was removed by the blocking get above.  The
            # draining helper balances every additional item it removes.
            task_done_safely(log_queue)


def start_file_log_worker(
    *,
    log_queue,
    stop_event,
    thread_factory=threading.Thread,
    log_error=None,
):
    # The file-log worker is the only thread that flushes logs to disk, so its
    # own failures cannot be reported through the file log (circular). Fall back
    # to stderr so a dead worker is still diagnosable.
    guard_log = log_error if log_error is not None else _default_diagnostic
    return start_daemon_thread(
        "file_log_worker",
        lambda: run_file_log_worker(
            log_queue=log_queue,
            stop_event=stop_event,
            on_error=guard_log,
        ),
        log_error=guard_log,
        thread_factory=thread_factory,
    )


def wait_for_file_log_worker(thread, timeout=None, *, log_error=None):
    """Wait for the file logger before final queue flush and process exit.

    Production shutdown uses the default ``None`` timeout so an in-flight
    batch can never be abandoned.  A finite timeout remains available for
    diagnostics and tests, but callers must treat ``False`` as a hard stop.
    """
    if thread is None:
        return True
    try:
        if timeout is None:
            thread.join()
        else:
            thread.join(timeout=max(0.0, float(timeout)))
        if thread.is_alive():
            if log_error is not None:
                try:
                    log_error("file_log_worker did not stop before shutdown timeout")
                except Exception:
                    pass
            return False
        return True
    except Exception as exc:
        if log_error is not None:
            try:
                log_error(f"Wait for file_log_worker failed: {exc!r}")
            except Exception:
                pass
        return False
