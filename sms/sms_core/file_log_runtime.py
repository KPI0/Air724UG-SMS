import inspect
import queue
import sys
import threading
import time

from sms_core.threading_runtime import start_daemon_thread


def drain_available_log_lines(log_queue, first_item):
    path, line = first_item
    batches = {path: [line]}
    while True:
        try:
            next_path, next_line = log_queue.get_nowait()
        except queue.Empty:
            break
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
            batches = drain_batches(log_queue, first_item)
            _write_batches_with_error_reporting(write_batches, batches, on_error)
        except Exception as exc:
            # A malformed queue item (e.g. not a (path, line) tuple) must not
            # kill the only thread that flushes logs to disk. Report and skip.
            if on_error is not None:
                try:
                    on_error(f"file_log_worker skipped a bad log item: {exc!r}")
                except Exception:
                    pass


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
