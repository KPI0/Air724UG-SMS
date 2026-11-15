from contextlib import nullcontext
import time


SERIAL_READ_CHUNK_BYTES = 4096


def read_serial_line_safely_runtime(
    serial_lock,
    get_serial,
    exception_cls,
    *,
    read_lock=None,
    read_context_guard=None,
    capture_read_chunk=None,
    wait_for_data=None,
    monotonic=time.monotonic,
):
    return _read_serial_line(
        serial_lock, get_serial, exception_cls, read_lock=read_lock,
        read_context_guard=read_context_guard, capture_read_chunk=capture_read_chunk,
        wait_for_data=wait_for_data, monotonic=monotonic,
    )


def _read_serial_line(
    serial_lock, get_serial, exception_cls, *, read_lock, read_context_guard,
    capture_read_chunk, wait_for_data, monotonic,
):
    read_guard = read_lock if read_lock is not None else (
        nullcontext() if capture_read_chunk is not None else serial_lock)
    serial_guard = serial_lock if read_guard is not serial_lock else nullcontext()
    try:
        with read_guard:
            with serial_guard:
                serial_obj = get_serial()
                if serial_obj is None or not serial_obj.is_open:
                    raise exception_cls("serial_obj is None (closed)")
            if capture_read_chunk is not None:
                return _read_contextual_chunk(
                    serial_obj, serial_guard, get_serial, read_context_guard or nullcontext,
                    capture_read_chunk, wait_for_data or time.sleep, monotonic,
                )
            # Keep a separate read lease until this bounded read returns. The normal
            # serial lock remains available for writes and immediate reply
            # processing, while close/reconnect waits for this OS read to finish.
            available = serial_obj.in_waiting
            size = min(SERIAL_READ_CHUNK_BYTES, max(1, available))
            return serial_obj.read(size)
    except Exception as exc:
        raise exception_cls(f"并发读取被中断: {exc}")


def _read_contextual_chunk(
    serial_obj, serial_guard, get_serial, context_guard, capture_chunk, wait_for_data, monotonic,
):
    timeout = getattr(serial_obj, "timeout", 0.3)
    deadline = monotonic() + max(0.0, float(0.3 if timeout is None else timeout))
    stopped = False
    while True:
        with serial_guard, context_guard():
            if get_serial() is not serial_obj or not serial_obj.is_open:
                raise RuntimeError("串口连接已变化，已取消旧读取")
            if stopped:
                return capture_chunk(b"")
            available = serial_obj.in_waiting
            if available > 0:
                # The read lease excludes competing readers. Take only bytes
                # already buffered, and bind their context before either a
                # sender or the UI can advance the call lifetime.
                raw = serial_obj.read(min(SERIAL_READ_CHUNK_BYTES, available))
                return capture_chunk(raw)
            remaining = deadline - monotonic()
            if remaining <= 0:
                return capture_chunk(b"")
        # Do not hold serial/call locks across an idle OS read: a new command
        # must be able to write and receive its reply during this same budget.
        stopped = bool(wait_for_data(min(0.01, remaining)))


def send_call_hangup_runtime(serial_lock, get_serial, write_command_result):
    with serial_lock:
        serial_obj = get_serial()
    return write_command_result(serial_obj, "ATH")


def safe_close_serial_runtime(
    serial_lock,
    get_serial,
    set_serial,
    unlock_port_mutex,
    *,
    read_lock=None,
    before_close=None,
):
    read_guard = read_lock if read_lock is not None else serial_lock
    serial_guard = serial_lock if read_guard is not serial_lock else nullcontext()
    with read_guard:
        with serial_guard:
            serial_obj = get_serial()
            if before_close is not None and before_close(serial_obj) is False:
                return False
            try:
                if serial_obj is not None:
                    serial_obj.close()
            except Exception:
                pass
            set_serial(None)
            unlock_port_mutex()
