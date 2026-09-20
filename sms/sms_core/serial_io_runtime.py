from contextlib import nullcontext


def read_serial_line_safely_runtime(
    serial_lock,
    get_serial,
    exception_cls,
    *,
    read_lock=None,
):
    return _read_serial_line(serial_lock, get_serial, exception_cls, read_lock=read_lock)


def _read_serial_line(serial_lock, get_serial, exception_cls, *, read_lock):
    read_guard = read_lock if read_lock is not None else serial_lock
    serial_guard = serial_lock if read_guard is not serial_lock else nullcontext()
    with read_guard:
        with serial_guard:
            serial_obj = get_serial()
            if serial_obj is None or not serial_obj.is_open:
                raise exception_cls("serial_obj is None (closed)")
        # Keep a separate read lease until readline() returns.  The normal
        # serial lock remains available for writes and immediate reply
        # processing, while close/reconnect waits for this OS read to finish.
        try:
            return serial_obj.readline()
        except Exception as exc:
            raise exception_cls(f"并发读取被中断: {exc}")


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
):
    read_guard = read_lock if read_lock is not None else serial_lock
    serial_guard = serial_lock if read_guard is not serial_lock else nullcontext()
    with read_guard:
        with serial_guard:
            serial_obj = get_serial()
            try:
                if serial_obj is not None:
                    serial_obj.close()
            except Exception:
                pass
            set_serial(None)
            unlock_port_mutex()
