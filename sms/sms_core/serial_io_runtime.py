def read_serial_line_safely_runtime(serial_lock, get_serial, exception_cls):
    with serial_lock:
        serial_obj = get_serial()
        if serial_obj is None or not serial_obj.is_open:
            raise exception_cls("serial_obj is None (closed)")
        current_serial = serial_obj

    try:
        return current_serial.readline()
    except Exception as exc:
        raise exception_cls(f"并发读取被中断: {exc}")


def send_call_hangup_runtime(serial_lock, get_serial, write_command_result):
    with serial_lock:
        serial_obj = get_serial()
    return write_command_result(serial_obj, "ATH")


def safe_close_serial_runtime(serial_lock, get_serial, set_serial, unlock_port_mutex):
    with serial_lock:
        serial_obj = get_serial()
        try:
            if serial_obj is not None:
                serial_obj.close()
        except Exception:
            pass
        set_serial(None)
        unlock_port_mutex()
