import threading
import time


IMEI_READ_COMMAND = "AT+CGSN"
IMEI_QUERY_WINDOW_SECONDS = 6.0


def notify_cloud_identity_changed_runtime(
    *,
    get_loop,
    get_ws,
    is_connected,
    runtime_imei,
    send_register,
    run_coroutine_threadsafe,
):
    try:
        loop = get_loop()
        ws = get_ws()
        if loop is None or not loop.is_running() or ws is None or not is_connected():
            return False
        if not runtime_imei():
            return False
        run_coroutine_threadsafe(send_register(ws), loop)
        return True
    except Exception:
        return False


def set_cloud_device_imei_runtime(
    imei,
    *,
    current_imei,
    normalize_imei,
    set_device_imei,
    set_verified,
    log,
    notify_identity_changed,
    source="",
):
    normalized = normalize_imei(imei)
    if normalized and not (14 <= len(normalized) <= 17):
        return False

    if normalized == normalize_imei(current_imei()):
        if normalized:
            set_verified(True)
        return True

    set_device_imei(normalized)
    set_verified(bool(normalized))

    if normalized:
        log(f"设备IMEI已更新：{normalized}")
        notify_identity_changed()

    return True


def maybe_capture_cloud_device_imei_runtime(
    line,
    *,
    query_deadline,
    set_query_deadline,
    imei_regex,
    set_device_imei,
    monotonic=time.monotonic,
):
    if query_deadline <= 0:
        return "inactive"
    if monotonic() > query_deadline:
        set_query_deadline(0.0)
        return "expired"

    text = str(line or "").strip()
    upper = text.upper()
    if not text or upper in ("OK", "ERROR") or "AT+CGSN" in upper:
        return "ignored"

    match = imei_regex.search(text)
    if not match:
        return "no_match"

    imei = match.group(1)
    if set_device_imei(imei, source=IMEI_READ_COMMAND):
        set_query_deadline(0.0)
        return "captured"
    return "unchanged"


def request_cloud_device_imei_worker(
    *,
    serial_lock,
    get_serial,
    write_command_result,
    set_query_deadline,
    monotonic=time.monotonic,
    push_serial_debug=None,
    cloud_log,
):
    try:
        with serial_lock:
            result = write_command_result(get_serial(), IMEI_READ_COMMAND)
            if not result.ok:
                cloud_log(f"读取IMEI失败：{result.error}")
                return False
            set_query_deadline(monotonic() + IMEI_QUERY_WINDOW_SECONDS)

        if push_serial_debug is not None:
            try:
                push_serial_debug(">>> 云端控制读取IMEI: AT+CGSN\\r\\n")
            except Exception:
                pass
        cloud_log("已发送读取IMEI指令：AT+CGSN")
        return True
    except Exception as exc:
        cloud_log(f"读取IMEI失败：{exc}")
        return False


def request_cloud_device_imei_runtime(
    *,
    serial_lock,
    get_serial,
    write_command_result,
    set_query_deadline,
    cloud_log,
    monotonic=time.monotonic,
    push_serial_debug=None,
    thread_factory=threading.Thread,
):
    def task():
        request_cloud_device_imei_worker(
            serial_lock=serial_lock,
            get_serial=get_serial,
            write_command_result=write_command_result,
            set_query_deadline=set_query_deadline,
            monotonic=monotonic,
            push_serial_debug=push_serial_debug,
            cloud_log=cloud_log,
        )

    try:
        thread_factory(target=task, daemon=True).start()
        return True, "已尝试发送读取IMEI指令"
    except Exception as exc:
        try:
            cloud_log(f"读取IMEI线程启动失败：{exc}")
        except Exception:
            pass
        return False, f"读取IMEI失败：{exc}"
