from dataclasses import dataclass
import re
import threading
import time

from sms_core.serial_debug import build_serial_command_payload
from sms_core.sms_pdu import encode_text_sms_pdus
from sms_core.threading_runtime import start_daemon_thread


@dataclass(frozen=True)
class SerialCommandResult:
    ok: bool
    error: str = ""


@dataclass(frozen=True)
class SmsPduSendResponse:
    ok: bool
    error: str = ""
    line: str = ""


SERIAL_LOG_PREFIX_RE = re.compile(r"^\s*\[[IWE]\]-\[[^\]]+\]\s*(?P<body>.*?)\s*$")
SMS_PDU_SEND_DEFAULT_TIMEOUT = 45.0
DEFAULT_SERIAL_WRITE_LOCK = threading.RLock()


class SmsPduSendWaiter:
    def __init__(self, label=""):
        self.label = str(label or "")
        self._event = threading.Event()
        self._lock = threading.Lock()
        self._response = None
        self._saw_cmgs = False

    def wait(self, timeout):
        if self._event.wait(float(timeout)):
            with self._lock:
                return self._response or SmsPduSendResponse(False, "短信发送结果未知")
        return SmsPduSendResponse(False, "等待短信发送确认超时")

    def cancel(self, error="短信发送已取消"):
        self.fail(error)

    def fail(self, error, line=""):
        self._complete(SmsPduSendResponse(False, str(error or "短信发送失败"), str(line or "")))

    def observe_line(self, line):
        body = _modem_response_body(line)
        if not body:
            return

        upper = body.upper()
        if (
            "+CMS ERROR" in upper
            or "+CME ERROR" in upper
            or upper == "ERROR"
            or upper.endswith(" ERROR")
            or " FALSE ERROR" in upper
        ):
            self._complete(SmsPduSendResponse(False, body, str(line or "")))
            return

        if "+CMGS" in upper:
            self._saw_cmgs = True
            if _response_has_final_ok(upper):
                self._complete(SmsPduSendResponse(True, "", str(line or "")))
            return

        if self._saw_cmgs and upper == "OK":
            self._complete(SmsPduSendResponse(True, "", str(line or "")))

    def done(self):
        return self._event.is_set()

    def _complete(self, response):
        with self._lock:
            if self._event.is_set():
                return
            self._response = response
            self._event.set()


class SmsPduSendCoordinator:
    def __init__(self):
        self._lock = threading.Lock()
        self._active = None

    def begin_segment(self, label=""):
        waiter = SmsPduSendWaiter(label)
        with self._lock:
            if self._active is not None and not self._active.done():
                waiter.fail("已有短信分片正在等待 Modem 确认")
                return waiter
            self._active = waiter
        return waiter

    def observe_line(self, line):
        with self._lock:
            waiter = self._active
        if waiter is not None:
            waiter.observe_line(line)

    def finish(self, waiter):
        with self._lock:
            if self._active is waiter:
                self._active = None


DEFAULT_SMS_PDU_SEND_COORDINATOR = SmsPduSendCoordinator()


def _is_serial_open(serial_obj):
    return serial_obj is not None and bool(getattr(serial_obj, "is_open", False))


def _modem_response_body(line):
    text = str(line or "").strip()
    match = SERIAL_LOG_PREFIX_RE.match(text)
    return (match.group("body") if match else text).strip()


def _response_has_final_ok(upper_body):
    text = str(upper_body or "").strip()
    return text == "OK" or text.endswith(" OK") or " TRUE OK" in text


def _sms_send_error(message, push_debug=None, port_ui=None):
    error = str(message or "短信发送失败")
    if push_debug:
        push_debug(f">>> 发送失败: {error}")
    if port_ui:
        port_ui(f"❌ 发送短信失败：{error}", "normal")
    return False


def _begin_sms_segment_response(response_coordinator, label):
    if response_coordinator is None:
        return None, SerialCommandResult(False, "未配置短信发送确认器，无法确认 Modem 发送结果")
    try:
        try:
            waiter = response_coordinator.begin_segment(label=label)
        except TypeError:
            waiter = response_coordinator.begin_segment()
    except Exception as exc:
        return None, SerialCommandResult(False, str(exc) or exc.__class__.__name__)
    if waiter is None:
        return None, SerialCommandResult(False, "短信发送确认器未返回等待对象")
    return waiter, SerialCommandResult(True)


def _finish_sms_segment_response(response_coordinator, waiter):
    if response_coordinator is None or waiter is None:
        return
    try:
        response_coordinator.finish(waiter)
    except Exception:
        pass


def _cancel_sms_segment_response(waiter, error):
    if waiter is None:
        return
    try:
        waiter.cancel(error)
    except Exception:
        pass


def _wait_sms_segment_response(waiter, timeout):
    try:
        return waiter.wait(timeout)
    except Exception as exc:
        return SmsPduSendResponse(False, str(exc) or exc.__class__.__name__)


def write_serial_command_result(serial_obj, command, append_crlf=True, push_debug=None):
    with DEFAULT_SERIAL_WRITE_LOCK:
        if not _is_serial_open(serial_obj):
            error = "串口未连接"
            if push_debug:
                push_debug(f">>> 发送失败: {error}")
            return SerialCommandResult(False, error)

        command_bytes, display_suffix = build_serial_command_payload(command, append_crlf)
        try:
            serial_obj.write(command_bytes)
            serial_obj.flush()
            if push_debug:
                push_debug(f">>> 发送: {command}{display_suffix}")
            return SerialCommandResult(True)
        except Exception as exc:
            error = str(exc) or exc.__class__.__name__
            if push_debug:
                push_debug(f">>> 发送失败: {error}")
            return SerialCommandResult(False, error)


def write_serial_command(serial_obj, command, append_crlf=True, push_debug=None):
    return write_serial_command_result(
        serial_obj,
        command,
        append_crlf=append_crlf,
        push_debug=push_debug,
    ).ok


def write_serial_command_sequence(serial_obj, commands, push_debug=None, delay_sec=0.3, sleep_func=time.sleep):
    if not _is_serial_open(serial_obj):
        if push_debug:
            push_debug(">>> 发送失败: 串口未连接")
        return False

    try:
        for index, command in enumerate(commands):
            serial_obj.write((command + "\r\n").encode("utf-8"))
            serial_obj.flush()
            if push_debug:
                push_debug(f">>> 发送: {command}\\r\\n")
            if index < len(commands) - 1:
                sleep_func(delay_sec)
        return True
    except Exception as exc:
        if push_debug:
            push_debug(f">>> 发送失败: {exc}")
        return False


def write_text_sms_pdu(
    serial_obj,
    phone,
    message,
    push_debug=None,
    port_ui=None,
    sleep_func=time.sleep,
    response_coordinator=None,
    segment_timeout=SMS_PDU_SEND_DEFAULT_TIMEOUT,
):
    if not _is_serial_open(serial_obj):
        if push_debug:
            push_debug(">>> 发送失败: 串口未连接")
        if port_ui:
            port_ui("❌ 发送短信失败：串口未连接", "normal")
        return False

    try:
        pdus = encode_text_sms_pdus(phone, message)
        if response_coordinator is None:
            return _sms_send_error(
                "未配置短信发送确认器，无法确认 Modem 发送结果",
                push_debug,
                port_ui,
            )
        if port_ui:
            port_ui(f"📤 发送短信至 {phone}：", "normal")
            port_ui(message, "sms")

        with DEFAULT_SERIAL_WRITE_LOCK:
            command = "AT+CMGF=0"
            serial_obj.write((command + "\r\n").encode("utf-8"))
            serial_obj.flush()
            if push_debug:
                push_debug(f">>> 发送: {command}\\r\\n")
            sleep_func(0.3)

            for index, (pdu_str, cmgs_len) in enumerate(pdus, start=1):
                command = f"AT+CMGS={cmgs_len}"
                serial_obj.write((command + "\r\n").encode("utf-8"))
                serial_obj.flush()
                if push_debug:
                    suffix = f" ({index}/{len(pdus)})" if len(pdus) > 1 else ""
                    push_debug(f">>> 发送: {command}{suffix}\\r\\n")
                sleep_func(1.0)

                suffix = f" ({index}/{len(pdus)})" if len(pdus) > 1 else ""
                waiter, begin_result = _begin_sms_segment_response(response_coordinator, suffix)
                if not begin_result.ok:
                    return _sms_send_error(begin_result.error, push_debug, port_ui)
                try:
                    serial_obj.write(pdu_str.encode("utf-8") + b"\x1a")
                    serial_obj.flush()
                except Exception as exc:
                    error = str(exc) or exc.__class__.__name__
                    _cancel_sms_segment_response(waiter, error)
                    _finish_sms_segment_response(response_coordinator, waiter)
                    return _sms_send_error(error, push_debug, port_ui)
                if push_debug:
                    push_debug(f">>> 发送 PDU 正文及 Ctrl+Z{suffix}，等待模组响应...")
                response = _wait_sms_segment_response(waiter, segment_timeout)
                _finish_sms_segment_response(response_coordinator, waiter)
                if not response.ok:
                    return _sms_send_error(response.error, push_debug, port_ui)
                if push_debug:
                    push_debug(f">>> 短信分片发送确认成功{suffix}")
                if index < len(pdus):
                    sleep_func(1.5)
        return True
    except Exception as exc:
        if push_debug:
            push_debug(f">>> 发送失败: {exc}")
        if port_ui:
            port_ui(f"❌ 发送短信失败：{exc}", "normal")
        return False


def _serial_available_locked(serial_lock, get_serial, push_debug=None, port_ui=None):
    with serial_lock:
        serial_obj = get_serial()
        ok = _is_serial_open(serial_obj)
    if ok:
        return True

    error = "串口未连接"
    if push_debug:
        push_debug(f">>> 发送失败: {error}")
    if port_ui:
        port_ui(f"❌ 发送短信失败：{error}", "normal")
    return False


def _write_bytes_locked(serial_lock, get_serial, payload, debug_message=None, push_debug=None):
    with serial_lock:
        serial_obj = get_serial()
        if not _is_serial_open(serial_obj):
            error = "串口未连接"
            if push_debug:
                push_debug(f">>> 发送失败: {error}")
            return SerialCommandResult(False, error)

        try:
            serial_obj.write(payload)
            serial_obj.flush()
        except Exception as exc:
            error = str(exc) or exc.__class__.__name__
            if push_debug:
                push_debug(f">>> 发送失败: {error}")
            return SerialCommandResult(False, error)

    if debug_message and push_debug:
        push_debug(debug_message)
    return SerialCommandResult(True)


def write_serial_command_sequence_locked(
    serial_lock,
    get_serial,
    commands,
    push_debug=None,
    delay_sec=0.3,
    sleep_func=time.sleep,
):
    with DEFAULT_SERIAL_WRITE_LOCK:
        commands = list(commands or [])
        if not _serial_available_locked(serial_lock, get_serial, push_debug=push_debug):
            return False

        for index, command in enumerate(commands):
            result = _write_bytes_locked(
                serial_lock,
                get_serial,
                (command + "\r\n").encode("utf-8"),
                f">>> 发送: {command}\\r\\n",
                push_debug,
            )
            if not result.ok:
                return False
            if index < len(commands) - 1:
                sleep_func(delay_sec)
        return True


def write_text_sms_pdu_locked(
    serial_lock,
    get_serial,
    phone,
    message,
    push_debug=None,
    port_ui=None,
    sleep_func=time.sleep,
    response_coordinator=None,
    segment_timeout=SMS_PDU_SEND_DEFAULT_TIMEOUT,
):
    try:
        pdus = encode_text_sms_pdus(phone, message)
        if response_coordinator is None:
            return _sms_send_error(
                "未配置短信发送确认器，无法确认 Modem 发送结果",
                push_debug,
                port_ui,
            )
        if port_ui:
            port_ui(f"📤 发送短信至 {phone}：", "normal")
            port_ui(message, "sms")

        with DEFAULT_SERIAL_WRITE_LOCK:
            if not _serial_available_locked(
                serial_lock,
                get_serial,
                push_debug=push_debug,
                port_ui=port_ui,
            ):
                return False

            command = "AT+CMGF=0"
            result = _write_bytes_locked(
                serial_lock,
                get_serial,
                (command + "\r\n").encode("utf-8"),
                f">>> 发送: {command}\\r\\n",
                push_debug,
            )
            if not result.ok:
                if port_ui:
                    port_ui(f"❌ 发送短信失败：{result.error}", "normal")
                return False
            sleep_func(0.3)

            for index, (pdu_str, cmgs_len) in enumerate(pdus, start=1):
                suffix = f" ({index}/{len(pdus)})" if len(pdus) > 1 else ""
                command = f"AT+CMGS={cmgs_len}"
                result = _write_bytes_locked(
                    serial_lock,
                    get_serial,
                    (command + "\r\n").encode("utf-8"),
                    f">>> 发送: {command}{suffix}\\r\\n",
                    push_debug,
                )
                if not result.ok:
                    if port_ui:
                        port_ui(f"❌ 发送短信失败：{result.error}", "normal")
                    return False
                sleep_func(1.0)

                waiter, begin_result = _begin_sms_segment_response(response_coordinator, suffix)
                if not begin_result.ok:
                    return _sms_send_error(begin_result.error, push_debug, port_ui)
                result = _write_bytes_locked(
                    serial_lock,
                    get_serial,
                    pdu_str.encode("utf-8") + b"\x1a",
                    f">>> 发送 PDU 正文及 Ctrl+Z{suffix}，等待模组响应...",
                    push_debug,
                )
                if not result.ok:
                    _cancel_sms_segment_response(waiter, result.error)
                    _finish_sms_segment_response(response_coordinator, waiter)
                    if port_ui:
                        port_ui(f"❌ 发送短信失败：{result.error}", "normal")
                    return False
                response = _wait_sms_segment_response(waiter, segment_timeout)
                _finish_sms_segment_response(response_coordinator, waiter)
                if not response.ok:
                    return _sms_send_error(response.error, push_debug, port_ui)
                if push_debug:
                    push_debug(f">>> 短信分片发送确认成功{suffix}")
                if index < len(pdus):
                    sleep_func(1.5)
        return True
    except Exception as exc:
        if push_debug:
            push_debug(f">>> 发送失败: {exc}")
        if port_ui:
            port_ui(f"❌ 发送短信失败：{exc}", "normal")
        return False


def _run_with_serial(serial_lock, get_serial, worker):
    with serial_lock:
        worker(get_serial())


def _run_serial_command_with_result(serial_lock, get_serial, command, append_crlf, push_debug, on_result):
    with serial_lock:
        result = write_serial_command_result(
            get_serial(),
            command,
            append_crlf=append_crlf,
            push_debug=push_debug,
        )
    if on_result:
        on_result(result)


def send_command_async(serial_lock, get_serial, command, append_crlf=True, push_debug=None, log_error=None):
    return start_daemon_thread(
        "serial_send_command",
        lambda: _run_with_serial(
            serial_lock,
            get_serial,
            lambda serial_obj: write_serial_command(
                serial_obj,
                command,
                append_crlf=append_crlf,
                push_debug=push_debug,
            ),
        ),
        log_error=log_error,
    )


def send_command_with_result_async(
    serial_lock,
    get_serial,
    command,
    append_crlf=True,
    push_debug=None,
    on_result=None,
    log_error=None,
):
    return start_daemon_thread(
        "serial_send_command_with_result",
        lambda: _run_serial_command_with_result(
            serial_lock,
            get_serial,
            command,
            append_crlf,
            push_debug,
            on_result,
        ),
        log_error=log_error,
    )


def send_command_sequence_async(
    serial_lock,
    get_serial,
    commands,
    push_debug=None,
    delay_sec=0.3,
    log_error=None,
    sleep_func=time.sleep,
):
    return start_daemon_thread(
        "serial_send_sequence",
        lambda: write_serial_command_sequence_locked(
            serial_lock,
            get_serial,
            commands,
            push_debug=push_debug,
            delay_sec=delay_sec,
            sleep_func=sleep_func,
        ),
        log_error=log_error,
    )


def send_text_sms_pdu_async(
    serial_lock,
    get_serial,
    phone,
    message,
    push_debug=None,
    port_ui=None,
    log_error=None,
    sleep_func=time.sleep,
    response_coordinator=None,
    segment_timeout=SMS_PDU_SEND_DEFAULT_TIMEOUT,
):
    coordinator = response_coordinator or DEFAULT_SMS_PDU_SEND_COORDINATOR
    return start_daemon_thread(
        "serial_send_sms_pdu",
        lambda: write_text_sms_pdu_locked(
            serial_lock,
            get_serial,
            phone,
            message,
            push_debug=push_debug,
            port_ui=port_ui,
            sleep_func=sleep_func,
            response_coordinator=coordinator,
            segment_timeout=segment_timeout,
        ),
        log_error=log_error,
    )
