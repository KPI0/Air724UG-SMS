from dataclasses import dataclass
import re
import threading
import time

from sms_core.serial_debug import build_serial_command_payload
from sms_core.sms_pdu import encode_text_sms_pdus
from sms_core.threading_runtime import WorkerThreadRegistry, start_daemon_thread


@dataclass(frozen=True)
class SerialCommandResult:
    ok: bool
    error: str = ""


@dataclass(frozen=True)
class SmsPduSendResponse:
    ok: bool
    error: str = ""
    line: str = ""


SERIAL_LOG_PREFIX_RE = re.compile(
    r"^\s*\[[IWE]\]-\[(?P<source>[^\]]+)\]\s*(?P<body>.*?)\s*$"
)
CMGS_RESULT_RE = re.compile(r"^\+CMGS\s*:\s*\d+\s*$", re.IGNORECASE)
CMGS_LUAT_RESULT_RE = re.compile(
    r"\+CMGS\s+AT\+CMGS=\d+.*(?:TRUE\s+OK|FALSE\s+ERROR|ERROR)\s*$",
    re.IGNORECASE,
)
MODEM_ERROR_RE = re.compile(
    r"^\+(?:CMS|CME)\s+ERROR(?:\s*:\s*.*)?$",
    re.IGNORECASE,
)
TRUSTED_MODEM_LOG_SOURCE = "ril.proatc"
TRUSTED_CMGS_RESULT_SOURCE = "lib_sms rsp"
SMS_PDU_SEND_DEFAULT_TIMEOUT = 45.0
SMS_PDU_PROMPT_TIMEOUT = 10.0
DEFAULT_SERIAL_TRANSACTION_LOCK = threading.RLock()
DEFAULT_SERIAL_WRITE_LOCK = threading.RLock()
_UNSET_CONNECTION = object()


class SmsPduSendWaiter:
    def __init__(self, label="", connection=None, cmgs_len=None):
        self.label = str(label or "")
        self.connection = connection
        self.cmgs_len = cmgs_len
        self._event = threading.Event()
        self._lock = threading.RLock()
        self._response = None
        self._saw_cmgs = False
        self._result_phase = False
        self._prompt_or_terminal_event = threading.Event()
        self._prompt_seen = False

    def wait(self, timeout):
        if self._event.wait(float(timeout)):
            with self._lock:
                return self._response or SmsPduSendResponse(False, "短信发送结果未知")
        return SmsPduSendResponse(False, "等待短信发送确认超时")

    def wait_prompt(self, timeout):
        self._prompt_or_terminal_event.wait(float(timeout))
        with self._lock:
            response = self._response
            prompt_seen = self._prompt_seen
        if response is not None and not response.ok:
            return response
        if prompt_seen:
            return SmsPduSendResponse(True)
        if response is not None:
            return SmsPduSendResponse(False, response.error or "未收到 Modem 的 > 提示符", response.line)
        return SmsPduSendResponse(False, "等待 Modem 的 > 提示符超时")

    def cancel(self, error="短信发送已取消"):
        self.fail(error)

    def mark_pdu_send_started(self):
        with self._lock:
            if self._response is not None or not self._prompt_seen:
                return False
            self._result_phase = True
            return True

    def write_pdu_if_ready(self, writer):
        """Atomically validate the prompt state and commit the PDU write."""
        with self._lock:
            if self._response is not None:
                return SerialCommandResult(
                    False,
                    self._response.error or "短信发送确认状态已失效",
                )
            if not self._prompt_seen:
                return SerialCommandResult(False, "未收到 Modem 的 > 提示符")
            self._result_phase = True
            return writer()

    def fail(self, error, line=""):
        self._complete(SmsPduSendResponse(False, str(error or "短信发送失败"), str(line or "")))

    def observe_line(self, line):
        body = _modem_response_body(line)
        if not body:
            return

        upper = body.upper()
        if _is_sms_send_error_line(line, body, upper, self.cmgs_len):
            self._complete(SmsPduSendResponse(False, body, str(line or "")))
            return

        if _is_sms_prompt_line(line, body):
            with self._lock:
                self._prompt_seen = True
            self._prompt_or_terminal_event.set()
            return

        with self._lock:
            result_phase = self._result_phase
        if result_phase and _is_cmgs_result_line(line, body, upper, self.cmgs_len):
            self._saw_cmgs = True
            if _response_has_final_ok(upper):
                self._complete(SmsPduSendResponse(True, "", str(line or "")))
            return

        if result_phase and self._saw_cmgs and _is_final_ok_line(line, body):
            self._complete(SmsPduSendResponse(True, "", str(line or "")))

    def done(self):
        return self._event.is_set()

    def _complete(self, response):
        with self._lock:
            if self._event.is_set():
                return
            self._response = response
            self._event.set()
            self._prompt_or_terminal_event.set()


class SmsPduSendCoordinator:
    def __init__(self):
        self._lock = threading.Lock()
        self._active = None

    def begin_segment(self, label="", connection=None, cmgs_len=None):
        waiter = SmsPduSendWaiter(label, connection=connection, cmgs_len=cmgs_len)
        with self._lock:
            if self._active is not None and not self._active.done():
                waiter.fail("已有短信分片正在等待 Modem 确认")
                return waiter
            self._active = waiter
        return waiter

    def observe_line(self, line, connection=_UNSET_CONNECTION):
        with self._lock:
            waiter = self._active
        if (
            waiter is not None
            and waiter.connection is not None
            and connection is not _UNSET_CONNECTION
            and waiter.connection is not connection
        ):
            return
        if waiter is not None:
            waiter.observe_line(line)

    def finish(self, waiter):
        with self._lock:
            if self._active is waiter:
                self._active = None

    def cancel_active(self, error="短信发送已取消"):
        with self._lock:
            waiter = self._active
        if waiter is None or waiter.done():
            return False
        waiter.cancel(error)
        return True


DEFAULT_SMS_PDU_SEND_COORDINATOR = SmsPduSendCoordinator()
DEFAULT_SMS_SEND_THREAD_REGISTRY = WorkerThreadRegistry()
DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY = WorkerThreadRegistry()


def _is_serial_open(serial_obj):
    return serial_obj is not None and bool(getattr(serial_obj, "is_open", False))


def _modem_response_body(line):
    text = str(line or "").strip()
    match = SERIAL_LOG_PREFIX_RE.match(text)
    return (match.group("body") if match else text).strip()


def _serial_log_source(line):
    match = SERIAL_LOG_PREFIX_RE.match(str(line or "").strip())
    return str(match.group("source") if match else "").strip().lower()


def _is_bare_modem_line(line, body):
    text = str(line or "").strip()
    return text == str(body or "").strip() and not SERIAL_LOG_PREFIX_RE.match(text)


def _is_trusted_modem_line(line, body):
    return (
        _is_bare_modem_line(line, body)
        or _serial_log_source(line) == TRUSTED_MODEM_LOG_SOURCE
    )


def _is_sms_prompt_line(line, body):
    return str(body or "").strip() == ">" and _is_trusted_modem_line(line, body)


def _is_final_ok_line(line, body):
    return str(body or "").strip().upper() == "OK" and _is_trusted_modem_line(line, body)


def _matches_current_cmgs_command(upper_body, cmgs_len):
    upper = str(upper_body or "").upper()
    if cmgs_len is None or "AT+CMGS=" not in upper:
        return True
    return f"AT+CMGS={cmgs_len}" in upper


def _is_luat_cmgs_result_line(line, body, upper_body, cmgs_len=None):
    if _serial_log_source(line) != TRUSTED_CMGS_RESULT_SOURCE:
        return False
    text = str(body or "").strip()
    upper = str(upper_body or "").strip()
    return bool(CMGS_LUAT_RESULT_RE.fullmatch(text)) and _matches_current_cmgs_command(
        upper,
        cmgs_len,
    )


def _is_cmgs_result_line(line, body, upper_body, cmgs_len=None):
    if CMGS_RESULT_RE.fullmatch(str(body or "").strip()):
        return _is_trusted_modem_line(line, body)
    return _is_luat_cmgs_result_line(line, body, upper_body, cmgs_len)


def _is_sms_send_error_line(line, body, upper_body, cmgs_len=None):
    text = str(body or "").strip()
    upper = str(upper_body or "").strip()
    if MODEM_ERROR_RE.fullmatch(text):
        return _is_trusted_modem_line(line, body)
    if _is_luat_cmgs_result_line(line, body, upper, cmgs_len):
        return "ERROR" in upper
    if upper == "ERROR":
        return _is_trusted_modem_line(line, body)
    return False


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


def _begin_sms_segment_response(response_coordinator, label, connection=None, cmgs_len=None):
    if response_coordinator is None:
        return None, SerialCommandResult(False, "未配置短信发送确认器，无法确认 Modem 发送结果")
    try:
        try:
            waiter = response_coordinator.begin_segment(
                label=label,
                connection=connection,
                cmgs_len=cmgs_len,
            )
        except TypeError:
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


def _wait_sms_prompt(waiter, timeout):
    try:
        wait_prompt = getattr(waiter, "wait_prompt", None)
        if not callable(wait_prompt):
            return SmsPduSendResponse(False, "短信确认器不支持等待 Modem 的 > 提示符")
        return wait_prompt(timeout)
    except Exception as exc:
        return SmsPduSendResponse(False, str(exc) or exc.__class__.__name__)


def _mark_sms_pdu_send_started(waiter):
    try:
        mark_started = getattr(waiter, "mark_pdu_send_started", None)
        if not callable(mark_started):
            return SmsPduSendResponse(False, "短信确认器不支持绑定 PDU 发送阶段")
        if not mark_started():
            return SmsPduSendResponse(False, "短信发送确认状态已失效")
        return SmsPduSendResponse(True)
    except Exception as exc:
        return SmsPduSendResponse(False, str(exc) or exc.__class__.__name__)


def _write_sms_pdu_after_prompt(waiter, writer):
    try:
        atomic_write = getattr(waiter, "write_pdu_if_ready", None)
        if callable(atomic_write):
            return atomic_write(writer)
    except Exception as exc:
        return SerialCommandResult(False, str(exc) or exc.__class__.__name__)

    phase_result = _mark_sms_pdu_send_started(waiter)
    if not phase_result.ok:
        return SerialCommandResult(False, phase_result.error)
    try:
        return writer()
    except Exception as exc:
        return SerialCommandResult(False, str(exc) or exc.__class__.__name__)


def _write_serial_obj_bytes(serial_obj, payload, debug_message=None, push_debug=None):
    with DEFAULT_SERIAL_WRITE_LOCK:
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


def write_serial_command_result(serial_obj, command, append_crlf=True, push_debug=None):
    with DEFAULT_SERIAL_TRANSACTION_LOCK:
        command_bytes, display_suffix = build_serial_command_payload(command, append_crlf)
        return _write_serial_obj_bytes(
            serial_obj,
            command_bytes,
            f">>> 发送: {command}{display_suffix}",
            push_debug,
        )


def write_serial_command(serial_obj, command, append_crlf=True, push_debug=None):
    return write_serial_command_result(
        serial_obj,
        command,
        append_crlf=append_crlf,
        push_debug=push_debug,
    ).ok


def write_serial_command_sequence(serial_obj, commands, push_debug=None, delay_sec=0.3, sleep_func=time.sleep):
    with DEFAULT_SERIAL_TRANSACTION_LOCK:
        for index, command in enumerate(commands):
            result = _write_serial_obj_bytes(
                serial_obj,
                (command + "\r\n").encode("utf-8"),
                f">>> 发送: {command}\\r\\n",
                push_debug,
            )
            if not result.ok:
                return False
            if index < len(commands) - 1:
                sleep_func(delay_sec)
        return True


def write_text_sms_pdu(
    serial_obj,
    phone,
    message,
    push_debug=None,
    port_ui=None,
    sleep_func=time.sleep,
    response_coordinator=None,
    segment_timeout=SMS_PDU_SEND_DEFAULT_TIMEOUT,
    prompt_timeout=None,
):
    if not _is_serial_open(serial_obj):
        if push_debug:
            push_debug(">>> 发送失败: 串口未连接")
        if port_ui:
            port_ui("❌ 发送短信失败：串口未连接", "normal")
        return False

    try:
        pdus = encode_text_sms_pdus(phone, message)
        if prompt_timeout is None:
            prompt_timeout = min(float(segment_timeout), SMS_PDU_PROMPT_TIMEOUT)
        if response_coordinator is None:
            return _sms_send_error(
                "未配置短信发送确认器，无法确认 Modem 发送结果",
                push_debug,
                port_ui,
            )
        if port_ui:
            port_ui(f"📤 发送短信至 {phone}：", "normal")
            port_ui(message, "sms")

        with DEFAULT_SERIAL_TRANSACTION_LOCK:
            command = "AT+CMGF=0"
            result = _write_serial_obj_bytes(
                serial_obj,
                (command + "\r\n").encode("utf-8"),
                f">>> 发送: {command}\\r\\n",
                push_debug,
            )
            if not result.ok:
                return _sms_send_error(result.error, push_debug, port_ui)
            sleep_func(0.3)

            for index, (pdu_str, cmgs_len) in enumerate(pdus, start=1):
                command = f"AT+CMGS={cmgs_len}"
                suffix = f" ({index}/{len(pdus)})" if len(pdus) > 1 else ""
                waiter, begin_result = _begin_sms_segment_response(
                    response_coordinator,
                    suffix,
                    connection=serial_obj,
                    cmgs_len=cmgs_len,
                )
                if not begin_result.ok:
                    return _sms_send_error(begin_result.error, push_debug, port_ui)
                result = _write_serial_obj_bytes(
                    serial_obj,
                    (command + "\r\n").encode("utf-8"),
                    f">>> 发送: {command}{suffix}\\r\\n",
                    push_debug,
                )
                if not result.ok:
                    _cancel_sms_segment_response(waiter, result.error)
                    _finish_sms_segment_response(response_coordinator, waiter)
                    return _sms_send_error(result.error, push_debug, port_ui)
                prompt_response = _wait_sms_prompt(waiter, prompt_timeout)
                if not prompt_response.ok:
                    _cancel_sms_segment_response(waiter, prompt_response.error)
                    _finish_sms_segment_response(response_coordinator, waiter)
                    return _sms_send_error(prompt_response.error, push_debug, port_ui)
                result = _write_sms_pdu_after_prompt(
                    waiter,
                    lambda: _write_serial_obj_bytes(
                        serial_obj,
                        pdu_str.encode("utf-8") + b"\x1a",
                        f">>> 发送 PDU 正文及 Ctrl+Z{suffix}，等待模组响应...",
                        push_debug,
                    ),
                )
                if not result.ok:
                    _cancel_sms_segment_response(waiter, result.error)
                    _finish_sms_segment_response(response_coordinator, waiter)
                    return _sms_send_error(result.error, push_debug, port_ui)
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


def _write_bytes_locked(
    serial_lock,
    get_serial,
    payload,
    debug_message=None,
    push_debug=None,
    expected_serial=None,
):
    with serial_lock:
        serial_obj = get_serial()
        if expected_serial is not None and serial_obj is not expected_serial:
            error = "串口连接已变化"
            if push_debug:
                push_debug(f">>> 发送失败: {error}")
            return SerialCommandResult(False, error)
        if not _is_serial_open(serial_obj):
            error = "串口未连接"
            if push_debug:
                push_debug(f">>> 发送失败: {error}")
            return SerialCommandResult(False, error)
        return _write_serial_obj_bytes(
            serial_obj,
            payload,
            debug_message,
            push_debug,
        )


def write_serial_command_sequence_locked(
    serial_lock,
    get_serial,
    commands,
    push_debug=None,
    delay_sec=0.3,
    sleep_func=time.sleep,
):
    with DEFAULT_SERIAL_TRANSACTION_LOCK:
        commands = list(commands or [])
        with serial_lock:
            transaction_serial = get_serial()
        if not _is_serial_open(transaction_serial):
            if push_debug:
                push_debug(">>> 发送失败: 串口未连接")
            return False

        for index, command in enumerate(commands):
            result = _write_bytes_locked(
                serial_lock,
                get_serial,
                (command + "\r\n").encode("utf-8"),
                f">>> 发送: {command}\\r\\n",
                push_debug,
                expected_serial=transaction_serial,
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
    prompt_timeout=None,
):
    try:
        pdus = encode_text_sms_pdus(phone, message)
        if prompt_timeout is None:
            prompt_timeout = min(float(segment_timeout), SMS_PDU_PROMPT_TIMEOUT)
        if response_coordinator is None:
            return _sms_send_error(
                "未配置短信发送确认器，无法确认 Modem 发送结果",
                push_debug,
                port_ui,
            )
        if port_ui:
            port_ui(f"📤 发送短信至 {phone}：", "normal")
            port_ui(message, "sms")

        with DEFAULT_SERIAL_TRANSACTION_LOCK:
            with serial_lock:
                transaction_serial = get_serial()
            if not _is_serial_open(transaction_serial):
                _sms_send_error("串口未连接", push_debug, port_ui)
                return False

            command = "AT+CMGF=0"
            result = _write_bytes_locked(
                serial_lock,
                get_serial,
                (command + "\r\n").encode("utf-8"),
                f">>> 发送: {command}\\r\\n",
                push_debug,
                expected_serial=transaction_serial,
            )
            if not result.ok:
                if port_ui:
                    port_ui(f"❌ 发送短信失败：{result.error}", "normal")
                return False
            sleep_func(0.3)

            for index, (pdu_str, cmgs_len) in enumerate(pdus, start=1):
                suffix = f" ({index}/{len(pdus)})" if len(pdus) > 1 else ""
                command = f"AT+CMGS={cmgs_len}"
                waiter, begin_result = _begin_sms_segment_response(
                    response_coordinator,
                    suffix,
                    connection=transaction_serial,
                    cmgs_len=cmgs_len,
                )
                if not begin_result.ok:
                    return _sms_send_error(begin_result.error, push_debug, port_ui)
                result = _write_bytes_locked(
                    serial_lock,
                    get_serial,
                    (command + "\r\n").encode("utf-8"),
                    f">>> 发送: {command}{suffix}\\r\\n",
                    push_debug,
                    expected_serial=transaction_serial,
                )
                if not result.ok:
                    _cancel_sms_segment_response(waiter, result.error)
                    _finish_sms_segment_response(response_coordinator, waiter)
                    if port_ui:
                        port_ui(f"❌ 发送短信失败：{result.error}", "normal")
                    return False
                prompt_response = _wait_sms_prompt(waiter, prompt_timeout)
                if not prompt_response.ok:
                    _cancel_sms_segment_response(waiter, prompt_response.error)
                    _finish_sms_segment_response(response_coordinator, waiter)
                    return _sms_send_error(prompt_response.error, push_debug, port_ui)
                result = _write_sms_pdu_after_prompt(
                    waiter,
                    lambda: _write_bytes_locked(
                        serial_lock,
                        get_serial,
                        pdu_str.encode("utf-8") + b"\x1a",
                        f">>> 发送 PDU 正文及 Ctrl+Z{suffix}，等待模组响应...",
                        push_debug,
                        expected_serial=transaction_serial,
                    ),
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
        serial_obj = get_serial()
    worker(serial_obj)


def _run_serial_command_with_result(serial_lock, get_serial, command, append_crlf, push_debug, on_result):
    with serial_lock:
        serial_obj = get_serial()
    result = write_serial_command_result(
        serial_obj,
        command,
        append_crlf=append_crlf,
        push_debug=push_debug,
    )
    if on_result:
        on_result(result)


def start_registered_serial_worker(
    name,
    target,
    *,
    log_error=None,
    thread_registry=DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY,
    thread_factory=threading.Thread,
):
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


def send_command_async(
    serial_lock,
    get_serial,
    command,
    append_crlf=True,
    push_debug=None,
    log_error=None,
    thread_registry=DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY,
):
    return start_registered_serial_worker(
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
        thread_registry=thread_registry,
    )


def send_command_with_result_async(
    serial_lock,
    get_serial,
    command,
    append_crlf=True,
    push_debug=None,
    on_result=None,
    log_error=None,
    thread_registry=DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY,
):
    return start_registered_serial_worker(
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
        thread_registry=thread_registry,
    )


def send_command_sequence_async(
    serial_lock,
    get_serial,
    commands,
    push_debug=None,
    delay_sec=0.3,
    log_error=None,
    sleep_func=time.sleep,
    thread_registry=DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY,
):
    return start_registered_serial_worker(
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
        thread_registry=thread_registry,
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
    prompt_timeout=None,
    thread_registry=DEFAULT_SMS_SEND_THREAD_REGISTRY,
):
    coordinator = response_coordinator or DEFAULT_SMS_PDU_SEND_COORDINATOR
    thread_holder = {}

    def register_thread(thread):
        thread_holder["thread"] = thread
        if thread_registry is not None:
            thread_registry.register(thread)

    def run_send():
        try:
            return write_text_sms_pdu_locked(
                serial_lock,
                get_serial,
                phone,
                message,
                push_debug=push_debug,
                port_ui=port_ui,
                sleep_func=sleep_func,
                response_coordinator=coordinator,
                segment_timeout=segment_timeout,
                prompt_timeout=prompt_timeout,
            )
        finally:
            if thread_registry is not None:
                thread_registry.unregister(thread_holder.get("thread"))

    try:
        return start_daemon_thread(
            "serial_send_sms_pdu",
            run_send,
            log_error=log_error,
            before_start=register_thread,
        )
    except Exception:
        if thread_registry is not None:
            thread_registry.unregister(thread_holder.get("thread"))
        raise
