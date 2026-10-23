from dataclasses import dataclass
from contextlib import contextmanager, nullcontext
import re
import threading
import time

from sms_core.serial_debug import build_serial_command_payload
from sms_core.call_command_context import VOICE_DIAL_RE
from sms_core.phone_numbers import normalize_sms_destination
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


@dataclass(frozen=True)
class AtCommandResponse:
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
CALL_TERMINAL_RESPONSE_RE = re.compile(
    r"^(?:NO\s+CARRIER|NO\s+ANSWER|BUSY)$",
    re.IGNORECASE,
)
TRUSTED_MODEM_LOG_SOURCE = "ril.proatc"
TRUSTED_CMGS_RESULT_SOURCE = "lib_sms rsp"
SMS_PDU_SEND_DEFAULT_TIMEOUT = 45.0
SMS_PDU_PROMPT_TIMEOUT = 10.0
SMS_PDU_ABORT_TIMEOUT = 2.0
AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT = 10.0
DEFAULT_SERIAL_TRANSACTION_LOCK = threading.RLock()
DEFAULT_SERIAL_WRITE_LOCK = threading.RLock()

# OTA reservations reject new commands instead of queueing them until reboot.
# The write lock protects both reservation changes and the final serial write.
_OTA_RESERVATIONS = {}


def reserve_serial_for_ota(serial_obj, owner, *, seconds=150):
    with DEFAULT_SERIAL_WRITE_LOCK:
        previous = _OTA_RESERVATIONS.get(id(serial_obj))
        if previous and previous[1] is owner and previous[2] > time.monotonic():
            _OTA_RESERVATIONS[id(serial_obj)] = (serial_obj, owner, time.monotonic() + seconds)
            return _is_serial_open(serial_obj)
    if not DEFAULT_SERIAL_TRANSACTION_LOCK.acquire(blocking=False):
        return False
    try:
        with DEFAULT_SERIAL_WRITE_LOCK:
            now = time.monotonic()
            for key, (_, _, expiry) in list(_OTA_RESERVATIONS.items()):
                if expiry <= now:
                    _OTA_RESERVATIONS.pop(key, None)
            previous = _OTA_RESERVATIONS.get(id(serial_obj))
            if not _is_serial_open(serial_obj) or (previous and previous[1] is not owner):
                return False
            _OTA_RESERVATIONS[id(serial_obj)] = (serial_obj, owner, now + seconds)
            return True
    finally:
        DEFAULT_SERIAL_TRANSACTION_LOCK.release()


def release_serial_from_ota(serial_obj, owner):
    with DEFAULT_SERIAL_WRITE_LOCK:
        previous = _OTA_RESERVATIONS.get(id(serial_obj))
        if previous and previous[1] is owner:
            _OTA_RESERVATIONS.pop(id(serial_obj), None)
_UNSET_CONNECTION = object()


class SmsPduSendWaiter:
    def __init__(self, label="", connection=None, cmgs_len=None, *, wait_for_write=False):
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
        self._write_started = not wait_for_write
        self._terminal_event = threading.Event()
        self._abort_started = False

    def mark_write_started(self):
        with self._lock:
            if self._response is not None:
                return False
            self._write_started = True
            return True

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

    def terminal_received(self):
        return self._terminal_event.is_set()

    def begin_abort(self):
        with self._lock:
            if self._terminal_event.is_set():
                return False
            self._abort_started = True
            return True

    def wait_terminal(self, timeout):
        # Cancellation wakes the caller, but does not end the modem transaction.
        return self._terminal_event.wait(float(timeout))

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
        with self._lock:
            if not self._write_started:
                return
        body = _modem_response_body(line)
        if not body:
            return

        upper = body.upper()
        if _is_sms_send_error_line(line, body, upper, self.cmgs_len):
            self._terminal_event.set()
            self._complete(SmsPduSendResponse(False, body, str(line or "")))
            return True

        if self._abort_started and _is_final_ok_line(line, body):
            self._terminal_event.set()
            return True

        if _is_sms_prompt_line(line, body):
            with self._lock:
                self._prompt_seen = True
            self._prompt_or_terminal_event.set()
            return True

        with self._lock:
            result_phase = self._result_phase
        if result_phase and _is_cmgs_result_line(line, body, upper, self.cmgs_len):
            self._saw_cmgs = True
            if _response_has_final_ok(upper):
                self._terminal_event.set()
                self._complete(SmsPduSendResponse(True, "", str(line or "")))
            return True

        if result_phase and self._saw_cmgs and _is_final_ok_line(line, body):
            self._terminal_event.set()
            self._complete(SmsPduSendResponse(True, "", str(line or "")))
        return _is_final_ok_line(line, body)

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

    def begin_segment(self, label="", connection=None, cmgs_len=None, *, wait_for_write=False):
        waiter = SmsPduSendWaiter(
            label, connection=connection, cmgs_len=cmgs_len, wait_for_write=wait_for_write,
        )
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
            return waiter.observe_line(line)

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


class AtCommandResponseWaiter:
    def __init__(self, command="", connection=None, *, wait_for_write=False):
        self.command = str(command or "").strip()
        self.connection = connection
        self._event = threading.Event()
        self._lock = threading.RLock()
        self._response = None
        self._write_started = not wait_for_write

    def mark_write_started(self):
        with self._lock:
            if self._response is not None:
                return False
            self._write_started = True
            return True

    def wait(self, timeout):
        if self._event.wait(float(timeout)):
            with self._lock:
                return self._response or AtCommandResponse(False, "AT 指令执行结果未知")
        return AtCommandResponse(False, "等待 Modem 指令响应超时")

    def cancel(self, error="AT 指令执行已取消"):
        self._complete(AtCommandResponse(False, str(error or "AT 指令执行已取消")))

    def observe_line(self, line):
        with self._lock:
            if not self._write_started:
                return
        body = _modem_response_body(line)
        if not body or not _is_trusted_modem_line(line, body):
            return

        upper = body.upper()
        # ATD has valid modem-level terminal responses that are not reported
        # as ERROR. Restrict this to ATD so an unrelated unsolicited call URC
        # cannot fail another command's waiter.
        if self.command.upper().startswith("ATD") and CALL_TERMINAL_RESPONSE_RE.fullmatch(body):
            self._complete(AtCommandResponse(False, body, str(line or "")))
            return
        if upper == "ERROR" or MODEM_ERROR_RE.fullmatch(body):
            self._complete(AtCommandResponse(False, body, str(line or "")))
            return
        if _is_final_ok_line(line, body):
            self._complete(AtCommandResponse(True, "", str(line or "")))

    def done(self):
        return self._event.is_set()

    def _complete(self, response):
        with self._lock:
            if self._event.is_set():
                return
            self._response = response
            self._event.set()


class AtCommandResponseCoordinator:
    def __init__(self):
        self._lock = threading.Lock()
        self._active = None

    def begin(self, command="", connection=None):
        waiter = AtCommandResponseWaiter(command, connection=connection, wait_for_write=True)
        with self._lock:
            if self._active is not None and not self._active.done():
                waiter.cancel("已有 AT 指令正在等待 Modem 响应")
                return waiter
            self._active = waiter
        return waiter

    def observe_line(self, line, connection=_UNSET_CONNECTION, *, call_state_snapshot=None):
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
            command = waiter.command.upper()
            if (call_state_snapshot is not None
                    and (VOICE_DIAL_RE.fullmatch(command) or re.fullmatch(r"AT[AH];?", command))
                    and not call_state_snapshot.is_response_current()):
                return
            waiter.observe_line(line)

    def finish(self, waiter):
        with self._lock:
            if self._active is waiter:
                self._active = None

    def cancel_active(self, error="AT 指令执行已取消"):
        with self._lock:
            waiter = self._active
        if waiter is None or waiter.done():
            return False
        waiter.cancel(error)
        return True


DEFAULT_SMS_PDU_SEND_COORDINATOR = SmsPduSendCoordinator()
DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR = AtCommandResponseCoordinator()
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
    command = re.search(r"AT\+CMGS=(\d+)", upper)
    return command is not None and command.group(1) == str(cmgs_len)


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
                wait_for_write=True,
            )
        except TypeError:
            try:
                waiter = response_coordinator.begin_segment(
                    label=label, connection=connection, cmgs_len=cmgs_len,
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


def _write_serial_obj_bytes(
    serial_obj, payload, debug_message=None, push_debug=None, *,
    write_guard=None, before_write=None, after_write=None, sms_waiter=None, flush_output=True,
):
    # Validate after every sender lock has been acquired, and keep the caller's
    # context locked through the write so a queued action cannot change calls.
    with DEFAULT_SERIAL_WRITE_LOCK, (write_guard() if write_guard else nullcontext(True)) as allowed:
        if not allowed:
            return SerialCommandResult(False, "指令所属状态已变化，已取消发送")
        recovery_error = serial_sms_recovery_error(serial_obj)
        if recovery_error:
            if push_debug:
                push_debug(f">>> 发送失败: {recovery_error}")
            return SerialCommandResult(False, recovery_error)
        reservation = _OTA_RESERVATIONS.get(id(serial_obj))
        if reservation:
            if reservation[2] > time.monotonic():
                if push_debug:
                    push_debug(">>> 发送失败: 固件更新中，请等待设备重启后再操作")
                return SerialCommandResult(False, "固件更新中，请等待设备重启后再操作")
            _OTA_RESERVATIONS.pop(id(serial_obj), None)
        if not _is_serial_open(serial_obj):
            error = "串口未连接"
            if push_debug:
                push_debug(f">>> 发送失败: {error}")
            return SerialCommandResult(False, error)

        try:
            def commit_bytes():
                if before_write is not None and before_write() is False:
                    return SerialCommandResult(False, "指令已取消，未写入串口")
                written = serial_obj.write(payload)
                if after_write is not None:
                    after_write(written)
                if written is not None and written != len(payload):
                    return SerialCommandResult(False, "串口写入不完整")
                return SerialCommandResult(True)

            result = (_write_sms_pdu_after_prompt(sms_waiter, commit_bytes)
                      if sms_waiter is not None else commit_bytes())
            if not result.ok:
                return result
            if flush_output:
                serial_obj.flush()
        except Exception as exc:
            error = str(exc) or exc.__class__.__name__
            if push_debug:
                push_debug(f">>> 发送失败: {error}")
            return SerialCommandResult(False, error)

    if debug_message and push_debug:
        push_debug(debug_message)
    return SerialCommandResult(True)


def serial_sms_recovery_error(serial_obj):
    error = getattr(serial_obj, "_sms_recovery_error", "")
    return error if isinstance(error, str) else ""


def _isolate_sms_serial(serial_obj):
    # Closing/reopening a host handle cannot prove the modem left PDU input.
    # Keep reads alive and reject every sender until the user resets the device.
    error = "短信事务收尾未获确认，已暂停此串口发送，请重启设备后重新连接"
    serial_obj._sms_recovery_error = error
    return error


def _send_sms_segment(
    response_coordinator, connection, cmgs_len, pdu_str, suffix,
    write_bytes, recover_bytes, isolate_serial, prompt_timeout, segment_timeout,
    abort_timeout,
):
    waiter, result = _begin_sms_segment_response(
        response_coordinator, suffix, connection=connection, cmgs_len=cmgs_len,
    )
    if not result.ok:
        return result
    cmgs_attempted = False
    pdu_may_be_submitted = False

    def begin_pdu():
        nonlocal pdu_may_be_submitted
        # A write exception may occur after Ctrl-Z reached the driver. Only an
        # explicit short count proves that the final byte was not submitted.
        pdu_may_be_submitted = True

    def record_pdu_write(written):
        nonlocal pdu_may_be_submitted
        if isinstance(written, int) and 0 <= written < len(pdu_str.encode("utf-8")) + 1:
            pdu_may_be_submitted = False

    def begin_cmgs():
        nonlocal cmgs_attempted
        marker = getattr(waiter, "mark_write_started", None)
        if callable(marker) and marker() is False:
            return False
        cmgs_attempted = True
        return True

    try:
        try:
            command = f"AT+CMGS={cmgs_len}"
            result = write_bytes(
                (command + "\r\n").encode("utf-8"),
                f">>> 发送: {command}{suffix}\\r\\n", before_write=begin_cmgs,
            )
            if result.ok:
                result = _wait_sms_prompt(waiter, prompt_timeout)
            if result.ok:
                result = write_bytes(
                    pdu_str.encode("utf-8") + b"\x1a",
                    f">>> 发送 PDU 正文及 Ctrl+Z{suffix}，等待模组响应...",
                    sms_waiter=waiter,
                    before_write=begin_pdu, after_write=record_pdu_write,
                )
            if result.ok:
                result = _wait_sms_segment_response(waiter, segment_timeout)
        except Exception as exc:
            result = SerialCommandResult(False, str(exc) or exc.__class__.__name__)
        if result.ok:
            return result

        _cancel_sms_segment_response(waiter, result.error)
        terminal_received = getattr(waiter, "terminal_received", lambda: False)
        if not cmgs_attempted or terminal_received():
            return result
        wait_terminal = getattr(waiter, "wait_terminal", None)
        recovered = False
        try:
            if not pdu_may_be_submitted:
                abort_result = recover_bytes(
                    b"\x1b", ">>> 取消短信输入，等待原事务结束",
                    before_write=getattr(waiter, "begin_abort", None),
                    flush_output=False,
                )
                recovered = terminal_received()
                can_wait = abort_result.ok
            else:
                # Ctrl+Z may already have submitted the SMS. Only its original
                # terminal response can safely release the transaction now.
                can_wait = True
            if not recovered and can_wait and callable(wait_terminal):
                recovered = wait_terminal(abort_timeout)
        except Exception:
            recovered = False
        if not recovered:
            recovery_error = isolate_serial()
            if recovery_error:
                return SerialCommandResult(False, f"{result.error}；{recovery_error}")
        return result
    finally:
        _finish_sms_segment_response(response_coordinator, waiter)


def write_serial_command_result(serial_obj, command, append_crlf=True, push_debug=None, *, write_guard=None):
    with DEFAULT_SERIAL_TRANSACTION_LOCK:
        command_bytes, display_suffix = build_serial_command_payload(command, append_crlf)
        return _write_serial_obj_bytes(
            serial_obj,
            command_bytes,
            f">>> 发送: {command}{display_suffix}",
            push_debug,
            write_guard=write_guard,
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
    *,
    abort_timeout=SMS_PDU_ABORT_TIMEOUT,
    command_response_coordinator=None,
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
            def isolate_serial():
                with DEFAULT_SERIAL_WRITE_LOCK:
                    return _isolate_sms_serial(serial_obj)

            result = _initialize_sms_pdu_mode(
                nullcontext(), lambda: serial_obj, serial_obj,
                command_response_coordinator or DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR,
                min(float(segment_timeout), AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT),
                isolate_serial, push_debug,
            )
            if not result.ok:
                return _sms_send_error(result.error, push_debug, port_ui)

            for index, (pdu_str, cmgs_len) in enumerate(pdus, start=1):
                suffix = f" ({index}/{len(pdus)})" if len(pdus) > 1 else ""

                def write_bytes(payload, debug_message, **kwargs):
                    return _write_serial_obj_bytes(
                        serial_obj, payload, debug_message, push_debug, **kwargs,
                    )

                response = _send_sms_segment(
                    response_coordinator, serial_obj, cmgs_len, pdu_str, suffix,
                    write_bytes, write_bytes, isolate_serial,
                    prompt_timeout, segment_timeout, abort_timeout,
                )
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
    before_write=None,
    after_write=None,
    call_write_guard=None,
    sms_waiter=None,
    flush_output=True,
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

        @contextmanager
        def write_guard():
            # Context-aware getters also validate authorization and generation.
            # Recheck after the final write-lock wait while serial_lock is held.
            with (call_write_guard() if call_write_guard else nullcontext(True)) as allowed:
                yield allowed and get_serial() is serial_obj

        return _write_serial_obj_bytes(
            serial_obj,
            payload,
            debug_message,
            push_debug,
            write_guard=write_guard,
            before_write=before_write,
            after_write=after_write,
            sms_waiter=sms_waiter,
            flush_output=flush_output,
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


def _write_confirmed_command_locked(
    serial_lock,
    get_serial,
    transaction_serial,
    command,
    response_coordinator,
    response_timeout,
    push_debug=None,
    *,
    append_crlf=True,
    on_modem_response=None,
    write_guard=None,
    before_write=None,
):
    if response_coordinator is None:
        return SerialCommandResult(False, "未配置 AT 指令响应确认器")
    try:
        waiter = response_coordinator.begin(
            command=command,
            connection=transaction_serial,
        )
    except Exception as exc:
        return SerialCommandResult(False, str(exc) or exc.__class__.__name__)

    try:
        def mark_write_started():
            marker = getattr(waiter, "mark_write_started", None)
            if callable(marker) and marker() is False:
                return False
            return before_write() if before_write is not None else True

        payload, display_suffix = build_serial_command_payload(command, append_crlf)
        result = _write_bytes_locked(
            serial_lock,
            get_serial,
            payload,
            f">>> 发送: {command}{display_suffix}" if push_debug else None,
            push_debug,
            expected_serial=transaction_serial,
            before_write=mark_write_started,
            call_write_guard=write_guard,
        )
        if not result.ok:
            try:
                waiter.cancel(result.error)
            except Exception:
                pass
            return result
        try:
            response = waiter.wait(response_timeout)
        except Exception as exc:
            response = AtCommandResponse(False, str(exc) or exc.__class__.__name__)
        if on_modem_response is not None:
            on_modem_response(response)
        return SerialCommandResult(bool(response.ok), str(response.error or ""))
    finally:
        response_coordinator.finish(waiter)


def write_serial_command_confirmed_locked(
    serial_lock,
    get_serial,
    command,
    *,
    response_coordinator,
    append_crlf=True,
    response_timeout=AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT,
    push_debug=None,
    on_modem_response=None,
    write_guard=None,
    before_write=None,
):
    """Confirm one command while preserving the serial editor's exact bytes."""
    with DEFAULT_SERIAL_TRANSACTION_LOCK:
        with serial_lock:
            transaction_serial = get_serial()
        if not _is_serial_open(transaction_serial):
            return SerialCommandResult(False, "串口未连接")
        return _write_confirmed_command_locked(
            serial_lock, get_serial, transaction_serial, command,
            response_coordinator, response_timeout, push_debug,
            append_crlf=append_crlf, on_modem_response=on_modem_response,
            write_guard=write_guard, before_write=before_write,
        )


def _initialize_sms_pdu_mode(
    serial_lock, get_serial, transaction_serial, response_coordinator,
    response_timeout, isolate_serial, push_debug=None, *, write_guard=None,
):
    attempted = False
    terminal_received = False

    def begin_write():
        nonlocal attempted
        attempted = True

    def on_response(response):
        nonlocal terminal_received
        # The AT waiter supplies a source-validated line for a modem terminal;
        # local cancellation and timeout responses have no received line.
        terminal_received = bool(getattr(response, "line", ""))

    result = _write_confirmed_command_locked(
        serial_lock, get_serial, transaction_serial, "AT+CMGF=0",
        response_coordinator, response_timeout, push_debug,
        before_write=begin_write, on_modem_response=on_response, write_guard=write_guard,
    )
    if not result.ok and attempted and not terminal_received:
        recovery_error = isolate_serial()
        if recovery_error:
            return SerialCommandResult(False, f"{result.error}；{recovery_error}")
    return result


def write_serial_command_sequence_confirmed_locked(
    serial_lock,
    get_serial,
    commands,
    *,
    response_coordinator,
    response_timeout=AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT,
    push_debug=None,
    write_guard=None,
    before_write=None,
):
    commands = [str(command or "").strip() for command in commands or ()]
    if not commands or any(not command for command in commands):
        return SerialCommandResult(False, "AT 指令序列不能为空")

    with DEFAULT_SERIAL_TRANSACTION_LOCK:
        with serial_lock:
            transaction_serial = get_serial()
        if not _is_serial_open(transaction_serial):
            return SerialCommandResult(False, "串口未连接")

        for command in commands:
            result = _write_confirmed_command_locked(
                serial_lock,
                get_serial,
                transaction_serial,
                command,
                response_coordinator,
                response_timeout,
                push_debug,
                write_guard=write_guard, before_write=before_write,
            )
            if not result.ok:
                return result
    return SerialCommandResult(True)


def write_text_sms_pdu_locked(
    serial_lock,
    get_serial,
    phone,
    message,
    push_debug=None,
    port_ui=None,
    sleep_func=time.sleep,
    response_coordinator=None,
    command_response_coordinator=None,
    segment_timeout=SMS_PDU_SEND_DEFAULT_TIMEOUT,
    prompt_timeout=None,
    *,
    get_recovery_serial=None,
    abort_timeout=SMS_PDU_ABORT_TIMEOUT,
    write_guard=None,
):
    try:
        phone = normalize_sms_destination(phone)
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

            recovery_getter = get_recovery_serial or get_serial

            def isolate_serial():
                with serial_lock, DEFAULT_SERIAL_WRITE_LOCK:
                    if recovery_getter() is transaction_serial:
                        return _isolate_sms_serial(transaction_serial)
                return ""

            result = _initialize_sms_pdu_mode(
                serial_lock,
                get_serial,
                transaction_serial,
                command_response_coordinator or DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR,
                min(float(segment_timeout), AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT),
                isolate_serial,
                push_debug,
                write_guard=write_guard,
            )
            if not result.ok:
                if port_ui:
                    port_ui(f"❌ 发送短信失败：{result.error}", "normal")
                return False

            for index, (pdu_str, cmgs_len) in enumerate(pdus, start=1):
                suffix = f" ({index}/{len(pdus)})" if len(pdus) > 1 else ""

                def write_bytes(payload, debug_message, **kwargs):
                    return _write_bytes_locked(
                        serial_lock, get_serial, payload, debug_message, push_debug,
                        expected_serial=transaction_serial, call_write_guard=write_guard, **kwargs,
                    )

                def recover_bytes(payload, debug_message, **kwargs):
                    return _write_bytes_locked(
                        serial_lock, recovery_getter, payload, debug_message, push_debug,
                        expected_serial=transaction_serial, **kwargs,
                    )

                response = _send_sms_segment(
                    response_coordinator, transaction_serial, cmgs_len, pdu_str, suffix,
                    write_bytes, recover_bytes, isolate_serial,
                    prompt_timeout, segment_timeout, abort_timeout,
                )
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


def _run_serial_command_with_result(serial_lock, get_serial, command, append_crlf, push_debug, on_result, write_guard=None):
    if write_guard is not None:
        # Match the transaction -> connection -> write order used by SMS and
        # confirmed commands. Do not release the connection before writing.
        with DEFAULT_SERIAL_TRANSACTION_LOCK, serial_lock:
            result = write_serial_command_result(
                get_serial(), command, append_crlf=append_crlf,
                push_debug=push_debug, write_guard=write_guard,
            )
    else:
        with serial_lock:
            serial_obj = get_serial()
        result = write_serial_command_result(
            serial_obj, command, append_crlf=append_crlf, push_debug=push_debug,
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
    *,
    write_guard=None,
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
            write_guard,
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
    on_sent=None,
    *,
    get_recovery_serial=None,
    abort_timeout=SMS_PDU_ABORT_TIMEOUT,
    command_response_coordinator=None,
):
    coordinator = response_coordinator or DEFAULT_SMS_PDU_SEND_COORDINATOR
    thread_holder = {}

    def register_thread(thread):
        thread_holder["thread"] = thread
        if thread_registry is not None:
            thread_registry.register(thread)

    def run_send():
        try:
            sent = write_text_sms_pdu_locked(
                serial_lock,
                get_serial,
                phone,
                message,
                push_debug=push_debug,
                port_ui=port_ui,
                sleep_func=sleep_func,
                response_coordinator=coordinator,
                command_response_coordinator=command_response_coordinator,
                segment_timeout=segment_timeout,
                prompt_timeout=prompt_timeout,
                get_recovery_serial=get_recovery_serial,
                abort_timeout=abort_timeout,
            )
            if sent and on_sent is not None:
                try:
                    on_sent()
                except Exception:
                    # Delivery is already confirmed. Do not report a send
                    # failure or retry it because a history callback failed.
                    if log_error is not None:
                        try:
                            log_error("短信已发送，但发送记录更新失败")
                        except Exception:
                            pass
            return sent
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
