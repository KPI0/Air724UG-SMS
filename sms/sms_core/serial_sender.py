from dataclasses import dataclass
import time

from sms_core.serial_debug import build_serial_command_payload
from sms_core.sms_pdu import encode_text_sms_pdu
from sms_core.threading_runtime import start_daemon_thread


@dataclass(frozen=True)
class SerialCommandResult:
    ok: bool
    error: str = ""


def _is_serial_open(serial_obj):
    return serial_obj is not None and bool(getattr(serial_obj, "is_open", False))


def write_serial_command_result(serial_obj, command, append_crlf=True, push_debug=None):
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
):
    if not _is_serial_open(serial_obj):
        if push_debug:
            push_debug(">>> 发送失败: 串口未连接")
        if port_ui:
            port_ui("❌ 发送短信失败：串口未连接", "normal")
        return False

    try:
        pdu_str, cmgs_len = encode_text_sms_pdu(phone, message)
        if port_ui:
            port_ui(f"📤 发送短信至 {phone}：", "normal")
            port_ui(message, "sms")

        command = "AT+CMGF=0"
        serial_obj.write((command + "\r\n").encode("utf-8"))
        serial_obj.flush()
        if push_debug:
            push_debug(f">>> 发送: {command}\\r\\n")
        sleep_func(0.3)

        command = f"AT+CMGS={cmgs_len}"
        serial_obj.write((command + "\r\n").encode("utf-8"))
        serial_obj.flush()
        if push_debug:
            push_debug(f">>> 发送: {command}\\r\\n")
        sleep_func(1.0)

        serial_obj.write(pdu_str.encode("utf-8") + b"\x1a")
        serial_obj.flush()
        if push_debug:
            push_debug(">>> 发送 PDU 正文及 Ctrl+Z，等待模组响应...")
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


def send_command_sequence_async(serial_lock, get_serial, commands, push_debug=None, delay_sec=0.3, log_error=None):
    return start_daemon_thread(
        "serial_send_sequence",
        lambda: _run_with_serial(
            serial_lock,
            get_serial,
            lambda serial_obj: write_serial_command_sequence(
                serial_obj,
                commands,
                push_debug=push_debug,
                delay_sec=delay_sec,
            ),
        ),
        log_error=log_error,
    )


def send_text_sms_pdu_async(serial_lock, get_serial, phone, message, push_debug=None, port_ui=None, log_error=None):
    return start_daemon_thread(
        "serial_send_sms_pdu",
        lambda: _run_with_serial(
            serial_lock,
            get_serial,
            lambda serial_obj: write_text_sms_pdu(
                serial_obj,
                phone,
                message,
                push_debug=push_debug,
                port_ui=port_ui,
            ),
        ),
        log_error=log_error,
    )
