#!/usr/bin/env python
"""Replay or synthesize Air724UG SMS serial logs through the desktop parser."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
SMS_SRC = ROOT / "sms"
if str(SMS_SRC) not in sys.path:
    sys.path.insert(0, str(SMS_SRC))

from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.serial_runtime import (
    SerialRuntimeCallbacks,
    SerialRuntimeConfig,
    SerialRuntimeState,
    handle_serial_runtime_line,
)


def _swap_number_digits(number: str) -> str:
    digits = str(number or "").lstrip("+")
    if len(digits) % 2:
        digits += "F"
    return "".join(digits[i + 1] + digits[i] for i in range(0, len(digits), 2))


def _encode_timestamp(timestamp: str) -> str:
    text = str(timestamp or "26/06/28,11:15:50+32")
    date_part, time_part = text.split(",", 1)
    year, month, day = date_part.split("/")
    hour, minute, second_tz = time_part.split(":")
    if "+" in second_tz:
        second, tz = second_tz.split("+", 1)
    else:
        second, tz = second_tz, "32"
    fields = [year, month, day, hour, minute, second, tz]
    return "".join(item[1] + item[0] for item in fields)


def _split_ucs2_segments(message: str, max_bytes: int = 134):
    segments = []
    current = []
    current_len = 0
    for char in str(message or ""):
        encoded_len = len(char.encode("utf-16-be"))
        if current and current_len + encoded_len > max_bytes:
            segments.append("".join(current))
            current = []
            current_len = 0
        current.append(char)
        current_len += encoded_len
    if current or not segments:
        segments.append("".join(current))
    return segments


def _incoming_ucs2_pdu(
    sender: str,
    message: str,
    timestamp: str,
    *,
    reference=0x2A,
    total=1,
    index=1,
    number_type=None,
):
    sender_text = str(sender or "")
    number_type = number_type or ("91" if sender_text.startswith("+") else "81")
    sender_digits = sender_text.lstrip("+")
    first_octet = "40" if total > 1 else "00"
    payload = str(message or "").encode("utf-16-be")
    user_data = payload
    if total > 1:
        user_data = bytes((0x05, 0x00, 0x03, reference & 0xFF, total & 0xFF, index & 0xFF)) + payload
    return (
        "00"
        + first_octet
        + f"{len(sender_digits):02X}"
        + number_type
        + _swap_number_digits(sender_digits)
        + "00"
        + "08"
        + _encode_timestamp(timestamp)
        + f"{len(user_data):02X}"
        + user_data.hex().upper()
    )


def _wrap_hex(text: str, width: int = 127):
    return [text[i:i + width] for i in range(0, len(text), width)]


def synthesize_sms_log(sender: str, timestamp: str, message: str, corrupt_callback: bool = True):
    segments = _split_ucs2_segments(message)
    total = len(segments)
    lines = []
    for index, segment in enumerate(segments, start=1):
        pdu = _incoming_ucs2_pdu(sender, segment, timestamp, total=total, index=index)
        lines.append(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,{(len(pdu) - 2) // 2}")
        lines.extend(_wrap_hex(pdu))
        lines.append("[I]-[TP-PID : ] 0 dcs:  8")

    callback_body = message
    continuation = ""
    if corrupt_callback and len(message) > 20:
        split_at = max(8, min(len(message) - 1, len(message) // 3))
        callback_body = message[:split_at] + "\ufffd"
        continuation = "\ufffd\ufffd" + message[split_at + 1:]

    lines.append(f"[I]-[handler_sms.smsCallback] {sender} {timestamp} {callback_body}")
    if continuation:
        lines.append(continuation)
    lines.append("[I]-[ril.proatc] OK")
    return lines


def replay_lines(lines):
    calls = []
    state = SerialRuntimeState.create(parse_sms_callback_head)
    config = SerialRuntimeConfig(
        keywords=[],
        log_unmatched_sms=False,
        log_dir=".",
        log_prefix="SIM",
        error_repeat_limit=3,
        call_filter_mode="Disabled",
        call_whitelist=[],
        call_blacklist=[],
    )
    callbacks = SerialRuntimeCallbacks(
        enqueue_third_push=lambda *args, **kwargs: calls.append(("push", args, kwargs)),
        send_cloud_sms_event=lambda *args: calls.append(("cloud_sms", args, {})),
        port_ui=lambda *args: calls.append(("ui", args, {})),
        play_alert=lambda: calls.append(("alert", (), {})),
        show_sms_popup=lambda *args: calls.append(("sms_popup", args, {})),
        file_log=lambda *args: calls.append(("file_log", args, {})),
        system_ui=lambda *args: calls.append(("system", args, {})),
        push_serial_debug=lambda *args: None,
        send_cloud_serial_log=lambda *args: None,
        capture_cloud_device_imei=lambda *args: None,
        set_temperature=lambda *args: None,
        set_signal=lambda *args: None,
        set_status=lambda *args: None,
        close_call_popup=lambda: None,
        send_call_hangup=lambda: None,
        show_call_popup=lambda *args: None,
        set_local_number=lambda *args: None,
    )
    repeat_state = {}

    for index, line in enumerate(lines):
        handle_serial_runtime_line(
            state,
            line,
            10.0 + index * 0.01,
            "SIM",
            False,
            config,
            callbacks,
            repeat_state,
        )
    handle_serial_runtime_line(
        state,
        "",
        10.0 + len(lines) * 0.01 + 5.5,
        "SIM",
        False,
        config,
        callbacks,
        repeat_state,
    )
    return calls


def main(argv=None):
    parser = argparse.ArgumentParser(description="Replay or synthesize SMS serial logs.")
    parser.add_argument("--log", help="Replay an existing serial log text file.")
    parser.add_argument("--sender", default="10086", help="Sender number for synthetic SMS.")
    parser.add_argument("--time", default="26/06/28,11:15:50+32", help="SMS timestamp, e.g. 26/06/28,11:15:50+32.")
    parser.add_argument("--message", help="Synthetic SMS body.")
    parser.add_argument("--no-corrupt", action="store_true", help="Do not corrupt the synthetic callback line.")
    parser.add_argument("--show-lines", action="store_true", help="Print generated/replayed serial lines.")
    args = parser.parse_args(argv)

    if args.log:
        lines = Path(args.log).read_text(encoding="utf-8", errors="replace").splitlines()
    else:
        message = args.message or (
            "中国电信温馨提醒:您刚才有漏接来电，回复\"1\"查看漏话，"
            "尊享来电识别、骚扰电话自动应答、漏接来电即时提醒。【号码百事通】"
        )
        lines = synthesize_sms_log(args.sender, args.time, message, corrupt_callback=not args.no_corrupt)

    if args.show_lines:
        print("=== SERIAL LINES ===")
        for line in lines:
            print(line)

    calls = replay_lines(lines)
    sms_popups = [args[0] for name, args, _kwargs in calls if name == "sms_popup"]
    cloud_events = [args for name, args, _kwargs in calls if name == "cloud_sms"]

    print("=== RESULT ===")
    print(f"sms_popup_count={len(sms_popups)}")
    print(f"cloud_sms_count={len(cloud_events)}")
    print(f"has_replacement={any(chr(0xFFFD) in item for item in sms_popups)}")
    for index, body in enumerate(sms_popups, start=1):
        print(f"popup[{index}]={body}")
    for index, event in enumerate(cloud_events, start=1):
        head, body = event
        print(f"cloud[{index}].head={head}")
        print(f"cloud[{index}].body={body}")


if __name__ == "__main__":
    main()
