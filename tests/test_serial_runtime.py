import unittest

from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.serial_runtime import (
    SerialLineDecoder,
    SerialRuntimeCallbacks,
    SerialRuntimeConfig,
    SerialRuntimeState,
    handle_serial_runtime_line,
    run_serial_runtime_thread,
    run_serial_thread_loop,
)


def _swap_number_digits(number):
    digits = str(number or "").lstrip("+")
    if len(digits) % 2:
        digits += "F"
    return "".join(digits[i + 1] + digits[i] for i in range(0, len(digits), 2))


def _incoming_ucs2_pdu(sender, message, *, reference=0x2A, total=1, index=1, timestamp_hex="62608211510523"):
    sender_text = str(sender or "")
    number_type = "91" if sender_text.startswith("+") else "81"
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
        + timestamp_hex
        + f"{len(user_data):02X}"
        + user_data.hex().upper()
    )


def parse_head(text):
    parts = str(text or "").split(" ", 1)
    return (parts[0], parts[1] if len(parts) > 1 else text)


def runtime_config(**overrides):
    values = {
        "keywords": [],
        "log_unmatched_sms": False,
        "log_dir": ".",
        "log_prefix": "COM5",
        "error_repeat_limit": 3,
        "call_filter_mode": "Disabled",
        "call_whitelist": [],
        "call_blacklist": [],
    }
    values.update(overrides)
    return SerialRuntimeConfig(**values)


def runtime_callbacks(calls):
    return SerialRuntimeCallbacks(
        enqueue_third_push=lambda *args, **kwargs: calls.append(("push", args, kwargs)),
        send_cloud_sms_event=lambda *args: calls.append(("cloud_sms", args)),
        port_ui=lambda *args: calls.append(("ui", args)),
        play_alert=lambda: calls.append(("alert",)),
        show_sms_popup=lambda *args: calls.append(("sms_popup", args)),
        file_log=lambda *args: calls.append(("file_log", args)),
        system_ui=lambda *args: calls.append(("system", args)),
        push_serial_debug=lambda *args: calls.append(("debug", args)),
        send_cloud_serial_log=lambda *args: calls.append(("cloud_log", args)),
        capture_cloud_device_imei=lambda *args: calls.append(("imei", args)),
        set_temperature=lambda *args: calls.append(("temp", args)),
        set_signal=lambda *args: calls.append(("signal", args)),
        set_status=lambda *args: calls.append(("status", args)),
        close_call_popup=lambda: calls.append(("close_popup",)),
        send_call_hangup=lambda: calls.append(("hangup",)),
        show_call_popup=lambda *args: calls.append(("call_popup", args)),
    )


def flush_settled_sms(state, calls, now, *, config=None, port="COM5"):
    handle_serial_runtime_line(
        state,
        "",
        now,
        port,
        False,
        config or runtime_config(),
        runtime_callbacks(calls),
        {},
    )


class SerialRuntimeTests(unittest.TestCase):
    def test_serial_line_decoder_keeps_split_utf8_character(self):
        decoder = SerialLineDecoder()
        text = "[I]-[handler_sms.smsCallback] +10086 您正在中国电信APP\r\n"
        raw = text.encode("utf-8")
        split_at = raw.index("中".encode("utf-8")) + 1

        self.assertEqual(decoder.feed(raw[:split_at]), [])
        self.assertEqual(decoder.feed(raw[split_at:]), [text.strip()])

    def test_serial_runtime_observes_sms_send_response_lines(self):
        calls = []
        state = SerialRuntimeState.create(parse_sms_callback_head)
        callbacks = runtime_callbacks(calls)
        callbacks = SerialRuntimeCallbacks(
            **{
                **callbacks.__dict__,
                "observe_sms_send_line": lambda line: calls.append(("sms_send_line", line)),
            }
        )

        handle_serial_runtime_line(
            state,
            "+CMGS: 12",
            1.0,
            "COM5",
            False,
            runtime_config(),
            callbacks,
            {},
        )

        self.assertIn(("sms_send_line", "+CMGS: 12"), calls)

    def test_serial_line_decoder_joins_newline_inserted_inside_utf8_character(self):
        decoder = SerialLineDecoder()
        text = "[I]-[handler_sms.smsCallback] +10086 您正在中国电信APP\r\n"
        raw = text.encode("utf-8")
        split_at = raw.index("中".encode("utf-8")) + 1

        self.assertEqual(decoder.feed(raw[:split_at] + b"\r\n"), [])
        self.assertEqual(decoder.feed(raw[split_at:]), [text.strip()])

    def test_serial_line_decoder_joins_three_byte_character_split_across_three_lines(self):
        decoder = SerialLineDecoder()
        prefix = "[I]-[handler_sms.smsCallback] 10699000 26/07/15,15:04:04+32 其中"
        character = "：".encode("utf-8")
        suffix = "中国电信2张；中国移动1张。\r\n".encode("utf-8")

        self.assertEqual(decoder.feed(prefix.encode("utf-8") + character[:1] + b"\r\n"), [])
        self.assertEqual(decoder.feed(character[1:2] + b"\r\n"), [])
        self.assertEqual(
            decoder.feed(character[2:] + suffix),
            [prefix + "：中国电信2张；中国移动1张。"],
        )

    def test_serial_line_decoder_preserves_real_newline_after_buffered_character_completes(self):
        decoder = SerialLineDecoder()
        character = "服".encode("utf-8")

        self.assertEqual(decoder.feed(b"customer " + character[:1] + b"\r\n"), [])
        self.assertEqual(decoder.feed(character[1:2] + b"\r\n"), [])
        self.assertEqual(decoder.feed(character[2:] + b" hotline\r\n"), ["customer 服 hotline"])

    def test_serial_line_decoder_keeps_timeout_from_flushing_partial_line(self):
        decoder = SerialLineDecoder()

        self.assertEqual(decoder.feed("您正在".encode("utf-8")), [])
        self.assertEqual(decoder.feed(b""), [])
        self.assertEqual(decoder.feed("中国电信APP\r\n".encode("utf-8")), ["您正在中国电信APP"])

    def test_serial_line_decoder_emits_modem_sms_prompt_without_newline(self):
        decoder = SerialLineDecoder()

        self.assertEqual(decoder.feed(b"> "), [">"])

    def test_serial_line_decoder_emits_prompt_when_prompt_arrives_in_chunks(self):
        decoder = SerialLineDecoder()

        self.assertEqual(decoder.feed(b"\r\n"), [""])
        self.assertEqual(decoder.feed(b"> "), [">"])

    def test_blank_line_flushes_expired_sms(self):
        calls = []
        state = SerialRuntimeState.create(parse_head)
        state.sms_collector.start("+8613123123123 hello code", now=10.0)

        result = handle_serial_runtime_line(
            state,
            "",
            20.0,
            "COM5",
            False,
            runtime_config(keywords=["hello"]),
            runtime_callbacks(calls),
            {},
        )

        self.assertTrue(result.continue_read)
        self.assertFalse(state.sms_collector.active)
        self.assertIn(("sms_popup", ("hello code",)), calls)
        self.assertIn(("alert",), calls)

    def test_ring_timeout_updates_status_and_closes_popup(self):
        calls = []
        state = SerialRuntimeState.create(parse_head)
        state.call_state.ring_timeout_target = 5.0
        state.call_state.last_clip_num = "10086"

        handle_serial_runtime_line(
            state,
            "",
            10.0,
            "COM5",
            True,
            runtime_config(),
            runtime_callbacks(calls),
            {},
        )

        self.assertEqual(state.call_state.ring_timeout_target, 0.0)
        self.assertEqual(state.call_state.last_clip_num, "")
        self.assertIn(("close_popup",), calls)
        self.assertTrue(any(item[0] == "status" for item in calls))

    def test_incoming_call_dispatches_push_status_and_popup(self):
        calls = []
        state = SerialRuntimeState.create(parse_head)

        result = handle_serial_runtime_line(
            state,
            '+CLIP: "+8613123123123",129',
            10.0,
            "COM5",
            False,
            runtime_config(),
            runtime_callbacks(calls),
            {},
        )

        self.assertFalse(result.continue_read)
        self.assertEqual(state.call_state.last_clip_num, "+8613123123123")
        self.assertEqual(state.call_state.ring_timeout_target, 22.0)
        self.assertIn(("call_popup", ("+8613123123123",)), calls)
        self.assertTrue(any(
            item[0] == "push"
            and item[2].get("event_type") == "call"
            and item[2].get("variables", {}).get("caller") == "+8613123123123"
            for item in calls
        ))

    def test_blocked_call_hangs_up_and_stops_processing(self):
        calls = []
        state = SerialRuntimeState.create(parse_head)

        result = handle_serial_runtime_line(
            state,
            '+CLIP: "+8613123123123",129',
            10.0,
            "COM5",
            False,
            runtime_config(call_filter_mode="Blacklist", call_blacklist=["13123123123"]),
            runtime_callbacks(calls),
            {},
        )

        self.assertTrue(result.continue_read)
        self.assertIn(("hangup",), calls)
        self.assertNotIn(("call_popup", ("+8613123123123",)), calls)

    def test_sms_callback_starts_collection(self):
        calls = []
        state = SerialRuntimeState.create(parse_head)

        result = handle_serial_runtime_line(
            state,
            "[I]-[handler_sms.smsCallback] +8613123123123 hello",
            10.0,
            "COM5",
            False,
            runtime_config(),
            runtime_callbacks(calls),
            {},
        )

        self.assertTrue(result.continue_read)
        self.assertTrue(state.sms_collector.active)
        self.assertEqual(state.sms_collector.callback_head, "+8613123123123 hello")

    def test_sms_callback_keeps_plus_digit_continuation_until_urc_boundary(self):
        calls = []
        state = SerialRuntimeState.create(parse_head)
        lines = [
            "[I]-[handler_sms.smsCallback] +8613123123123 first line",
            "+100.00 元到账",
            "+CMTI: \"SM\",1",
        ]

        for index, line in enumerate(lines):
            handle_serial_runtime_line(
                state,
                line,
                10.0 + index,
                "COM5",
                False,
                runtime_config(),
                runtime_callbacks(calls),
                {},
            )

        self.assertFalse(state.sms_collector.active)
        self.assertIn(("sms_popup", ("first line\n+100.00 元到账",)), calls)
        self.assertTrue(any(
            item[0] == "push" and item[1][0] == "first line\n+100.00 元到账"
            for item in calls
        ))

    def test_sms_callback_uses_cached_pdu_when_long_log_text_is_corrupted(self):
        calls = []
        state = SerialRuntimeState.create(parse_sms_callback_head)
        body = "中国电信温馨提醒:尊享来电识别【号码百事通】"
        part1 = body[:12]
        part2 = body[12:]
        lines = [
            "[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80",
            _incoming_ucs2_pdu("10086", part1, total=2, index=1),
            "[I]-[TP-PID : ] 0 dcs:  8",
            "[I]-[lib_sms rsp] +CMGR AT+CMGR=2 true OK +CMGR: 0,,80",
            _incoming_ucs2_pdu("10086", part2, total=2, index=2),
            "[I]-[TP-PID : ] 0 dcs:  8",
            "[I]-[handler_sms.smsCallback] 10086 26/06/28,11:15:50+32 中国电信温馨提醒:尊享来电�",
            "��别【号码百事通】",
            "[I]-[ril.proatc] OK",
        ]

        for index, line in enumerate(lines):
            handle_serial_runtime_line(
                state,
                line,
                10.0 + index,
                "COM5",
                False,
                runtime_config(),
                runtime_callbacks(calls),
                {},
            )
        flush_settled_sms(state, calls, 10.0 + len(lines) + 6.0)

        self.assertIn(("sms_popup", (body,)), calls)
        self.assertIn(("cloud_sms", ("10086 26/06/28,11:15:50+32 " + body, body)), calls)
        self.assertFalse(any(
            item[0] == "sms_popup" and "\ufffd" in item[1][0]
            for item in calls
        ))

    def test_sms_callback_assembles_cached_pdu_when_multipart_timestamps_match_assembler_window(self):
        calls = []
        state = SerialRuntimeState.create(parse_sms_callback_head)
        body = "中国电信温馨提醒:尊享来电识别【号码百事通】"
        part1 = body[:12]
        part2 = body[12:]
        lines = [
            "[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80",
            _incoming_ucs2_pdu("10086", part1, total=2, index=1),
            "[I]-[TP-PID : ] 0 dcs:  8",
            "[I]-[lib_sms rsp] +CMGR AT+CMGR=2 true OK +CMGR: 0,,80",
            _incoming_ucs2_pdu(
                "10086",
                part2,
                total=2,
                index=2,
                timestamp_hex="62608211511523",
            ),
            "[I]-[TP-PID : ] 0 dcs:  8",
            "[I]-[handler_sms.smsCallback] 10086 26/06/28,11:15:50+32 " + body[:18] + "\ufffd",
            "\ufffd\ufffd" + body[18:],
            "[I]-[ril.proatc] OK",
        ]

        for index, line in enumerate(lines):
            handle_serial_runtime_line(
                state,
                line,
                10.0 + index,
                "COM5",
                False,
                runtime_config(),
                runtime_callbacks(calls),
                {},
            )
        flush_settled_sms(state, calls, 10.0 + len(lines) + 6.0)

        self.assertIn(("sms_popup", (body,)), calls)
        self.assertIn(("cloud_sms", ("10086 26/06/28,11:15:50+32 " + body, body)), calls)

    def test_concat_progress_logs_stay_out_of_user_facing_logs(self):
        calls = []
        state = SerialRuntimeState.create(parse_sms_callback_head)
        callbacks = runtime_callbacks(calls)
        body = "截止到2026年06月29日12时21分，您的费用情况如下：更多流量回复9"
        part1 = body[:24]
        part2 = body[24:]
        lines = [
            "[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80",
            _incoming_ucs2_pdu("10001", part1, reference=0x8A, total=2, index=1),
            "[I]-[TP-PID : ] 0 dcs:  8",
            "[I]-[handler_sms.smsCallback] 10001 26/06/28,11:15:50+32 " + part1,
            "[I]-[ril.proatc] OK",
            "[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80",
            _incoming_ucs2_pdu("10001", part2, reference=0x8A, total=2, index=2),
            "[I]-[TP-PID : ] 0 dcs:  8",
            "[I]-[handler_sms.smsCallback] 10001 26/06/28,11:15:50+32 " + part2,
            "[I]-[ril.proatc] OK",
        ]

        for index, line in enumerate(lines):
            handle_serial_runtime_line(
                state,
                line,
                10.0 + index,
                "COM60",
                False,
                runtime_config(log_dir="logs", log_prefix="COM60"),
                callbacks,
                {},
            )
        flush_settled_sms(
            state,
            calls,
            10.0 + len(lines) + 6.0,
            config=runtime_config(log_dir="logs", log_prefix="COM60"),
            port="COM60",
        )

        self.assertFalse(any(
            item[0] == "system" and "SMS CONCAT" in str(item[1])
            for item in calls
        ))
        self.assertTrue(any(
            item[0] == "debug" and "SMS CONCAT" in str(item[1])
            for item in calls
        ))
        self.assertFalse(any(
            item[0] == "file_log" and "SMS CONCAT" in str(item[1])
            for item in calls
        ))
        self.assertIn(("sms_popup", (body,)), calls)

    def test_merged_multiline_lua_callback_is_shown_after_concat_pdu_parts(self):
        calls = []
        state = SerialRuntimeState.create(parse_sms_callback_head)
        callbacks = runtime_callbacks(calls)
        body = (
            "截止到2026年06月29日 12时33分，您的费用情况如下：\n"
            "1、本月已产生话费5.20元，通用话费余额：352.66元，本月已充值8.16元\n"
            "2、套餐使用情况：\n"
            "通用流量：本月总量400MB（其中上月结转200MB），已使用5.38MB，剩余394.62MB\n"
            "3、更多查询方式：\n"
            "1）关注“湖南电信”微信公众号\n"
            "2）登录官方链接 t.hn.189.cn/RVVRFzuu\n"
            "3）拨打查费专线1000111\n"
            "4）更多流量回复9"
        )
        parts = [body[:70], body[70:140], body[140:210], body[210:]]
        lines = []
        for index, part in enumerate(parts, start=1):
            lines.extend([
                f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,156",
                _incoming_ucs2_pdu(
                    "10001",
                    part,
                    reference=0x50,
                    total=4,
                    index=index,
                    timestamp_hex="62609221332523",
                ),
                "[I]-[TP-PID : ] 0 dcs:  8",
            ])
        callback_lines = body.split("\n")
        lines.extend([
            "[I]-[handler_sms.smsCallback] 10001 26/06/29,12:33:52+32 " + callback_lines[0],
            *callback_lines[1:],
            "[I]-[usbmsc.write] usb storage free size: 12288/82432B",
        ])

        for index, line in enumerate(lines):
            handle_serial_runtime_line(
                state,
                line,
                10.0 + index,
                "COM60",
                False,
                runtime_config(log_dir="logs", log_prefix="COM60"),
                callbacks,
                {},
            )
        flush_settled_sms(
            state,
            calls,
            10.0 + len(lines) + 6.0,
            config=runtime_config(log_dir="logs", log_prefix="COM60"),
            port="COM60",
        )

        self.assertTrue(any(
            item[0] == "sms_popup"
            and "截止到2026年06月29日" in item[1][0]
            and "4）更多流量回复9" in item[1][0]
            for item in calls
        ))
        self.assertFalse(state.sms_collector.active)
        self.assertEqual(state.long_sms_assembler._pending, {})

    def test_sms_callback_collects_pdu_split_at_odd_hex_line_length(self):
        calls = []
        state = SerialRuntimeState.create(parse_sms_callback_head)
        body = "中国电信温馨提醒:尊享来电识别【号码百事通】"
        pdu = _incoming_ucs2_pdu("10086", body)
        lines = [
            "[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80",
            pdu[:127],
            pdu[127:],
            "[I]-[TP-PID : ] 0 dcs:  8",
            "[I]-[handler_sms.smsCallback] 10086 26/06/28,11:15:50+32 中国电信温馨提醒:尊享来电�",
            "[I]-[ril.proatc] OK",
        ]

        for index, line in enumerate(lines):
            handle_serial_runtime_line(
                state,
                line,
                10.0 + index,
                "COM5",
                False,
                runtime_config(),
                runtime_callbacks(calls),
                {},
            )

        self.assertIn(("sms_popup", (body,)), calls)

    def test_run_serial_thread_loop_reads_lines_until_stopped(self):
        calls = []
        keep_running = [True, True, False, False]
        raw_lines = [b"first\r\n", b"second\r\n"]

        run_serial_thread_loop(
            should_continue=lambda: keep_running.pop(0),
            get_target_port=lambda: "COM5",
            resolve_target_port=lambda: "COM5",
            set_connecting_status=lambda port: calls.append(("connecting", port)),
            open_and_initialize_serial=lambda port: calls.append(("open", port)),
            on_connected_port=lambda port: calls.append(("connected", port)),
            read_serial_line=lambda: raw_lines.pop(0),
            handle_line=lambda line: calls.append(("line", line)),
            handle_error=lambda error, port: calls.append(("error", str(error), port)) or False,
            wait_before_retry=lambda: calls.append(("wait",)),
            safe_close_serial=lambda: calls.append(("close",)),
        )

        self.assertEqual(calls, [
            ("connecting", "COM5"),
            ("open", "COM5"),
            ("connected", "COM5"),
            ("line", "first"),
            ("close",),
        ])

    def test_run_serial_thread_loop_combines_split_utf8_line(self):
        calls = []
        text = "[I]-[handler_sms.smsCallback] +10086 您正在中国电信APP\r\n"
        raw = text.encode("utf-8")
        split_at = raw.index("中".encode("utf-8")) + 1
        raw_lines = [raw[:split_at], raw[split_at:]]
        keep_running = [True, True, True, False, False]

        run_serial_thread_loop(
            should_continue=lambda: keep_running.pop(0),
            get_target_port=lambda: "COM5",
            resolve_target_port=lambda: "COM5",
            set_connecting_status=lambda port: calls.append(("connecting", port)),
            open_and_initialize_serial=lambda port: calls.append(("open", port)),
            on_connected_port=lambda port: calls.append(("connected", port)),
            read_serial_line=lambda: raw_lines.pop(0),
            handle_line=lambda line: calls.append(("line", line)),
            handle_error=lambda error, port: calls.append(("error", str(error), port)) or False,
            wait_before_retry=lambda: calls.append(("wait",)),
            safe_close_serial=lambda: calls.append(("close",)),
        )

        self.assertIn(("line", text.strip()), calls)
        self.assertNotIn(("line", "[I]-[handler_sms.smsCallback] +10086 您正在"), calls)

    def test_run_serial_thread_loop_combines_utf8_split_by_inserted_newline(self):
        calls = []
        text = "[I]-[handler_sms.smsCallback] +10086 您正在中国电信APP\r\n"
        raw = text.encode("utf-8")
        split_at = raw.index("中".encode("utf-8")) + 1
        raw_lines = [raw[:split_at] + b"\r\n", raw[split_at:]]
        keep_running = [True, True, True, False, False]

        run_serial_thread_loop(
            should_continue=lambda: keep_running.pop(0),
            get_target_port=lambda: "COM5",
            resolve_target_port=lambda: "COM5",
            set_connecting_status=lambda port: calls.append(("connecting", port)),
            open_and_initialize_serial=lambda port: calls.append(("open", port)),
            on_connected_port=lambda port: calls.append(("connected", port)),
            read_serial_line=lambda: raw_lines.pop(0),
            handle_line=lambda line: calls.append(("line", line)),
            handle_error=lambda error, port: calls.append(("error", str(error), port)) or False,
            wait_before_retry=lambda: calls.append(("wait",)),
            safe_close_serial=lambda: calls.append(("close",)),
        )

        self.assertIn(("line", text.strip()), calls)
        self.assertFalse(any("�" in item[1] for item in calls if item[0] == "line"))

    def test_run_serial_thread_loop_skips_missing_target_port(self):
        calls = []
        keep_running = [True, False]
        times = [1.0, 1.001]

        run_serial_thread_loop(
            should_continue=lambda: keep_running.pop(0),
            get_target_port=lambda: "COM5",
            resolve_target_port=lambda: None,
            set_connecting_status=lambda port: calls.append(("connecting", port)),
            open_and_initialize_serial=lambda port: calls.append(("open", port)),
            on_connected_port=lambda port: calls.append(("connected", port)),
            read_serial_line=lambda: b"",
            handle_line=lambda line: calls.append(("line", line)),
            handle_error=lambda error, port: False,
            wait_before_retry=lambda: calls.append(("wait",)),
            safe_close_serial=lambda: calls.append(("close",)),
            monotonic=lambda: times.pop(0),
        )

        self.assertEqual(calls, [("wait",), ("close",)])

    def test_run_serial_thread_loop_does_not_double_wait_when_resolver_waited(self):
        calls = []
        keep_running = [True, False]
        times = [1.0, 2.0]

        run_serial_thread_loop(
            should_continue=lambda: keep_running.pop(0),
            get_target_port=lambda: "COM5",
            resolve_target_port=lambda: None,
            set_connecting_status=lambda port: calls.append(("connecting", port)),
            open_and_initialize_serial=lambda port: calls.append(("open", port)),
            on_connected_port=lambda port: calls.append(("connected", port)),
            read_serial_line=lambda: b"",
            handle_line=lambda line: calls.append(("line", line)),
            handle_error=lambda error, port: False,
            wait_before_retry=lambda: calls.append(("wait",)),
            safe_close_serial=lambda: calls.append(("close",)),
            monotonic=lambda: times.pop(0),
        )

        self.assertEqual(calls, [("close",)])

    def test_run_serial_thread_loop_waits_after_unhandled_error(self):
        calls = []
        keep_running = [True, False]

        run_serial_thread_loop(
            should_continue=lambda: keep_running.pop(0),
            get_target_port=lambda: "COM5",
            resolve_target_port=lambda: "COM5",
            set_connecting_status=lambda port: calls.append(("connecting", port)),
            open_and_initialize_serial=lambda port: (_ for _ in ()).throw(RuntimeError("down")),
            on_connected_port=lambda port: calls.append(("connected", port)),
            read_serial_line=lambda: b"",
            handle_line=lambda line: calls.append(("line", line)),
            handle_error=lambda error, port: calls.append(("error", str(error), port)) or False,
            wait_before_retry=lambda: calls.append(("wait",)),
            safe_close_serial=lambda: calls.append(("close",)),
        )

        self.assertEqual(calls, [
            ("connecting", "COM5"),
            ("error", "down", "COM5"),
            ("wait",),
            ("close",),
        ])

    def test_run_serial_thread_loop_skips_wait_after_handled_error(self):
        calls = []
        keep_running = [True, False]

        run_serial_thread_loop(
            should_continue=lambda: keep_running.pop(0),
            get_target_port=lambda: "COM5",
            resolve_target_port=lambda: "COM5",
            set_connecting_status=lambda port: calls.append(("connecting", port)),
            open_and_initialize_serial=lambda port: (_ for _ in ()).throw(RuntimeError("gone")),
            on_connected_port=lambda port: calls.append(("connected", port)),
            read_serial_line=lambda: b"",
            handle_line=lambda line: calls.append(("line", line)),
            handle_error=lambda error, port: calls.append(("error", str(error), port)) or True,
            wait_before_retry=lambda: calls.append(("wait",)),
            safe_close_serial=lambda: calls.append(("close",)),
        )

        self.assertEqual(calls, [
            ("connecting", "COM5"),
            ("error", "gone", "COM5"),
            ("close",),
        ])

    def test_run_serial_thread_loop_suppresses_disconnect_log_while_stopping(self):
        calls = []
        keep_running = [True, False]

        run_serial_thread_loop(
            should_continue=lambda: keep_running.pop(0),
            get_target_port=lambda: "COM5",
            resolve_target_port=lambda: "COM5",
            set_connecting_status=lambda port: calls.append(("connecting", port)),
            open_and_initialize_serial=lambda port: (_ for _ in ()).throw(RuntimeError("closed")),
            on_connected_port=lambda port: calls.append(("connected", port)),
            read_serial_line=lambda: b"",
            handle_line=lambda line: calls.append(("line", line)),
            handle_error=lambda error, port: calls.append(("error", str(error), port)) or False,
            wait_before_retry=lambda: calls.append(("wait",)),
            safe_close_serial=lambda: calls.append(("close",)),
            is_stopping=lambda: True,
        )

        self.assertEqual(calls, [("connecting", "COM5"), ("close",)])

    def test_run_serial_thread_loop_closes_serial_when_error_handler_raises(self):
        calls = []

        with self.assertRaisesRegex(RuntimeError, "handler failed"):
            run_serial_thread_loop(
                should_continue=lambda: True,
                get_target_port=lambda: "COM5",
                resolve_target_port=lambda: "COM5",
                set_connecting_status=lambda port: calls.append(("connecting", port)),
                open_and_initialize_serial=lambda port: (_ for _ in ()).throw(RuntimeError("down")),
                on_connected_port=lambda port: calls.append(("connected", port)),
                read_serial_line=lambda: b"",
                handle_line=lambda line: calls.append(("line", line)),
                handle_error=lambda error, port: (_ for _ in ()).throw(RuntimeError("handler failed")),
                wait_before_retry=lambda: calls.append(("wait",)),
                safe_close_serial=lambda: calls.append(("close",)),
            )

        self.assertEqual(calls, [("connecting", "COM5"), ("close",)])

    def test_run_serial_runtime_thread_syncs_call_state_after_line(self):
        calls = []
        app_state = [(7.0, "10086")]

        def fake_loop(**kwargs):
            kwargs["handle_line"]("RING")

        def fake_handle_line(state, line, now, current_port, popup_active, config, callbacks, ignore_repeat_state):
            calls.append((line, now, current_port, popup_active, config.log_prefix, ignore_repeat_state))
            self.assertEqual(state.call_state.ring_timeout_target, 0.0)
            self.assertEqual(state.call_state.current_dial_num, "")
            state.call_state.ring_timeout_target = 9.0
            state.call_state.current_dial_num = "10010"

        run_serial_runtime_thread(
            parse_callback_head=parse_head,
            get_runtime_config=lambda: runtime_config(log_prefix="COM9"),
            callbacks=runtime_callbacks([]),
            get_call_state=lambda: app_state[-1],
            set_call_state=lambda ring_timeout, dial_num: app_state.append((ring_timeout, dial_num)),
            popup_active=lambda: True,
            ignore_repeat_state={"seen": 1},
            should_continue=lambda: True,
            get_target_port=lambda: "COM9",
            resolve_target_port=lambda: "COM9",
            set_connecting_status=lambda *_: None,
            open_and_initialize_serial=lambda *_: None,
            on_connected_port=lambda *_: None,
            read_serial_line=lambda: b"",
            handle_disconnect=lambda *_: False,
            wait_before_retry=lambda: None,
            safe_close_serial=lambda: None,
            clock=lambda: 123.0,
            run_loop=fake_loop,
            handle_runtime_line=fake_handle_line,
        )

        self.assertEqual(app_state, [(7.0, "10086"), (0.0, ""), (9.0, "10010")])
        self.assertEqual(calls, [("RING", 123.0, "COM9", True, "COM9", {"seen": 1})])

    def test_run_serial_runtime_thread_refreshes_config_after_connect(self):
        app_state = [(0.0, "")]
        config_prefix = ["system"]
        log_unmatched = [True]
        seen_prefixes = []
        seen_log_flags = []

        def fake_loop(**kwargs):
            kwargs["on_connected_port"]("COM8")
            kwargs["handle_line"]("SMS")
            log_unmatched[0] = False
            kwargs["handle_line"]("SMS2")

        def fake_handle_line(state, line, now, current_port, popup_active, config, callbacks, ignore_repeat_state):
            seen_prefixes.append(config.log_prefix)
            seen_log_flags.append(config.log_unmatched_sms)

        run_serial_runtime_thread(
            parse_callback_head=parse_head,
            get_runtime_config=lambda: runtime_config(
                log_prefix=config_prefix[0],
                log_unmatched_sms=log_unmatched[0],
            ),
            callbacks=runtime_callbacks([]),
            get_call_state=lambda: app_state[-1],
            set_call_state=lambda ring_timeout, dial_num: app_state.append((ring_timeout, dial_num)),
            popup_active=lambda: False,
            ignore_repeat_state={},
            should_continue=lambda: True,
            get_target_port=lambda: "COM8",
            resolve_target_port=lambda: "COM8",
            set_connecting_status=lambda *_: None,
            open_and_initialize_serial=lambda *_: None,
            on_connected_port=lambda port: config_prefix.__setitem__(0, port),
            read_serial_line=lambda: b"",
            handle_disconnect=lambda *_: False,
            wait_before_retry=lambda: None,
            safe_close_serial=lambda: None,
            run_loop=fake_loop,
            handle_runtime_line=fake_handle_line,
        )

        self.assertEqual(seen_prefixes, ["COM8", "COM8"])
        self.assertEqual(seen_log_flags, [True, False])

    def test_run_serial_runtime_thread_resets_call_state_before_disconnect_callback(self):
        app_state = []
        disconnects = []

        def fake_loop(**kwargs):
            kwargs["handle_error"](RuntimeError("gone"), "COM5")

        run_serial_runtime_thread(
            parse_callback_head=parse_head,
            get_runtime_config=runtime_config,
            callbacks=runtime_callbacks([]),
            get_call_state=lambda: (8.0, "10086"),
            set_call_state=lambda ring_timeout, dial_num: app_state.append((ring_timeout, dial_num)),
            popup_active=lambda: False,
            ignore_repeat_state={},
            should_continue=lambda: True,
            get_target_port=lambda: "COM5",
            resolve_target_port=lambda: "COM5",
            set_connecting_status=lambda *_: None,
            open_and_initialize_serial=lambda *_: None,
            on_connected_port=lambda *_: None,
            read_serial_line=lambda: b"",
            handle_disconnect=lambda error, port: disconnects.append((str(error), port)) or True,
            wait_before_retry=lambda: None,
            safe_close_serial=lambda: None,
            run_loop=fake_loop,
        )

        self.assertEqual(app_state, [(0.0, ""), (0.0, "")])
        self.assertEqual(disconnects, [("gone", "COM5")])

    def test_reset_sms_state_clears_collector_pdu_cache_and_assembler(self):
        state = SerialRuntimeState.create(parse_head)
        state.sms_collector.start("10086 body", 1.0)
        state.sms_pdu_cache._collecting = True
        state.sms_pdu_cache._pdu_lines.append("0011")
        state.sms_pdu_cache._segments_by_key[("10086", 8, 42, 2)] = [object()]
        state.long_sms_assembler._pending[("10086", 8, 42, 2, 0)] = {"parts": {1: "old"}}
        state.long_sms_assembler._completed[("10086", 8, 42, 2, 0)] = {"parts": {1: "old"}}

        state.reset_sms_state()

        self.assertFalse(state.sms_collector.active)
        self.assertFalse(state.sms_pdu_cache._collecting)
        self.assertEqual(state.sms_pdu_cache._pdu_lines, [])
        self.assertEqual(state.sms_pdu_cache._segments_by_key, {})
        self.assertEqual(state.long_sms_assembler._pending, {})
        self.assertEqual(state.long_sms_assembler._completed, {})


if __name__ == "__main__":
    unittest.main()
