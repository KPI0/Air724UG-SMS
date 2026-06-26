import os
import tempfile
import unittest
from datetime import datetime

from sms_core.cloud_security import (
    prune_replay_cache,
    read_unix_timestamp,
    replay_key,
    safe_preview,
)
from sms_core.log_cleanup import cleanup_old_logs_in_dir, parse_date_from_log_filename
from sms_core.phone_numbers import normalize_call_number
from sms_core.serial_debug import (
    COMMON_SERIAL_COMMANDS,
    HANGUP_COMMAND,
    SERIAL_DEBUG_MAX_STORE_LINES,
    SERIAL_DEBUG_MAX_VISIBLE_LINES,
    build_dial_command,
    build_own_number_commands,
    build_pin_change_command,
    build_pin_lock_command,
    build_pin_unlock_command,
    build_puk_unlock_command,
    build_serial_command_payload,
    build_sn_command,
    normalize_dial_number,
    normalize_own_number,
    quick_command_label,
)
from sms_core.serial_parsers import (
    evaluate_call_filter,
    is_call_connected_event,
    is_hangup_event,
    is_new_clip,
    is_ring_line,
    is_sms_collection_boundary,
    parse_clip_number,
    parse_cesq_rsrp,
    parse_serial_debug_insights,
    parse_temperature,
)
from sms_core.serial_sender import (
    write_serial_command,
    write_serial_command_sequence,
    write_text_sms_pdu,
)
from sms_core.serial_sms import SMS_CALLBACK_PREFIX, SmsPendingCollector, callback_body_from_line
from sms_core.sms_processing import (
    build_unmatched_sms_log_entries,
    process_pending_sms,
    repeat_count_message,
    sms_keyword_hit,
)
from sms_core.sms_pdu import encode_text_sms_pdu
from sms_core.app_launch import (
    decode_restart_args,
    encode_restart_args,
    get_clean_restart_env,
    get_launch_target_and_args,
)
from sms_core.updates import (
    build_release_api_urls,
    format_proxy_test_result,
    normalize_proxy_base,
    pick_zip_asset,
    test_update_proxy_connectivity,
)
from sms_core.windows_shortcuts import sanitize_shortcut_name
from sms_core.windows_runtime import (
    is_existing_instance_error,
    normalize_serial_device_path,
    port_mutex_name,
)


class CoreHelperTests(unittest.TestCase):
    class FakeSerial:
        def __init__(self, is_open=True):
            self.is_open = is_open
            self.writes = []
            self.flush_count = 0

        def write(self, payload):
            self.writes.append(payload)

        def flush(self):
            self.flush_count += 1

    def test_normalize_call_number_strips_china_prefix(self):
        self.assertEqual(normalize_call_number("+8613812345678"), "13812345678")
        self.assertEqual(normalize_call_number("8613812345678"), "13812345678")
        self.assertEqual(normalize_call_number("10086"), "10086")

    def test_encode_text_sms_pdu_uses_ucs2_and_cmgs_tpdu_length(self):
        pdu, cmgs_len = encode_text_sms_pdu("+8613812345678", "\u6d4b\u8bd5")

        self.assertTrue(pdu.startswith("0011000D91"))
        self.assertIn("6D4B8BD5", pdu)
        self.assertEqual(cmgs_len, (len(pdu) - 2) // 2)

    def test_parse_date_from_log_filename(self):
        self.assertEqual(
            parse_date_from_log_filename("sms_COM5_2026-06-08.txt").isoformat(),
            "2026-06-08",
        )
        self.assertIsNone(parse_date_from_log_filename("README.md"))

    def test_cleanup_old_logs_in_dir(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_path = os.path.join(tmp, "sms_COM5_2026-05-01.txt")
            keep_path = os.path.join(tmp, "sms_COM5_2026-06-07.txt")
            other_path = os.path.join(tmp, "other_2026-05-01.txt")

            for path in (old_path, keep_path, other_path):
                with open(path, "w", encoding="utf-8") as f:
                    f.write("x")

            deleted = cleanup_old_logs_in_dir(tmp, 30, now=datetime(2026, 6, 8))

            self.assertEqual(deleted, 1)
            self.assertFalse(os.path.exists(old_path))
            self.assertTrue(os.path.exists(keep_path))
            self.assertTrue(os.path.exists(other_path))

    def test_safe_preview_masks_secret_fields(self):
        preview = safe_preview('{"secret":"abc","nested":{"token":"def","value":1}}')

        self.assertIn('"secret": "***"', preview)
        self.assertIn('"token": "***"', preview)
        self.assertNotIn("abc", preview)
        self.assertNotIn("def", preview)

    def test_read_unix_timestamp_accepts_milliseconds(self):
        self.assertEqual(
            read_unix_timestamp({"timestamp": "1710000000123"}),
            (1710000000, "1710000000123"),
        )
        self.assertEqual(read_unix_timestamp({"ts": "1710000000"}), (1710000000, "1710000000"))
        self.assertEqual(read_unix_timestamp({}), (None, None))

    def test_replay_key_prefers_nonce_and_hashes_without_nonce(self):
        self.assertEqual(replay_key({"nonce": "abc"}, 1710000000), "nonce:abc")
        key = replay_key({"type": "cmd", "target_imei": "123", "command": "AT"}, 1710000000)
        self.assertTrue(key.startswith("fp:"))
        self.assertEqual(len(key), 67)

    def test_prune_replay_cache_removes_expired_and_oldest_entries(self):
        cache = {"old": 10, "keep": 95, "extra": 96}

        prune_replay_cache(cache, now_ts=100, window_seconds=20, max_size=1)

        self.assertEqual(cache, {"extra": 96})

    def test_update_helpers(self):
        self.assertEqual(normalize_proxy_base("gh-proxy.com"), "https://gh-proxy.com/")
        self.assertEqual(
            build_release_api_urls("KPI0", "Air724UG-SMS", "proxy.example", now=123)[0],
            "https://proxy.example/repos/KPI0/Air724UG-SMS/releases/latest?t=123",
        )

        asset = pick_zip_asset({
            "assets": [
                {"name": "sms-small.zip", "size": 10},
                {"name": "sms-large.zip", "size": 20},
                {"name": "sms.exe", "size": 100},
            ]
        })
        self.assertEqual(asset["name"], "sms-large.zip")

    def test_update_proxy_connectivity_formats_successes(self):
        calls = []

        def fake_get_json(url, timeout=0, retries=0):
            calls.append(("json", url, timeout, retries))
            return {
                "assets": [
                    {
                        "name": "sms.zip",
                        "size": 10,
                        "browser_download_url": "https://github.com/KPI0/Air724UG-SMS/releases/download/v1/sms.zip",
                    }
                ]
            }

        def fake_probe(url, timeout=0, retries=0):
            calls.append(("probe", url, timeout, retries))
            return True, "HTTP 200"

        result = test_update_proxy_connectivity(
            "KPI0",
            "Air724UG-SMS",
            "api.proxy",
            "download.proxy",
            get_json=fake_get_json,
            probe=fake_probe,
        )
        formatted = format_proxy_test_result(result)

        self.assertEqual(result["checks"][0], ("https://api.proxy/", True, "OK"))
        self.assertTrue(result["download_ok"])
        self.assertIn("API 代理可用", formatted)
        self.assertIn("下载代理可用", formatted)
        self.assertEqual(calls[0][0], "json")
        self.assertEqual(calls[1][0], "probe")

    def test_app_launch_helpers(self):
        payload = encode_restart_args(["--flag", "中文", 3])
        self.assertEqual(decode_restart_args(payload), ["--flag", "中文", 3])
        self.assertEqual(decode_restart_args(""), [])

        env = get_clean_restart_env({
            "_MEIPASS": "old",
            "TCL_LIBRARY": "old",
            "KEEP": "1",
        })
        self.assertEqual(env["KEEP"], "1")
        self.assertEqual(env["PYINSTALLER_RESET_ENVIRONMENT"], "1")
        self.assertNotIn("_MEIPASS", env)
        self.assertNotIn("TCL_LIBRARY", env)

        target, script_arg, workdir = get_launch_target_and_args(
            frozen=True,
            executable=r"C:\app\sms.exe",
            argv0="ignored.pyw",
        )
        self.assertEqual(target, os.path.abspath(r"C:\app\sms.exe"))
        self.assertEqual(script_arg, "")
        self.assertEqual(workdir, os.path.dirname(os.path.abspath(r"C:\app\sms.exe")))

    def test_shortcut_name_sanitizer(self):
        self.assertEqual(sanitize_shortcut_name('sms:/bad*name'), "sms__bad_name.lnk")
        self.assertEqual(sanitize_shortcut_name("Already.lnk"), "Already.lnk")
        self.assertEqual(sanitize_shortcut_name("   "), "sms.lnk")

    def test_windows_runtime_pure_helpers(self):
        self.assertEqual(port_mutex_name("COM5"), "Air724UG_PORT_COM5")
        self.assertEqual(normalize_serial_device_path("COM5"), r"\\.\COM5")
        self.assertEqual(normalize_serial_device_path(r"\\.\COM7"), r"\\.\COM7")
        self.assertEqual(normalize_serial_device_path("  "), "")
        self.assertTrue(is_existing_instance_error(183))
        self.assertTrue(is_existing_instance_error(5))
        self.assertFalse(is_existing_instance_error(0))

    def test_serial_status_parsers(self):
        self.assertEqual(
            parse_temperature("[I]-[ril.proatc] +RFTEMPERATURE: 28.84"),
            "28.84",
        )
        self.assertEqual(
            parse_cesq_rsrp("[I]-[ril.proatc] +CESQ: 99,99,255,255,26,49"),
            "49",
        )

        insights = parse_serial_debug_insights('[I]-[ril.proatc] +COPS: 0,2,"46011",7')
        self.assertEqual(insights, [">>> \u8bc6\u522b\u5230\u7f51\u7edc\u8fd0\u8425\u5546\uff1a\u4e2d\u56fd\u7535\u4fe1 (46011)"])

        self.assertEqual(
            parse_serial_debug_insights("[I]-[ril.proatc] +CPIN: READY"),
            [">>> \u8bc6\u522b\u5230SIM\u5361\u72b6\u6001\uff1aPIN\u7801\u9501\u672a\u5f00\u542f (SIM\u5361\u6b63\u5e38)"],
        )

    def test_serial_call_event_helpers(self):
        self.assertEqual(parse_clip_number('+CLIP: "+8613812345678",129'), "+8613812345678")
        self.assertEqual(parse_clip_number("no caller"), "\u672a\u77e5\u53f7\u7801")

        self.assertEqual(
            evaluate_call_filter("+8613812345678", "Whitelist", ["13812345678"], []),
            (False, ""),
        )
        self.assertEqual(
            evaluate_call_filter("+8613812345678", "Whitelist", ["10086"], []),
            (True, "\u4e0d\u5728\u767d\u540d\u5355"),
        )
        self.assertEqual(
            evaluate_call_filter("+8613812345678", "Blacklist", [], ["13812345678"]),
            (True, "\u547d\u4e2d\u9ed1\u540d\u5355"),
        )

        self.assertTrue(is_new_clip("10086", "10010", 10.0, 9.0))
        self.assertFalse(is_new_clip("10086", "10086", 10.0, 7.0))
        self.assertTrue(is_ring_line("[I]-[ril] RING"))
        self.assertTrue(is_hangup_event('[I]-[ril] +CIEV: "CALL",0'))
        self.assertTrue(is_call_connected_event(' +CIEV: "CALL",1'))
        self.assertTrue(is_sms_collection_boundary("[I]-[handler] next"))
        self.assertFalse(is_sms_collection_boundary("[\u6e29\u99a8\u63d0\u793a] text"))

    def test_sms_pending_collector(self):
        def parse_head(text):
            return ("+8613812345678", text.split(" ", 1)[1])

        collector = SmsPendingCollector(parse_head, initial_timeout=1.0, fragment_timeout=0.4, max_follow_lines=2)
        line = SMS_CALLBACK_PREFIX + " +8613812345678 26/06/08,12:00:00+32 first"

        self.assertEqual(callback_body_from_line(line), "+8613812345678 26/06/08,12:00:00+32 first")
        self.assertTrue(collector.start(callback_body_from_line(line), now=10.0))
        self.assertTrue(collector.active)
        self.assertFalse(collector.expired(10.5))
        self.assertTrue(collector.expired(11.1))

        self.assertEqual(collector.consume_line("second", now=11.2), "consumed")
        self.assertEqual(collector.consume_line("third", now=11.3), "flush")
        pending = collector.flush()

        self.assertEqual(pending.callback_head, "+8613812345678 26/06/08,12:00:00+32 first")
        self.assertEqual(pending.full_msg, "26/06/08,12:00:00+32 firstsecondthird")
        self.assertEqual(pending.display_lines[0], "\U0001f4e9 \u6536\u5230\u77ed\u4fe1\uff1a")
        self.assertFalse(collector.active)

    def test_sms_processing_helpers(self):
        self.assertTrue(sms_keyword_hit("hello KEY", ["key"]))
        self.assertFalse(sms_keyword_hit("hello", ["key"]))
        self.assertTrue(sms_keyword_hit("hello", []))

        entries = build_unmatched_sms_log_entries(
            r"C:\logs",
            "COM:5",
            "body",
            now=datetime(2026, 6, 8, 12, 30, 0),
        )
        self.assertEqual(len(entries), 2)
        self.assertTrue(entries[0][0].endswith(os.path.join("sms_COM_5_2026-06-08.txt")))
        self.assertIn("body", entries[1][1])

        state = {}
        self.assertEqual(repeat_count_message(state, "k", "msg", 3), "msg")
        self.assertEqual(repeat_count_message(state, "k", "msg", 3), "msg")
        self.assertIn("msg", repeat_count_message(state, "k", "msg", 3))
        self.assertIsNone(repeat_count_message(state, "k", "msg", 3))

    def test_process_pending_sms_shows_keyword_match(self):
        class Pending:
            callback_head = "head"
            full_msg = "hello keyword"
            display_lines = ["header", "hello keyword"]

        events = []
        result = process_pending_sms(
            Pending(),
            ["keyword"],
            True,
            r"C:\logs",
            "COM5",
            {},
            3,
            lambda msg, event_type="sms": events.append(("push", msg, event_type)),
            lambda head, msg: events.append(("cloud", head, msg)),
            lambda msg, tag="normal": events.append(("ui", msg, tag)),
            lambda: events.append(("sound",)),
            lambda msg: events.append(("popup", msg)),
            lambda item: events.append(("file", item)),
            lambda msg, tag="normal": events.append(("system", msg, tag)),
        )

        self.assertEqual(result, "shown")
        self.assertIn(("push", "hello keyword", "sms"), events)
        self.assertIn(("cloud", "head", "hello keyword"), events)
        self.assertIn(("ui", "header", "normal"), events)
        self.assertIn(("ui", "hello keyword", "sms"), events)
        self.assertIn(("sound",), events)
        self.assertIn(("popup", "hello keyword"), events)
        self.assertFalse([event for event in events if event[0] == "file"])

    def test_process_pending_sms_logs_unmatched_message(self):
        class Pending:
            callback_head = "head"
            full_msg = "hello"
            display_lines = ["header", "hello"]

        events = []
        repeat_state = {}
        result = process_pending_sms(
            Pending(),
            ["missing"],
            True,
            r"C:\logs",
            "COM5",
            repeat_state,
            3,
            lambda msg, event_type="sms": events.append(("push", msg, event_type)),
            lambda head, msg: events.append(("cloud", head, msg)),
            lambda msg, tag="normal": events.append(("ui", msg, tag)),
            lambda: events.append(("sound",)),
            lambda msg: events.append(("popup", msg)),
            lambda item: events.append(("file", item)),
            lambda msg, tag="normal": events.append(("system", msg, tag)),
        )

        self.assertEqual(result, "ignored")
        self.assertIn(("push", "hello", "sms"), events)
        self.assertIn(("cloud", "head", "hello"), events)
        self.assertEqual(len([event for event in events if event[0] == "file"]), 2)
        self.assertEqual(len([event for event in events if event[0] == "system"]), 1)
        self.assertFalse([event for event in events if event[0] in ("ui", "sound", "popup")])

    def test_lte_cell_parser(self):
        insights = parse_serial_debug_insights(
            "[I]-[ril.proatc] +EEMLTESVC: 1120,17,0,12345,0,0,0,0,0,67890"
        )

        self.assertEqual(insights[0], ">>> \u89e3\u6790\u5230\u57fa\u7ad9\u5b9a\u4f4d\u6570\u636e\uff1a")
        self.assertIn("460", insights[1])
        self.assertIn("11", insights[2])

    def test_serial_debug_helpers(self):
        self.assertGreater(SERIAL_DEBUG_MAX_STORE_LINES, SERIAL_DEBUG_MAX_VISIBLE_LINES)
        self.assertIn(("AT", "\u6d4b\u8bd5\u901a\u4fe1"), COMMON_SERIAL_COMMANDS)
        self.assertEqual(quick_command_label("AT", "test"), "AT  (test)")

        payload, suffix = build_serial_command_payload("AT", append_crlf=True)
        self.assertEqual(payload, b"AT\r\n")
        self.assertEqual(suffix, "\\r\\n")

        payload, suffix = build_serial_command_payload("AT", append_crlf=False)
        self.assertEqual(payload, b"AT")
        self.assertEqual(suffix, "")

    def test_serial_debug_command_builders(self):
        self.assertEqual(build_pin_unlock_command("1234"), 'AT+CPIN="1234"')
        self.assertEqual(build_puk_unlock_command("12345678", "1111"), 'AT+CPIN="12345678","1111"')
        self.assertEqual(build_pin_lock_command("1234", enable=True), 'AT+CLCK="SC",1,"1234"')
        self.assertEqual(build_pin_lock_command("1234", enable=False), 'AT+CLCK="SC",0,"1234"')
        self.assertEqual(build_pin_change_command("1234", "5678"), 'AT+CPWD="SC","1234","5678"')
        self.assertEqual(normalize_own_number("13812345678"), "+8613812345678")
        self.assertEqual(
            build_own_number_commands("13812345678"),
            ('AT+CPBS="ON"', 'AT+CPBW=1,"+8613812345678",145', "AT+CNUM"),
        )
        self.assertEqual(build_sn_command("ABC123"), "AT+WISN=ABC123")
        self.assertEqual(normalize_dial_number("+8613812345678"), "13812345678")
        self.assertEqual(build_dial_command("8613812345678"), "ATD13812345678;")
        self.assertEqual(HANGUP_COMMAND, "ATH")

    def test_serial_sender_writes_command_and_sequence(self):
        serial_obj = self.FakeSerial()
        debug_lines = []

        self.assertTrue(write_serial_command(serial_obj, "AT", push_debug=debug_lines.append))
        self.assertEqual(serial_obj.writes, [b"AT\r\n"])
        self.assertEqual(debug_lines[-1], ">>> 发送: AT\\r\\n")

        self.assertTrue(
            write_serial_command_sequence(
                serial_obj,
                ["AT", "ATI"],
                push_debug=debug_lines.append,
                sleep_func=lambda _seconds: None,
            )
        )
        self.assertEqual(serial_obj.writes[-2:], [b"AT\r\n", b"ATI\r\n"])
        self.assertIn(">>> 发送: ATI\\r\\n", debug_lines)

    def test_serial_sender_reports_closed_port(self):
        serial_obj = self.FakeSerial(is_open=False)
        debug_lines = []
        ui_lines = []

        self.assertFalse(write_serial_command(serial_obj, "AT", push_debug=debug_lines.append))
        self.assertFalse(
            write_text_sms_pdu(
                serial_obj,
                "+8613812345678",
                "测试",
                push_debug=debug_lines.append,
                port_ui=lambda *args: ui_lines.append(args),
                sleep_func=lambda _seconds: None,
            )
        )
        self.assertEqual(debug_lines, [">>> 发送失败: 串口未连接", ">>> 发送失败: 串口未连接"])
        self.assertEqual(ui_lines, [("❌ 发送短信失败：串口未连接", "normal")])

    def test_serial_sender_writes_text_sms_pdu(self):
        serial_obj = self.FakeSerial()
        debug_lines = []
        ui_lines = []

        self.assertTrue(
            write_text_sms_pdu(
                serial_obj,
                "+8613812345678",
                "测试",
                push_debug=debug_lines.append,
                port_ui=lambda *args: ui_lines.append(args),
                sleep_func=lambda _seconds: None,
            )
        )

        self.assertEqual(serial_obj.writes[0], b"AT+CMGF=0\r\n")
        self.assertTrue(serial_obj.writes[1].startswith(b"AT+CMGS="))
        self.assertTrue(serial_obj.writes[2].endswith(b"\x1a"))
        self.assertIn(">>> 发送 PDU 正文及 Ctrl+Z，等待模组响应...", debug_lines)
        self.assertEqual(ui_lines[0], ("📤 发送短信至 +8613812345678：", "normal"))


if __name__ == "__main__":
    unittest.main()
