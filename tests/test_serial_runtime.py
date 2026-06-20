import unittest

from sms_core.serial_runtime import (
    SerialRuntimeCallbacks,
    SerialRuntimeConfig,
    SerialRuntimeState,
    handle_serial_runtime_line,
    run_serial_runtime_thread,
    run_serial_thread_loop,
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


class SerialRuntimeTests(unittest.TestCase):
    def test_blank_line_flushes_expired_sms(self):
        calls = []
        state = SerialRuntimeState.create(parse_head)
        state.sms_collector.start("+8613812345678 hello code", now=10.0)

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
            '+CLIP: "+8613812345678",129',
            10.0,
            "COM5",
            False,
            runtime_config(),
            runtime_callbacks(calls),
            {},
        )

        self.assertFalse(result.continue_read)
        self.assertEqual(state.call_state.last_clip_num, "+8613812345678")
        self.assertEqual(state.call_state.ring_timeout_target, 22.0)
        self.assertIn(("call_popup", ("+8613812345678",)), calls)
        self.assertTrue(any(item[0] == "push" and item[2] == {"event_type": "call"} for item in calls))

    def test_blocked_call_hangs_up_and_stops_processing(self):
        calls = []
        state = SerialRuntimeState.create(parse_head)

        result = handle_serial_runtime_line(
            state,
            '+CLIP: "+8613812345678",129',
            10.0,
            "COM5",
            False,
            runtime_config(call_filter_mode="Blacklist", call_blacklist=["13812345678"]),
            runtime_callbacks(calls),
            {},
        )

        self.assertTrue(result.continue_read)
        self.assertIn(("hangup",), calls)
        self.assertNotIn(("call_popup", ("+8613812345678",)), calls)

    def test_sms_callback_starts_collection(self):
        calls = []
        state = SerialRuntimeState.create(parse_head)

        result = handle_serial_runtime_line(
            state,
            "[I]-[handler_sms.smsCallback] +8613812345678 hello",
            10.0,
            "COM5",
            False,
            runtime_config(),
            runtime_callbacks(calls),
            {},
        )

        self.assertTrue(result.continue_read)
        self.assertTrue(state.sms_collector.active)
        self.assertEqual(state.sms_collector.callback_head, "+8613812345678 hello")

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

    def test_run_serial_thread_loop_skips_missing_target_port(self):
        calls = []
        keep_running = [True, False]

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


if __name__ == "__main__":
    unittest.main()
