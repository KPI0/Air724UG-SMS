import unittest

from sms_core.serial_app_runtime import (
    build_serial_app_wiring,
    build_serial_runtime_callbacks,
    build_serial_runtime_config,
    run_serial_app_from_wiring,
    run_serial_reader_namespace_runtime,
    run_serial_app_runtime,
)


def app_settings(**overrides):
    values = {
        "keywords": lambda: ["otp"],
        "log_unmatched_sms": lambda: True,
        "log_dir": lambda: "logs",
        "log_prefix": lambda: "COM5",
        "error_repeat_limit": lambda: 3,
        "call_filter_mode": lambda: "Blacklist",
        "call_whitelist": lambda: ["10086"],
        "call_blacklist": lambda: ["10010"],
    }
    values.update(overrides)
    return values


def app_callbacks(calls):
    return {
        "enqueue_third_push": lambda *args: calls.append(("push", args)),
        "send_cloud_sms_event": lambda *args: calls.append(("cloud_sms", args)),
        "port_ui": lambda *args: calls.append(("port_ui", args)),
        "play_alert": lambda: calls.append(("alert",)),
        "show_sms_popup": lambda *args: calls.append(("sms_popup", args)),
        "file_log": lambda *args: calls.append(("file", args)),
        "system_ui": lambda *args: calls.append(("system", args)),
        "push_serial_debug": lambda *args: calls.append(("debug", args)),
        "send_cloud_serial_log": lambda *args: calls.append(("cloud_log", args)),
        "capture_cloud_device_imei": lambda *args: calls.append(("imei", args)),
        "set_temperature": lambda *args: calls.append(("temp", args)),
        "set_signal": lambda *args: calls.append(("signal", args)),
        "set_status": lambda *args: calls.append(("status", args)),
        "close_call_popup": lambda: calls.append(("close_popup",)),
        "send_call_hangup": lambda: calls.append(("hangup",)),
        "show_call_popup": lambda *args: calls.append(("call_popup", args)),
        "schedule_connected_log": lambda *args: calls.append(("connected_log", args)),
        "serial_error_ui": lambda *args: calls.append(("serial_error", args)),
    }


class SerialAppRuntimeTests(unittest.TestCase):
    def test_build_serial_runtime_config_reads_settings(self):
        config = build_serial_runtime_config(app_settings())

        self.assertEqual(config.keywords, ["otp"])
        self.assertTrue(config.log_unmatched_sms)
        self.assertEqual(config.log_dir, "logs")
        self.assertEqual(config.log_prefix, "COM5")
        self.assertEqual(config.call_blacklist, ["10010"])

    def test_build_serial_runtime_callbacks_maps_callbacks(self):
        calls = []
        callbacks = build_serial_runtime_callbacks(app_callbacks(calls))

        callbacks.set_status("ready", "green")
        callbacks.close_call_popup()

        self.assertEqual(calls, [("status", ("ready", "green")), ("close_popup",)])

    def test_build_serial_app_wiring_maps_grouped_dependencies(self):
        calls = []
        settings, callbacks, state, io, reconnect = build_serial_app_wiring(
            values=app_settings(),
            callbacks=app_callbacks(calls),
            state_access={
                "get_call_state": lambda: (0, ""),
                "set_call_state": lambda *_: None,
                "popup_active": lambda: False,
                "ignore_repeat_state": {"seen": True},
                "serial_running": lambda: True,
                "port": lambda: "COM5",
                "baud": lambda: 115200,
                "set_log_prefix": lambda value: calls.append(("prefix", value)),
            },
            io_callbacks={
                "resolve_target_port": lambda: "COM5",
                "open_and_initialize_serial": lambda port: calls.append(("open", port)),
                "read_serial_line": lambda: b"",
                "safe_close_serial": lambda: calls.append(("close",)),
            },
            reconnect_callbacks={
                "stop_requested": lambda: False,
                "interval": lambda: 2,
                "wakeup_wait": lambda timeout=None: calls.append(("wait", timeout)),
                "wakeup_clear": lambda: calls.append(("clear",)),
                "try_manual_rebind_after_error": lambda error: False,
            },
        )

        self.assertEqual(settings["log_prefix"](), "COM5")
        callbacks["set_status"]("ready", "green")
        self.assertEqual(state["ignore_repeat_state"], {"seen": True})
        self.assertEqual(io["resolve_target_port"](), "COM5")
        self.assertFalse(reconnect["stop_requested"]())
        self.assertIn(("status", ("ready", "green")), calls)

    def test_run_serial_app_from_wiring_builds_and_runs(self):
        calls = []

        def run_app(**kwargs):
            calls.append(kwargs)
            return "ran"

        result = run_serial_app_from_wiring(
            parse_callback_head="parser",
            values=app_settings(),
            callbacks=app_callbacks(calls),
            state_access={
                "get_call_state": lambda: (0, ""),
                "set_call_state": lambda *_: None,
                "popup_active": lambda: False,
                "ignore_repeat_state": {},
                "serial_running": lambda: True,
                "port": lambda: "COM5",
                "baud": lambda: 115200,
                "set_log_prefix": lambda value: None,
            },
            io_callbacks={
                "resolve_target_port": lambda: "COM5",
                "open_and_initialize_serial": lambda port: None,
                "read_serial_line": lambda: b"",
                "safe_close_serial": lambda: None,
            },
            reconnect_callbacks={
                "stop_requested": lambda: False,
                "interval": lambda: 2,
                "wakeup_wait": lambda timeout=None: None,
                "wakeup_clear": lambda: None,
                "try_manual_rebind_after_error": lambda error: False,
            },
            apply_disconnect_effects="disconnect",
            run_app=run_app,
        )

        self.assertEqual(result, "ran")
        self.assertEqual(calls[0]["parse_callback_head"], "parser")
        self.assertEqual(calls[0]["apply_disconnect_effects"], "disconnect")
        self.assertEqual(calls[0]["settings"]["keywords"](), ["otp"])

    def test_run_serial_reader_namespace_runtime_adapts_globals_namespace(self):
        calls = []

        class FakeQueue:
            def put_nowait(self, value):
                calls.append(("file", value))

        class FakeEvent:
            def __init__(self):
                self.wait_calls = []
                self.cleared = False

            def is_set(self):
                return False

            def wait(self, timeout=None):
                self.wait_calls.append(timeout)

            def clear(self):
                self.cleared = True

        class FakeSmsCoordinator:
            def observe_line(self, line, connection=None):
                calls.append(("sms_line", line, connection))

            def cancel_active(self, error):
                calls.append(("sms_cancel", error))
                return True

        stop_event = FakeEvent()
        wakeup_event = FakeEvent()
        serial_obj = object()
        namespace = {
            "KEYWORDS": ["otp"],
            "LOG_UNMATCHED_SMS": True,
            "LOG_DIR": "logs",
            "LOG_PREFIX": "COM5",
            "ERROR_REPEAT_LIMIT": 3,
            "CALL_FILTER_MODE": "Blacklist",
            "CALL_WHITELIST": ["10086"],
            "CALL_BLACKLIST": ["10010"],
            "enqueue_third_push": lambda *args: calls.append(("push", args)),
            "_cloud_send_sms_event": lambda *args: calls.append(("cloud_sms", args)),
            "port_ui": lambda *args: calls.append(("port_ui", args)),
            "play_alert": lambda: calls.append(("alert",)),
            "show_sms_popup": lambda *args: calls.append(("sms_popup", args)),
            "FILE_LOG_Q": FakeQueue(),
            "system_ui": lambda *args: calls.append(("system", args)),
            "_push_serial_debug": lambda *args: calls.append(("debug", args)),
            "_cloud_send_serial_log": lambda *args: calls.append(("cloud_log", args)),
            "notify_cloud_channel_status": lambda *args: calls.append(("channel_status", args)),
            "_maybe_capture_cloud_device_imei": lambda *args: calls.append(("imei", args)),
            "set_temperature": lambda *args: calls.append(("temp", args)),
            "set_signal": lambda *args: calls.append(("signal", args)),
            "set_status": lambda *args: calls.append(("status", args)),
            "close_call_popup": lambda: calls.append(("close_popup",)),
            "send_call_hangup_command": lambda: calls.append(("hangup",)),
            "show_call_popup": lambda *args: calls.append(("call_popup", args)),
            "schedule_delayed_connected_log": lambda *args: calls.append(("connected", args)),
            "serial_error_ui": lambda *args: calls.append(("error", args)),
            "get_serial_call_state": lambda: (0.0, ""),
            "set_serial_call_state": lambda *args: calls.append(("call_state", args)),
            "current_call_popup": None,
            "_sms_ignore_repeat_state": {"seen": True},
            "serial_running": True,
            "PORT": "COM5",
            "BAUD": 115200,
            "set_serial_log_prefix": lambda value: calls.append(("prefix", value)),
            "resolve_serial_target_port": lambda: "COM5",
            "open_and_initialize_serial": lambda port: calls.append(("open", port)),
            "read_serial_line_safely": lambda: b"",
            "safe_close_serial": lambda: calls.append(("close_serial",)),
            "serial_stop_event": stop_event,
            "RECONNECT_INTERVAL": 2,
            "serial_wakeup_event": wakeup_event,
            "try_manual_rebind_after_error": lambda error: False,
            "serial_obj": serial_obj,
            "SMS_SEND_COORDINATOR": FakeSmsCoordinator(),
        }

        def run_reader(**kwargs):
            calls.append(kwargs)
            kwargs["values"]["log_prefix"]()
            kwargs["callbacks"]["file_log"]("line")
            kwargs["callbacks"]["observe_sms_send_line"]("OK")
            kwargs["callbacks"]["cancel_sms_send"]("gone")
            kwargs["callbacks"]["notify_cloud_channel_status"](True)
            self.assertFalse(kwargs["state_access"]["popup_active"]())
            self.assertFalse(kwargs["reconnect_callbacks"]["stop_requested"]())
            return "ran"

        result = run_serial_reader_namespace_runtime(
            namespace,
            parse_callback_head="parser",
            apply_disconnect_effects="disconnect",
            run_reader=run_reader,
        )

        self.assertEqual(result, "ran")
        self.assertEqual(calls[0]["parse_callback_head"], "parser")
        self.assertEqual(calls[0]["apply_disconnect_effects"], "disconnect")
        self.assertIn(("file", "line"), calls)
        self.assertIn(("sms_line", "OK", serial_obj), calls)
        self.assertIn(("sms_cancel", "gone"), calls)
        self.assertIn(("channel_status", (True,)), calls)

    def test_run_serial_app_runtime_wires_loop_callbacks(self):
        calls = []
        log_prefixes = []
        waited = []
        cleared = []
        call_state = []
        captured = {}

        def fake_run_serial_runtime_thread(**kwargs):
            captured.update(kwargs)
            self.assertTrue(kwargs["should_continue"]())
            kwargs["set_connecting_status"]("COM7")
            kwargs["on_connected_port"]("COM7:")
            kwargs["wait_before_retry"]()
            return "state"

        result = run_serial_app_runtime(
            parse_callback_head=lambda text: ("sender", text),
            settings=app_settings(),
            callbacks=app_callbacks(calls),
            state={
                "get_call_state": lambda: (0.0, ""),
                "set_call_state": lambda *args: call_state.append(args),
                "popup_active": lambda: False,
                "ignore_repeat_state": {"seen": 1},
                "serial_running": lambda: True,
                "port": lambda: "COM7",
                "baud": lambda: 115200,
                "set_log_prefix": log_prefixes.append,
            },
            io={
                "resolve_target_port": lambda: "COM7",
                "open_and_initialize_serial": lambda port: calls.append(("open", port)),
                "read_serial_line": lambda: b"",
                "safe_close_serial": lambda: calls.append(("close_serial",)),
            },
            reconnect={
                "stop_requested": lambda: False,
                "interval": lambda: 2,
                "wakeup_wait": lambda timeout=None: waited.append(timeout),
                "wakeup_clear": lambda: cleared.append("cleared"),
                "try_manual_rebind_after_error": lambda error: False,
            },
            apply_disconnect_effects=lambda *args: calls.append(("disconnect", args)),
            run_runtime_thread=fake_run_serial_runtime_thread,
        )

        self.assertEqual(result, "state")
        self.assertEqual(calls[0][0], "status")
        self.assertIn("COM7", calls[0][1][0])
        self.assertEqual(log_prefixes, ["COM7_"])
        self.assertIn(("connected_log", ("COM7", 115200)), calls)
        self.assertEqual(waited, [2])
        self.assertEqual(cleared, ["cleared"])
        self.assertEqual(captured["ignore_repeat_state"], {"seen": 1})

    def test_run_serial_app_runtime_handles_disconnect(self):
        calls = []
        log_prefixes = []
        callbacks = app_callbacks(calls)
        callbacks["cancel_sms_send"] = lambda error: calls.append(("cancel_sms", error))

        def fake_run_serial_runtime_thread(**kwargs):
            return kwargs["handle_disconnect"](RuntimeError("gone"), "COM7")

        result = run_serial_app_runtime(
            parse_callback_head=lambda text: ("sender", text),
            settings=app_settings(),
            callbacks=callbacks,
            state={
                "get_call_state": lambda: (0.0, ""),
                "set_call_state": lambda *_: None,
                "popup_active": lambda: False,
                "ignore_repeat_state": {},
                "serial_running": lambda: True,
                "port": lambda: "COM5",
                "baud": lambda: 115200,
                "set_log_prefix": log_prefixes.append,
            },
            io={
                "resolve_target_port": lambda: "COM7",
                "open_and_initialize_serial": lambda port: None,
                "read_serial_line": lambda: b"",
                "safe_close_serial": lambda: None,
            },
            reconnect={
                "stop_requested": lambda: False,
                "interval": lambda: 2,
                "wakeup_wait": lambda timeout=None: None,
                "wakeup_clear": lambda: None,
                "try_manual_rebind_after_error": lambda error: calls.append(("rebind", str(error))) or True,
            },
            apply_disconnect_effects=lambda *args: calls.append(("disconnect", args)),
            run_runtime_thread=fake_run_serial_runtime_thread,
        )

        self.assertTrue(result)
        self.assertEqual(log_prefixes, ["system"])
        self.assertEqual(calls[0], ("close_popup",))
        self.assertEqual(calls[1], ("cancel_sms", "串口连接已断开：gone"))
        self.assertEqual(calls[2][0], "disconnect")
        self.assertEqual(calls[-1], ("rebind", "gone"))


if __name__ == "__main__":
    unittest.main()
