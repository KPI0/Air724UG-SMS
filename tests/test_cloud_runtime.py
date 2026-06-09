import unittest
import configparser

from sms_core.cloud_runtime import (
    CloudControlSettings,
    cloud_auto_upload_action,
    cloud_backoff_sleep_ticks,
    cloud_control_form_values,
    cloud_control_save_kwargs,
    cloud_control_state,
    cloud_restart_attempt_action,
    cloud_restarting_status,
    cloud_start_thread_action,
    cloud_stopped_status,
    next_cloud_backoff,
    read_cloud_control_settings,
    restart_cloud_control_runtime,
    start_cloud_control_runtime,
    stop_cloud_control_runtime,
    update_cloud_control_settings,
    validate_cloud_start,
    write_cloud_control_settings,
)


class CloudRuntimeTests(unittest.TestCase):
    class FakeLock:
        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

    class FakeEvent:
        def __init__(self, set_value=False):
            self.value = set_value
            self.cleared = False

        def is_set(self):
            return self.value

        def clear(self):
            self.value = False
            self.cleared = True

    class FakeThread:
        def __init__(self, target=None, args=(), daemon=False, alive=False):
            self.target = target
            self.args = args
            self.daemon = daemon
            self.alive = alive
            self.started = False
            self.joined = []

        def is_alive(self):
            return self.alive

        def start(self):
            self.started = True
            if self.target:
                self.target(*self.args)

        def join(self, timeout=None):
            self.joined.append(timeout)
            self.alive = False

    def test_validate_cloud_start_requires_websockets(self):
        result = validate_cloud_start(False, "ws://127.0.0.1:8765", "secret")

        self.assertFalse(result.ok)
        self.assertEqual(result.status_text, "🌐 缺少依赖")
        self.assertEqual(result.status_color, "#cc0000")
        self.assertIn("websockets", result.warning_message)

    def test_validate_cloud_start_requires_url(self):
        result = validate_cloud_start(True, "  ", "secret")

        self.assertFalse(result.ok)
        self.assertEqual(result.status_text, "🌐 未配置")
        self.assertIn("WebSocket", result.warning_message)

    def test_validate_cloud_start_rejects_non_websocket_scheme(self):
        result = validate_cloud_start(True, "http://example.com/ws", "secret")

        self.assertFalse(result.ok)
        self.assertEqual(result.url, "http://example.com/ws")
        self.assertEqual(result.status_text, "🌐 地址错误")

    def test_validate_cloud_start_requires_secret(self):
        result = validate_cloud_start(True, "ws://127.0.0.1:8765", " ")

        self.assertFalse(result.ok)
        self.assertEqual(result.url, "ws://127.0.0.1:8765/websocket")
        self.assertEqual(result.status_text, "🌐 密码未配置")

    def test_validate_cloud_start_normalizes_valid_url(self):
        result = validate_cloud_start(True, "ws://127.0.0.1:8765", "secret")

        self.assertTrue(result.ok)
        self.assertEqual(result.url, "ws://127.0.0.1:8765/websocket")

    def test_cloud_stopped_status(self):
        self.assertEqual(cloud_stopped_status(False), "🌐 已关闭")
        self.assertEqual(cloud_stopped_status(True), "🌐 已断开")

    def test_next_cloud_backoff(self):
        self.assertEqual(next_cloud_backoff(2), 3.0)
        self.assertEqual(next_cloud_backoff(100), 60.0)
        self.assertEqual(next_cloud_backoff("bad"), 1.5)

    def test_cloud_backoff_sleep_ticks(self):
        self.assertEqual(cloud_backoff_sleep_ticks(2), 20)
        self.assertEqual(cloud_backoff_sleep_ticks("bad"), 10)

    def test_cloud_control_state_and_form_values(self):
        state = cloud_control_state(True, False, "ws://host", "secret", 5)

        self.assertEqual(state["enabled"], True)
        self.assertEqual(state["auto_upload"], False)
        self.assertEqual(state["url"], "ws://host")
        self.assertEqual(state["secret"], "secret")
        self.assertEqual(state["reconnect_interval"], 5)

        values = cloud_control_form_values((False, "ws://host", "7", "secret", True), enabled_override=True)
        self.assertTrue(values.enabled)
        self.assertEqual(values.url, "ws://host")
        self.assertEqual(values.reconnect_interval, "7")
        self.assertEqual(values.device_secret, "secret")
        self.assertTrue(values.auto_upload)

    def test_cloud_control_save_kwargs(self):
        self.assertEqual(
            cloud_control_save_kwargs((True, "ws://host", 5, "secret", False)),
            {
                "enabled": True,
                "url": "ws://host",
                "reconnect_interval": 5,
                "device_secret": "secret",
                "auto_upload": False,
            },
        )
        self.assertFalse(
            cloud_control_save_kwargs((True, "ws://host", 5, "secret", False), enabled_override=False)["enabled"]
        )

    def test_cloud_control_settings_read_update_and_write(self):
        config = configparser.ConfigParser()
        config["cloud_control"] = {
            "enabled": "1",
            "url": "ws://127.0.0.1:8765",
            "device_imei": "old",
            "device_secret": " secret ",
            "reconnect_interval": "0",
            "auto_upload": "1",
        }

        settings = read_cloud_control_settings(config)
        self.assertEqual(settings, CloudControlSettings(
            enabled=True,
            url="ws://127.0.0.1:8765/websocket",
            reconnect_interval=1,
            device_secret="secret",
            auto_upload=True,
        ))

        updated = update_cloud_control_settings(
            settings,
            enabled=False,
            url="wss://example.com",
            reconnect_interval="bad",
            device_secret=" next ",
            auto_upload=False,
        )
        self.assertEqual(updated.url, "wss://example.com/websocket")
        self.assertEqual(updated.reconnect_interval, 5)
        self.assertEqual(updated.device_secret, "next")

        write_cloud_control_settings(config, updated)
        self.assertFalse(config.has_option("cloud_control", "device_imei"))
        self.assertEqual(config.get("cloud_control", "enabled"), "0")
        self.assertEqual(config.get("cloud_control", "auto_upload"), "0")
        self.assertEqual(config.get("cloud_control", "reconnect_interval"), "5")

    def test_cloud_auto_upload_action(self):
        self.assertEqual(cloud_auto_upload_action(False, True, True, True, True), "register")
        self.assertEqual(cloud_auto_upload_action(True, False, False, False, False), "unregister")
        self.assertEqual(cloud_auto_upload_action(False, True, False, True, True), "")

    def test_cloud_thread_lifecycle_actions(self):
        self.assertEqual(cloud_restarting_status(), "🌐 正在重启")
        self.assertEqual(cloud_start_thread_action(True, True), "restarting")
        self.assertEqual(cloud_start_thread_action(True, False), "already_running")
        self.assertEqual(cloud_start_thread_action(False, True), "start")

        self.assertEqual(cloud_restart_attempt_action(1, 2, True, False, False), "cancel")
        self.assertEqual(cloud_restart_attempt_action(1, 1, False, False, False), "cancel")
        self.assertEqual(cloud_restart_attempt_action(1, 1, True, True, True), "wait")
        self.assertEqual(cloud_restart_attempt_action(1, 1, True, True, False), "start")

    def test_start_cloud_control_runtime_reports_validation_errors(self):
        calls = []

        ok = start_cloud_control_runtime(
            websockets_available=False,
            url="ws://host",
            device_secret="secret",
            reconnect_interval=5,
            show_errors=True,
            validate_start=validate_cloud_start,
            set_cloud_status=lambda text, color: calls.append(("status", text, color)),
            log_missing_dependency=lambda: calls.append(("missing",)),
            show_warning=lambda title, message: calls.append(("warning", title, message)),
            runtime_imei=lambda: "imei",
            request_device_imei=lambda: calls.append(("imei",)),
            lock=self.FakeLock(),
            get_thread=lambda: None,
            set_thread=lambda thread: calls.append(("thread", thread)),
            stop_event=self.FakeEvent(),
            thread_factory=self.FakeThread,
            thread_target=lambda *_: calls.append(("target",)),
        )

        self.assertFalse(ok)
        self.assertEqual(calls[0][0], "status")
        self.assertIn(("missing",), calls)
        self.assertEqual(calls[-1][0], "warning")

    def test_start_cloud_control_runtime_starts_thread_and_requests_imei(self):
        calls = []
        stop_event = self.FakeEvent(set_value=True)
        stored = []

        ok = start_cloud_control_runtime(
            websockets_available=True,
            url="ws://host",
            device_secret="secret",
            reconnect_interval=7,
            show_errors=False,
            validate_start=validate_cloud_start,
            set_cloud_status=lambda text, color: calls.append(("status", text, color)),
            log_missing_dependency=lambda: calls.append(("missing",)),
            show_warning=lambda title, message: calls.append(("warning", title, message)),
            runtime_imei=lambda: "",
            request_device_imei=lambda: calls.append(("imei",)),
            lock=self.FakeLock(),
            get_thread=lambda: None,
            set_thread=stored.append,
            stop_event=stop_event,
            thread_factory=self.FakeThread,
            thread_target=lambda url, interval: calls.append(("target", url, interval)),
        )

        self.assertTrue(ok)
        self.assertTrue(stop_event.cleared)
        self.assertEqual(calls, [("imei",), ("target", "ws://host/websocket", 7)])
        self.assertEqual(len(stored), 1)
        self.assertTrue(stored[0].started)
        self.assertTrue(stored[0].daemon)

    def test_start_cloud_control_runtime_returns_true_when_already_running(self):
        calls = []
        thread = self.FakeThread(alive=True)

        ok = start_cloud_control_runtime(
            websockets_available=True,
            url="ws://host",
            device_secret="secret",
            reconnect_interval=7,
            show_errors=False,
            validate_start=validate_cloud_start,
            set_cloud_status=lambda text, color: calls.append(("status", text, color)),
            log_missing_dependency=lambda: calls.append(("missing",)),
            show_warning=lambda title, message: calls.append(("warning", title, message)),
            runtime_imei=lambda: "imei",
            request_device_imei=lambda: calls.append(("imei",)),
            lock=self.FakeLock(),
            get_thread=lambda: thread,
            set_thread=lambda new_thread: calls.append(("thread", new_thread)),
            stop_event=self.FakeEvent(False),
            thread_factory=self.FakeThread,
            thread_target=lambda *_: calls.append(("target",)),
        )

        self.assertTrue(ok)
        self.assertEqual(calls, [])

    def test_stop_cloud_control_runtime_resets_state_and_schedules_unregister(self):
        calls = []

        class FakeLoop:
            def is_running(self):
                return True

        class FakeStopEvent:
            def set(self):
                calls.append(("stop",))

        async def unregister(ws):
            return ws

        def run_coro(coro, loop):
            calls.append(("coro", loop))
            coro.close()
            return "future"

        stop_cloud_control_runtime(
            update_status=True,
            enabled=True,
            stop_event=FakeStopEvent(),
            set_connected=lambda value: calls.append(("connected", value)),
            set_authorized=lambda value: calls.append(("authorized", value)),
            reset_serial_log_state=lambda: calls.append(("reset_log",)),
            get_loop=lambda: FakeLoop(),
            get_ws=lambda: "ws",
            schedule_unregister_then_close=unregister,
            set_ws=lambda value: calls.append(("ws", value)),
            set_cloud_status=lambda text, color: calls.append(("status", text, color)),
            run_coroutine_threadsafe=run_coro,
        )

        self.assertEqual(calls[0], ("stop",))
        self.assertIn(("connected", False), calls)
        self.assertIn(("authorized", False), calls)
        self.assertIn(("reset_log",), calls)
        self.assertIn(("ws", None), calls)
        self.assertEqual(calls[-1][0], "status")
        self.assertEqual(calls[-1][2], "#666666")

    def test_stop_cloud_control_runtime_skips_unregister_and_status_when_unavailable(self):
        calls = []

        class FakeStopEvent:
            def set(self):
                calls.append(("stop",))

        stop_cloud_control_runtime(
            update_status=False,
            enabled=False,
            stop_event=FakeStopEvent(),
            set_connected=lambda value: calls.append(("connected", value)),
            set_authorized=lambda value: calls.append(("authorized", value)),
            reset_serial_log_state=lambda: calls.append(("reset_log",)),
            get_loop=lambda: None,
            get_ws=lambda: "ws",
            schedule_unregister_then_close=lambda ws: calls.append(("unregister", ws)),
            set_ws=lambda value: calls.append(("ws", value)),
            set_cloud_status=lambda text, color: calls.append(("status", text, color)),
            run_coroutine_threadsafe=lambda coro, loop: calls.append(("coro", loop)),
        )

        self.assertIn(("ws", None), calls)
        self.assertNotIn(("status", "unused", "unused"), calls)
        self.assertFalse(any(call[0] == "coro" for call in calls))

    def test_restart_cloud_control_runtime_stops_and_restarts(self):
        calls = []
        restart_seq = [0]
        old_thread = self.FakeThread(alive=True)

        def increment_restart_seq():
            restart_seq[0] += 1
            return restart_seq[0]

        restart_thread = restart_cloud_control_runtime(
            show_errors=True,
            lock=self.FakeLock(),
            increment_restart_seq=increment_restart_seq,
            get_restart_seq=lambda: restart_seq[0],
            get_thread=lambda: old_thread,
            stop_control=lambda **kwargs: calls.append(("stop", kwargs)),
            tk_alive=lambda: True,
            stop_event=self.FakeEvent(False),
            set_cloud_status=lambda text, color: calls.append(("status", text, color)),
            schedule_after=lambda delay, callback: calls.append(("after", delay, callback)),
            ui_post=lambda callback: callback(),
            start_control=lambda **kwargs: calls.append(("start", kwargs)),
            thread_factory=self.FakeThread,
        )

        self.assertTrue(restart_thread.started)
        self.assertEqual(old_thread.joined, [2.0])
        self.assertIn(("stop", {"update_status": False}), calls)
        self.assertIn(("start", {"show_errors": True}), calls)


if __name__ == "__main__":
    unittest.main()
