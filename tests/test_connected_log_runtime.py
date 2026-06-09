import unittest

from sms_core.connected_log_runtime import (
    connected_status_update,
    run_delayed_connected_log_runtime,
    startup_ui_delay_ms,
    start_delayed_connected_log_runtime,
)


class ConnectedLogRuntimeTests(unittest.TestCase):
    def test_startup_ui_delay_ms_clamps_elapsed_time(self):
        self.assertEqual(
            startup_ui_delay_ms(10.0, 2.0, monotonic=lambda: 10.5),
            1500,
        )
        self.assertEqual(
            startup_ui_delay_ms(10.0, 2.0, monotonic=lambda: 13.0),
            0,
        )

    def test_connected_status_update_skips_call_states(self):
        for current in ("响铃中", "通话中", "呼叫中"):
            self.assertIsNone(connected_status_update(current, "COM1"))

    def test_connected_status_update_formats_idle_state(self):
        self.assertEqual(
            connected_status_update(
                "等待中",
                "COM3",
                format_connected_status=lambda port: f"connected:{port}",
            ),
            ("connected:COM3", "green"),
        )

    def test_run_delayed_connected_log_runtime_posts_status_update(self):
        calls = []
        posted = []

        def ui_post(fn):
            posted.append(fn)

        def root_after(delay_ms, fn):
            calls.append(("after", delay_ms))
            fn()

        run_delayed_connected_log_runtime(
            "COM7",
            115200,
            delay=2,
            sleep=lambda delay: calls.append(("sleep", delay)),
            reset_auto_connect_state=lambda: calls.append(("reset",)),
            clear_serial_error_repeat_state=lambda: calls.append(("clear",)),
            system_ui=lambda msg: calls.append(("ui", msg)),
            ui_post=ui_post,
            root_after=root_after,
            get_status=lambda: "空闲",
            set_status=lambda text, color: calls.append(("status", text, color)),
            app_start_mono=10.0,
            start_ui_delay=2.0,
            monotonic=lambda: 10.5,
            format_connected_status=lambda port: f"connected:{port}",
        )

        self.assertEqual(calls[:3], [("sleep", 2), ("reset",), ("clear",)])
        self.assertIn(("ui", "🔌 串口已连接：COM7 @ 115200"), calls)
        self.assertEqual(len(posted), 1)

        posted[0]()

        self.assertIn(("after", 1500), calls)
        self.assertIn(("status", "connected:COM7", "green"), calls)

    def test_run_delayed_connected_log_runtime_falls_back_when_after_fails(self):
        calls = []

        def root_after(_delay_ms, _fn):
            raise RuntimeError("after failed")

        run_delayed_connected_log_runtime(
            "COM7",
            115200,
            delay=0,
            sleep=lambda _delay: None,
            reset_auto_connect_state=lambda: None,
            clear_serial_error_repeat_state=lambda: None,
            system_ui=lambda _msg: None,
            ui_post=lambda fn: fn(),
            root_after=root_after,
            get_status=lambda: "空闲",
            set_status=lambda text, color: calls.append((text, color)),
            app_start_mono=0.0,
            start_ui_delay=0.0,
            monotonic=lambda: 0.0,
            format_connected_status=lambda port: f"connected:{port}",
        )

        self.assertEqual(calls, [("connected:COM7", "green")])

    def test_start_delayed_connected_log_runtime_starts_thread(self):
        calls = []

        class FakeThread:
            def __init__(self, target, daemon):
                self.target = target
                calls.append(("thread", daemon))

            def start(self):
                calls.append(("start",))

        thread = start_delayed_connected_log_runtime(
            "COM1",
            9600,
            thread_factory=FakeThread,
            delay=0,
            sleep=lambda _delay: None,
            reset_auto_connect_state=lambda: None,
            clear_serial_error_repeat_state=lambda: None,
            system_ui=lambda _msg: None,
            ui_post=lambda _fn: None,
            root_after=lambda _delay_ms, _fn: None,
            get_status=lambda: "",
            set_status=lambda _text, _color: None,
            app_start_mono=0.0,
            start_ui_delay=0.0,
        )

        self.assertIsInstance(thread, FakeThread)
        self.assertEqual(calls, [("thread", True), ("start",)])


if __name__ == "__main__":
    unittest.main()
