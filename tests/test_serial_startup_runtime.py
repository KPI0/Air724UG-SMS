import unittest

from sms_core.serial_startup_runtime import (
    open_and_initialize_serial_runtime,
    resolve_serial_target_port_runtime,
)


class FakePort:
    def __init__(self, device):
        self.device = device


class FakeLock:
    def __init__(self):
        self.entered = 0
        self.exited = 0

    def __enter__(self):
        self.entered += 1
        return self

    def __exit__(self, exc_type, exc, tb):
        self.exited += 1
        return False


class SerialStartupRuntimeTests(unittest.TestCase):
    def test_resolve_serial_target_port_uses_luat_port_in_auto_mode(self):
        calls = []

        result = resolve_serial_target_port_runtime(
            mode="Auto",
            current_port="COM1",
            reconnect_interval=2,
            find_luat_best_port=lambda: ("COM7", "LUAT Modem"),
            list_ports=lambda: [],
            is_port_locked=lambda port: False,
            auto_connect_ui=lambda message: calls.append(("auto", message)),
            set_status=lambda *args: calls.append(("status", args)),
            wakeup_wait=lambda **kwargs: calls.append(("wait", kwargs)),
            wakeup_clear=lambda: calls.append(("clear",)),
        )

        self.assertEqual(result, "COM7")
        self.assertEqual(calls[0][0], "auto")

    def test_resolve_serial_target_port_uses_single_unlocked_port(self):
        calls = []

        result = resolve_serial_target_port_runtime(
            mode="Auto",
            current_port="",
            reconnect_interval=2,
            find_luat_best_port=lambda: (None, None),
            list_ports=lambda: [FakePort("COM5")],
            is_port_locked=lambda port: False,
            auto_connect_ui=lambda message: calls.append(("auto", message)),
            set_status=lambda *args: calls.append(("status", args)),
            wakeup_wait=lambda **kwargs: calls.append(("wait", kwargs)),
            wakeup_clear=lambda: calls.append(("clear",)),
        )

        self.assertEqual(result, "COM5")
        self.assertEqual(calls[0][0], "auto")

    def test_resolve_serial_target_port_waits_when_auto_has_no_candidate(self):
        calls = []

        result = resolve_serial_target_port_runtime(
            mode="Auto",
            current_port="",
            reconnect_interval=3,
            find_luat_best_port=lambda: (None, None),
            list_ports=lambda: [FakePort("COM1"), FakePort("COM2")],
            is_port_locked=lambda port: False,
            auto_connect_ui=lambda message: calls.append(("auto", message)),
            set_status=lambda *args: calls.append(("status", args)),
            wakeup_wait=lambda **kwargs: calls.append(("wait", kwargs)),
            wakeup_clear=lambda: calls.append(("clear",)),
        )

        self.assertIsNone(result)
        self.assertEqual(calls[0][0], "status")
        self.assertEqual(calls[1], ("wait", {"timeout": 3}))
        self.assertEqual(calls[2], ("clear",))

    def test_resolve_serial_target_port_reports_missing_manual_port(self):
        calls = []

        result = resolve_serial_target_port_runtime(
            mode="Manual",
            current_port="",
            reconnect_interval=4,
            find_luat_best_port=lambda: (None, None),
            list_ports=lambda: [],
            is_port_locked=lambda port: False,
            auto_connect_ui=lambda message: calls.append(("auto", message)),
            set_status=lambda *args: calls.append(("status", args)),
            wakeup_wait=lambda **kwargs: calls.append(("wait", kwargs)),
            wakeup_clear=lambda: calls.append(("clear",)),
            sleep=lambda seconds: calls.append(("sleep", seconds)),
        )

        self.assertIsNone(result)
        self.assertEqual(calls[0][0], "status")
        self.assertEqual(calls[1], ("sleep", 4))

    def test_open_and_initialize_serial_runtime_initializes_port(self):
        calls = []
        lock = FakeLock()

        result = open_and_initialize_serial_runtime(
            target_port="COM7",
            baud=115200,
            mode="Auto",
            serial_lock=lock,
            open_serial=lambda port, baud: calls.append(("open", port, baud)) or "serial",
            set_serial_obj=lambda value: calls.append(("serial_obj", value)),
            set_port=lambda value: calls.append(("port", value)),
            lock_port_mutex=lambda port: calls.append(("mutex", port)),
            set_cloud_imei_query_deadline=lambda value: calls.append(("deadline", value)),
            serial_error_ui=lambda *args, **kwargs: calls.append(("error", args, kwargs)),
            set_status=lambda *args: calls.append(("status", args)),
            write_command=lambda serial_obj, command: calls.append(("write", serial_obj, command)),
            monotonic=lambda: 10.0,
        )

        self.assertEqual(result, "serial")
        self.assertEqual(lock.entered, 1)
        self.assertEqual(lock.exited, 1)
        self.assertIn(("open", "COM7", 115200), calls)
        self.assertIn(("serial_obj", "serial"), calls)
        self.assertIn(("mutex", "COM7"), calls)
        self.assertIn(("write", "serial", "AT+CLIP=1"), calls)
        self.assertIn(("deadline", 16.0), calls)
        self.assertIn(("write", "serial", "AT+CGSN"), calls)
        self.assertIn(("port", "COM7"), calls)

    def test_open_and_initialize_serial_runtime_reports_denied_port(self):
        calls = []

        with self.assertRaises(RuntimeError):
            open_and_initialize_serial_runtime(
                target_port="COM7",
                baud=115200,
                mode="Manual",
                serial_lock=FakeLock(),
                open_serial=lambda port, baud: (_ for _ in ()).throw(RuntimeError("Access is denied")),
                set_serial_obj=lambda value: calls.append(("serial_obj", value)),
                set_port=lambda value: calls.append(("port", value)),
                lock_port_mutex=lambda port: calls.append(("mutex", port)),
                set_cloud_imei_query_deadline=lambda value: calls.append(("deadline", value)),
                serial_error_ui=lambda *args, **kwargs: calls.append(("error", args, kwargs)),
                set_status=lambda *args: calls.append(("status", args)),
                is_open_denied=lambda text: "denied" in text.lower(),
                open_denied_repeat_key=lambda port: f"repeat:{port}",
            )

        self.assertEqual(calls[0][0], "error")
        self.assertEqual(calls[0][2], {"repeat_key": "repeat:COM7"})
        self.assertEqual(calls[1][0], "status")


if __name__ == "__main__":
    unittest.main()
