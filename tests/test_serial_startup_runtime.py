import threading
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
        self.depth = 0

    def __enter__(self):
        self.entered += 1
        self.depth += 1
        return self

    def __exit__(self, exc_type, exc, tb):
        self.exited += 1
        self.depth -= 1
        return False

    @property
    def locked(self):
        return self.depth > 0


class AttemptTrackingRLock:
    def __init__(self):
        self._lock = threading.RLock()
        self._state_lock = threading.Lock()
        self._attempts = 0
        self.second_attempt = threading.Event()

    def __enter__(self):
        with self._state_lock:
            self._attempts += 1
            if self._attempts == 2:
                self.second_attempt.set()
        self._lock.acquire()
        return self

    def __exit__(self, exc_type, exc, tb):
        self._lock.release()
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
            write_command=lambda serial_obj, command: calls.append(
                ("write", serial_obj, command, lock.locked)
            ),
            monotonic=lambda: 10.0,
        )

        self.assertEqual(result, "serial")
        self.assertEqual(lock.entered, 1)
        self.assertEqual(lock.exited, 1)
        self.assertIn(("open", "COM7", 115200), calls)
        self.assertIn(("serial_obj", "serial"), calls)
        self.assertIn(("mutex", "COM7"), calls)
        self.assertIn(("write", "serial", "AT+CLIP=1", False), calls)
        self.assertIn(("deadline", 16.0), calls)
        self.assertIn(("write", "serial", "AT+CGSN", False), calls)
        self.assertIn(("write", "serial", "AT+CNUM", False), calls)
        self.assertIn(("port", "COM7"), calls)

    def test_reconnect_waiting_for_transaction_does_not_hold_serial_lock(self):
        transaction_lock = AttemptTrackingRLock()
        serial_lock = threading.Lock()
        sender_holds_transaction = threading.Event()
        allow_sender_to_continue = threading.Event()
        sender_done = threading.Event()
        startup_done = threading.Event()
        calls = []

        def sender():
            with transaction_lock:
                sender_holds_transaction.set()
                allow_sender_to_continue.wait(1)
                with serial_lock:
                    calls.append(("sender", "serial_lock"))
            sender_done.set()

        def startup():
            open_and_initialize_serial_runtime(
                target_port="COM7",
                baud=115200,
                mode="Auto",
                serial_lock=serial_lock,
                transaction_lock=transaction_lock,
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
            startup_done.set()

        sender_thread = threading.Thread(target=sender, daemon=True)
        startup_thread = threading.Thread(target=startup, daemon=True)
        sender_thread.start()
        self.assertTrue(sender_holds_transaction.wait(1))
        startup_thread.start()
        self.assertTrue(transaction_lock.second_attempt.wait(1))

        acquired = serial_lock.acquire(timeout=0.5)
        self.assertTrue(acquired, "重连等待事务锁时不应占用 serial_lock")
        if acquired:
            serial_lock.release()

        allow_sender_to_continue.set()
        sender_thread.join(1)
        startup_thread.join(1)

        self.assertTrue(sender_done.is_set())
        self.assertTrue(startup_done.is_set())
        self.assertFalse(sender_thread.is_alive())
        self.assertFalse(startup_thread.is_alive())
        self.assertIn(("sender", "serial_lock"), calls)
        self.assertIn(("write", "serial", "AT+CLIP=1"), calls)

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
