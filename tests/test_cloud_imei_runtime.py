import unittest
import re

from sms_core.cloud_imei_runtime import (
    IMEI_READ_COMMAND,
    maybe_capture_cloud_device_imei_runtime,
    notify_cloud_identity_changed_runtime,
    request_cloud_device_imei_runtime,
    request_cloud_device_imei_worker,
    set_cloud_device_imei_runtime,
)
from sms_core.serial_sender import SerialCommandResult


class DummyLock:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


class ImmediateThread:
    def __init__(self, target, daemon):
        self.target = target
        self.daemon = daemon
        self.started = False

    def start(self):
        self.started = True
        self.target()


class FakeLoop:
    def __init__(self, running=True):
        self.running = running

    def is_running(self):
        return self.running


class CloudImeiRuntimeTests(unittest.TestCase):
    def test_notify_cloud_identity_changed_runtime_schedules_register(self):
        calls = []
        loop = FakeLoop()
        ws = object()

        async def send_register(next_ws):
            return next_ws

        result = notify_cloud_identity_changed_runtime(
            get_loop=lambda: loop,
            get_ws=lambda: ws,
            is_connected=lambda: True,
            runtime_imei=lambda: "123456789012345",
            send_register=send_register,
            run_coroutine_threadsafe=lambda coro, next_loop: calls.append((coro, next_loop)),
        )

        self.assertTrue(result)
        self.assertEqual(calls[0][1], loop)
        calls[0][0].close()

    def test_notify_cloud_identity_changed_runtime_skips_unavailable_states(self):
        base = {
            "get_loop": lambda: FakeLoop(),
            "get_ws": lambda: object(),
            "is_connected": lambda: True,
            "runtime_imei": lambda: "123456789012345",
            "send_register": lambda ws: "coro",
            "run_coroutine_threadsafe": lambda *_args: None,
        }

        self.assertFalse(notify_cloud_identity_changed_runtime(**{**base, "get_loop": lambda: None}))
        self.assertFalse(notify_cloud_identity_changed_runtime(**{**base, "get_loop": lambda: FakeLoop(False)}))
        self.assertFalse(notify_cloud_identity_changed_runtime(**{**base, "get_ws": lambda: None}))
        self.assertFalse(notify_cloud_identity_changed_runtime(**{**base, "is_connected": lambda: False}))
        self.assertFalse(notify_cloud_identity_changed_runtime(**{**base, "runtime_imei": lambda: ""}))

    def test_notify_cloud_identity_changed_runtime_closes_coro_when_submit_fails(self):
        created = []

        async def send_register(_ws):
            return None

        def build_register(ws):
            coro = send_register(ws)
            created.append(coro)
            return coro

        result = notify_cloud_identity_changed_runtime(
            get_loop=lambda: FakeLoop(),
            get_ws=lambda: object(),
            is_connected=lambda: True,
            runtime_imei=lambda: "123456789012345",
            send_register=build_register,
            run_coroutine_threadsafe=lambda _coro, _loop: (_ for _ in ()).throw(
                RuntimeError("loop rejected")
            ),
        )

        self.assertFalse(result)
        self.assertEqual(len(created), 1)
        self.assertIsNone(created[0].cr_frame)

    def test_set_cloud_device_imei_runtime_rejects_invalid_length(self):
        calls = []

        result = set_cloud_device_imei_runtime(
            "123",
            current_imei=lambda: "",
            normalize_imei=lambda value: re.sub(r"\D", "", str(value or "")),
            set_device_imei=lambda value: calls.append(("imei", value)),
            set_verified=lambda value: calls.append(("verified", value)),
            log=lambda message: calls.append(("log", message)),
            notify_identity_changed=lambda: calls.append(("notify",)),
        )

        self.assertFalse(result)
        self.assertEqual(calls, [])

    def test_set_cloud_device_imei_runtime_marks_existing_verified(self):
        calls = []

        result = set_cloud_device_imei_runtime(
            "IMEI: 123456789012345",
            current_imei=lambda: "123456789012345",
            normalize_imei=lambda value: re.sub(r"\D", "", str(value or "")),
            set_device_imei=lambda value: calls.append(("imei", value)),
            set_verified=lambda value: calls.append(("verified", value)),
            log=lambda message: calls.append(("log", message)),
            notify_identity_changed=lambda: calls.append(("notify",)),
        )

        self.assertTrue(result)
        self.assertEqual(calls, [("verified", True)])

    def test_set_cloud_device_imei_runtime_updates_logs_and_notifies(self):
        calls = []

        result = set_cloud_device_imei_runtime(
            "IMEI: 123456789012345",
            current_imei=lambda: "",
            normalize_imei=lambda value: re.sub(r"\D", "", str(value or "")),
            set_device_imei=lambda value: calls.append(("imei", value)),
            set_verified=lambda value: calls.append(("verified", value)),
            log=lambda message: calls.append(("log", message)),
            notify_identity_changed=lambda: calls.append(("notify",)),
        )

        self.assertTrue(result)
        self.assertEqual(calls[0], ("imei", "123456789012345"))
        self.assertEqual(calls[1], ("verified", True))
        self.assertIn("123456789012345", calls[2][1])
        self.assertEqual(calls[3], ("notify",))

    def test_maybe_capture_cloud_device_imei_runtime_ignores_inactive_and_echo(self):
        regex = re.compile(r"\b(\d{14,17})\b")
        calls = []

        self.assertEqual(
            maybe_capture_cloud_device_imei_runtime(
                "123456789012345",
                query_deadline=0.0,
                set_query_deadline=lambda value: calls.append(("deadline", value)),
                imei_regex=regex,
                set_device_imei=lambda *_args, **_kwargs: True,
                monotonic=lambda: 1.0,
            ),
            "inactive",
        )
        self.assertEqual(
            maybe_capture_cloud_device_imei_runtime(
                "AT+CGSN",
                query_deadline=10.0,
                set_query_deadline=lambda value: calls.append(("deadline", value)),
                imei_regex=regex,
                set_device_imei=lambda *_args, **_kwargs: True,
                monotonic=lambda: 1.0,
            ),
            "ignored",
        )
        self.assertEqual(calls, [])

    def test_maybe_capture_cloud_device_imei_runtime_expires_deadline(self):
        calls = []

        result = maybe_capture_cloud_device_imei_runtime(
            "123456789012345",
            query_deadline=10.0,
            set_query_deadline=lambda value: calls.append(("deadline", value)),
            imei_regex=re.compile(r"\b(\d{14,17})\b"),
            set_device_imei=lambda *_args, **_kwargs: True,
            monotonic=lambda: 11.0,
        )

        self.assertEqual(result, "expired")
        self.assertEqual(calls, [("deadline", 0.0)])

    def test_maybe_capture_cloud_device_imei_runtime_captures_imei(self):
        calls = []

        result = maybe_capture_cloud_device_imei_runtime(
            "IMEI: 123456789012345",
            query_deadline=10.0,
            set_query_deadline=lambda value: calls.append(("deadline", value)),
            imei_regex=re.compile(r"\b(\d{14,17})\b"),
            set_device_imei=lambda imei, source: calls.append(("imei", imei, source)) or True,
            monotonic=lambda: 1.0,
        )

        self.assertEqual(result, "captured")
        self.assertEqual(calls, [("imei", "123456789012345", IMEI_READ_COMMAND), ("deadline", 0.0)])

    def test_maybe_capture_cloud_device_imei_runtime_reports_no_match_and_unchanged(self):
        regex = re.compile(r"\b(\d{14,17})\b")

        self.assertEqual(
            maybe_capture_cloud_device_imei_runtime(
                "short 123",
                query_deadline=10.0,
                set_query_deadline=lambda value: None,
                imei_regex=regex,
                set_device_imei=lambda *_args, **_kwargs: True,
                monotonic=lambda: 1.0,
            ),
            "no_match",
        )
        self.assertEqual(
            maybe_capture_cloud_device_imei_runtime(
                "123456789012345",
                query_deadline=10.0,
                set_query_deadline=lambda value: None,
                imei_regex=regex,
                set_device_imei=lambda *_args, **_kwargs: False,
                monotonic=lambda: 1.0,
            ),
            "unchanged",
        )

    def test_request_cloud_device_imei_worker_sends_command_and_sets_deadline(self):
        calls = []

        ok = request_cloud_device_imei_worker(
            serial_lock=DummyLock(),
            get_serial=lambda: "serial",
            write_command_result=lambda serial_obj, command: calls.append(("write", serial_obj, command)) or SerialCommandResult(True),
            set_query_deadline=lambda deadline: calls.append(("deadline", deadline)),
            monotonic=lambda: 10.0,
            push_serial_debug=lambda line: calls.append(("debug", line)),
            cloud_log=lambda message: calls.append(("log", message)),
        )

        self.assertTrue(ok)
        self.assertIn(("write", "serial", IMEI_READ_COMMAND), calls)
        self.assertIn(("deadline", 16.0), calls)
        self.assertTrue(any(item[0] == "debug" for item in calls))
        self.assertTrue(any(item[0] == "log" and "AT+CGSN" in item[1] for item in calls))

    def test_request_cloud_device_imei_worker_logs_write_failure(self):
        calls = []

        ok = request_cloud_device_imei_worker(
            serial_lock=DummyLock(),
            get_serial=lambda: "serial",
            write_command_result=lambda serial_obj, command: SerialCommandResult(False, "closed"),
            set_query_deadline=lambda deadline: calls.append(("deadline", deadline)),
            cloud_log=lambda message: calls.append(("log", message)),
        )

        self.assertFalse(ok)
        self.assertEqual(calls, [("log", "读取IMEI失败：closed")])

    def test_request_cloud_device_imei_runtime_starts_worker_thread(self):
        calls = []

        result = request_cloud_device_imei_runtime(
            serial_lock=DummyLock(),
            get_serial=lambda: "serial",
            write_command_result=lambda serial_obj, command: SerialCommandResult(True),
            set_query_deadline=lambda deadline: calls.append(("deadline", deadline)),
            cloud_log=lambda message: calls.append(("log", message)),
            monotonic=lambda: 1.0,
            thread_factory=ImmediateThread,
        )

        self.assertEqual(result, (True, "已尝试发送读取IMEI指令"))
        self.assertIn(("deadline", 7.0), calls)

    def test_request_cloud_device_imei_runtime_reports_thread_start_failure(self):
        class FailingThread:
            def __init__(self, target, daemon):
                pass

            def start(self):
                raise RuntimeError("boom")

        logs = []

        ok, message = request_cloud_device_imei_runtime(
            serial_lock=DummyLock(),
            get_serial=lambda: "serial",
            write_command_result=lambda serial_obj, command: SerialCommandResult(True),
            set_query_deadline=lambda deadline: None,
            cloud_log=logs.append,
            thread_factory=FailingThread,
        )

        self.assertFalse(ok)
        self.assertIn("boom", message)
        self.assertTrue(any("线程启动失败" in item for item in logs))


if __name__ == "__main__":
    unittest.main()
