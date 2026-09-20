"""IMEI query ordering with synthetic serial replies and bounded threads."""

import re
import threading
import unittest
from types import SimpleNamespace

from sms_core.cloud_imei_runtime import request_cloud_device_imei_worker
from sms_core.cloud_state_namespace_runtime import maybe_capture_cloud_device_imei_namespace_runtime
from sms_core.serial_io_runtime import read_serial_line_safely_runtime
from sms_core.serial_sender import SerialCommandResult, write_serial_command_result


class FakeSerial:
    is_open = True

    def write(self, data):
        return len(data)

    def flush(self):
        pass


class CloudImeiQueryOrderingTests(unittest.TestCase):
    def setUp(self):
        self.clock = 100.0
        self.captured = []
        self.namespace = {
            "serial_lock": threading.Lock(),
            "serial_read_lock": threading.Lock(),
            "serial_obj": FakeSerial(),
            "cloud_imei_query_deadline": 0.0,
            "IMEI_REGEX": re.compile(r"\b(\d{14,17})\b"),
            "time": SimpleNamespace(monotonic=lambda: self.clock),
            "_set_cloud_device_imei": lambda value, **_kwargs: self.captured.append(value) or True,
        }

    def query(self, **overrides):
        kwargs = {
            "serial_lock": self.namespace["serial_lock"],
            "get_serial": lambda: self.namespace["serial_obj"],
            "write_command_result": write_serial_command_result,
            "set_query_deadline": lambda value: self.namespace.__setitem__("cloud_imei_query_deadline", value),
            "monotonic": lambda: self.clock,
            "cloud_log": lambda _message: None,
        }
        kwargs.update(overrides)
        return request_cloud_device_imei_worker(**kwargs)

    def capture(self):
        return maybe_capture_cloud_device_imei_namespace_runtime(self.namespace, "000000000000001")

    def test_reply_during_flush_is_captured_and_does_not_reopen_window(self):
        sent, processed = threading.Event(), threading.Event()
        decisions, errors = [], []

        class ReplySerial(FakeSerial):
            def write(self, data):
                sent.set()
                return len(data)

            def readline(self):
                if not sent.wait(2):
                    raise RuntimeError("synthetic command was not written")
                return b"000000000000001\r\n"

            def flush(self):
                if not processed.wait(2):
                    raise RuntimeError("synthetic reader could not process reply")

        self.namespace["serial_obj"] = ReplySerial()

        def read_reply():
            try:
                line = read_serial_line_safely_runtime(
                    self.namespace["serial_lock"],
                    lambda: self.namespace["serial_obj"],
                    RuntimeError,
                    read_lock=self.namespace["serial_read_lock"],
                ).decode("ascii").strip()
                decisions.append(maybe_capture_cloud_device_imei_namespace_runtime(self.namespace, line))
            except Exception as exc:
                errors.append(type(exc).__name__)
            finally:
                processed.set()

        reader = threading.Thread(target=read_reply)
        reader.start()
        try:
            accepted = self.query()
        finally:
            reader.join(timeout=3)
        self.assertFalse(reader.is_alive())
        self.assertEqual(errors, [])
        self.assertTrue(accepted)
        self.assertEqual(decisions, ["captured"])
        self.assertEqual(len(self.captured), 1)
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)

    def test_delayed_reply_consumes_the_query_window(self):
        self.assertTrue(self.query())
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 106.0)
        self.clock = 103.0
        self.assertEqual(self.capture(), "captured")
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)

    def test_write_failure_or_exception_clears_its_window(self):
        for raises in (False, True):
            with self.subTest(raises=raises):
                observed = []

                def fail(_serial, _command):
                    observed.append(self.namespace["cloud_imei_query_deadline"])
                    if raises:
                        raise RuntimeError("synthetic failure")
                    return SerialCommandResult(False, "synthetic failure")

                self.assertFalse(self.query(write_command_result=fail))
                self.assertEqual(observed, [106.0])
                self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)

    def test_failure_does_not_clear_a_newer_query_window(self):
        def fail_after_replacement(_serial, _command):
            self.namespace["cloud_imei_query_deadline"] = 206.0
            return SerialCommandResult(False, "synthetic failure")

        self.assertFalse(self.query(
            write_command_result=fail_after_replacement,
            get_query_deadline=lambda: self.namespace["cloud_imei_query_deadline"],
        ))
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 206.0)

    def test_queued_query_starts_window_after_locks_and_uses_current_serial(self):
        for dependency in ("transaction_lock", "write_lock"):
            with self.subTest(dependency=dependency):
                self.namespace["cloud_imei_query_deadline"] = 0.0
                self.clock = 100.0
                current_serial = FakeSerial()
                owner = self

                class DeferredLock:
                    def __enter__(self):
                        owner.assertEqual(owner.namespace["cloud_imei_query_deadline"], 0.0)
                        owner.clock = 150.0
                        owner.namespace["serial_obj"] = current_serial
                        return self

                    def __exit__(self, *_args):
                        return False

                observed = []

                def write(serial, _command):
                    observed.append((serial is current_serial, self.namespace["cloud_imei_query_deadline"]))
                    return SerialCommandResult(True)

                self.assertTrue(self.query(write_command_result=write, **{dependency: DeferredLock()}))
                self.assertEqual(observed, [(True, 156.0)])

    def test_capture_serializes_deadline_read_and_completion(self):
        observed = []

        def capture(_line, **kwargs):
            observed.append(self.namespace["serial_lock"].locked())
            kwargs["set_query_deadline"](0.0)
            return "captured"

        self.namespace["cloud_imei_query_deadline"] = 106.0
        maybe_capture_cloud_device_imei_namespace_runtime(
            self.namespace, "000000000000001", capture_runtime=capture
        )
        self.assertEqual(observed, [True])

    def test_failed_query_finishes_before_a_concurrent_query_opens_its_window(self):
        first_writing, second_waiting, release_first = (
            threading.Event(), threading.Event(), threading.Event()
        )
        owner = self

        class OrderedLock:
            def __init__(self):
                self.lock = threading.RLock()

            def __enter__(self):
                if threading.current_thread().name == "audit-imei-second":
                    second_waiting.set()
                self.lock.acquire()
                return self

            def __exit__(self, *_args):
                self.lock.release()

        lock = OrderedLock()
        results, observed = {}, []
        ticks = iter((100.0, 200.0))

        def write(_serial, _command):
            observed.append(owner.namespace["cloud_imei_query_deadline"])
            if threading.current_thread().name == "audit-imei-first":
                first_writing.set()
                if not release_first.wait(2):
                    raise RuntimeError("synthetic first query was not released")
                return SerialCommandResult(False, "synthetic first failure")
            return SerialCommandResult(True)

        def run(name):
            results[name] = self.query(
                write_command_result=write,
                transaction_lock=lock,
                monotonic=lambda: next(ticks),
                get_query_deadline=lambda: self.namespace["cloud_imei_query_deadline"],
            )

        first = threading.Thread(target=run, args=("first",), name="audit-imei-first")
        second = threading.Thread(target=run, args=("second",), name="audit-imei-second")
        first.start()
        second_started = False
        try:
            self.assertTrue(first_writing.wait(2))
            second.start()
            second_started = True
            self.assertTrue(second_waiting.wait(2))
            self.assertEqual(observed, [106.0])
        finally:
            release_first.set()
            first.join(timeout=3)
            if second_started:
                second.join(timeout=3)
        self.assertFalse(first.is_alive())
        self.assertFalse(second.is_alive())
        self.assertEqual(results, {"first": False, "second": True})
        self.assertEqual(observed, [106.0, 206.0])
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 206.0)


if __name__ == "__main__":
    unittest.main()
