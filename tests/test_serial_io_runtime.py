import threading
import unittest

from sms_core.serial_io_runtime import (
    read_serial_line_safely_runtime,
    safe_close_serial_runtime,
    send_call_hangup_runtime,
)


class SerialError(Exception):
    pass


class FakeSerial:
    def __init__(self, *, open=True, line=b"OK\r\n", error=None):
        self.is_open = open
        self.line = line
        self.error = error

    def readline(self):
        if self.error is not None:
            raise self.error
        return self.line


class ClosableSerial:
    def __init__(self, *, error=None):
        self.error = error
        self.closed = False

    def close(self):
        self.closed = True
        if self.error is not None:
            raise self.error


class SerialIoRuntimeTests(unittest.TestCase):
    def test_read_serial_line_safely_runtime_reads_open_serial(self):
        serial_obj = FakeSerial(line=b"+CSQ\r\n")

        result = read_serial_line_safely_runtime(
            threading.Lock(),
            lambda: serial_obj,
            SerialError,
        )

        self.assertEqual(result, b"+CSQ\r\n")

    def test_read_serial_line_safely_runtime_rejects_closed_serial(self):
        with self.assertRaisesRegex(SerialError, "serial_obj is None"):
            read_serial_line_safely_runtime(
                threading.Lock(),
                lambda: FakeSerial(open=False),
                SerialError,
            )

    def test_read_serial_line_safely_runtime_wraps_read_errors(self):
        with self.assertRaisesRegex(SerialError, "并发读取被中断"):
            read_serial_line_safely_runtime(
                threading.Lock(),
                lambda: FakeSerial(error=RuntimeError("boom")),
                SerialError,
            )

    def test_send_call_hangup_runtime_writes_ath_command(self):
        calls = []
        serial_obj = object()

        result = send_call_hangup_runtime(
            threading.Lock(),
            lambda: serial_obj,
            lambda next_serial, command: calls.append((next_serial, command)) or "ok",
        )

        self.assertEqual(result, "ok")
        self.assertEqual(calls, [(serial_obj, "ATH")])

    def test_safe_close_serial_runtime_closes_clears_and_unlocks(self):
        calls = []
        serial_obj = ClosableSerial()

        safe_close_serial_runtime(
            threading.Lock(),
            lambda: serial_obj,
            lambda value: calls.append(("set", value)),
            lambda: calls.append(("unlock",)),
        )

        self.assertTrue(serial_obj.closed)
        self.assertEqual(calls, [("set", None), ("unlock",)])

    def test_safe_close_serial_runtime_clears_and_unlocks_after_close_error(self):
        calls = []
        serial_obj = ClosableSerial(error=RuntimeError("boom"))

        safe_close_serial_runtime(
            threading.Lock(),
            lambda: serial_obj,
            lambda value: calls.append(("set", value)),
            lambda: calls.append(("unlock",)),
        )

        self.assertTrue(serial_obj.closed)
        self.assertEqual(calls, [("set", None), ("unlock",)])


if __name__ == "__main__":
    unittest.main()
