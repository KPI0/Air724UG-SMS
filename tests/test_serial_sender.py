import threading
import unittest

from sms_core.serial_sender import (
    SerialCommandResult,
    send_command_with_result_async,
    write_serial_command_result,
)


class FakeSerial:
    def __init__(self, is_open=True, fail_write=False):
        self.is_open = is_open
        self.fail_write = fail_write
        self.writes = []
        self.flush_count = 0

    def write(self, payload):
        if self.fail_write:
            raise RuntimeError("write failed")
        self.writes.append(payload)

    def flush(self):
        self.flush_count += 1


class SerialSenderResultTests(unittest.TestCase):
    def test_write_serial_command_result_reports_success(self):
        serial_obj = FakeSerial()

        result = write_serial_command_result(serial_obj, "ATA")

        self.assertEqual(result, SerialCommandResult(True))
        self.assertEqual(serial_obj.writes, [b"ATA\r\n"])
        self.assertEqual(serial_obj.flush_count, 1)

    def test_write_serial_command_result_reports_closed_port(self):
        result = write_serial_command_result(FakeSerial(is_open=False), "ATA")

        self.assertFalse(result.ok)
        self.assertEqual(result.error, "串口未连接")

    def test_write_serial_command_result_reports_write_error(self):
        result = write_serial_command_result(FakeSerial(fail_write=True), "ATA")

        self.assertFalse(result.ok)
        self.assertEqual(result.error, "write failed")

    def test_send_command_with_result_async_invokes_callback(self):
        serial_obj = FakeSerial()
        lock = threading.RLock()
        done = threading.Event()
        results = []

        thread = send_command_with_result_async(
            lock,
            lambda: serial_obj,
            "ATH",
            on_result=lambda result: (results.append(result), done.set()),
        )

        thread.join(1)
        self.assertTrue(done.is_set())
        self.assertEqual(results, [SerialCommandResult(True)])
        self.assertEqual(serial_obj.writes, [b"ATH\r\n"])


if __name__ == "__main__":
    unittest.main()
