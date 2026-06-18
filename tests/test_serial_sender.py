import threading
import unittest

from sms_core.serial_sender import (
    SerialCommandResult,
    send_command_with_result_async,
    write_text_sms_pdu,
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

    def test_write_text_sms_pdu_sends_long_message_segments(self):
        serial_obj = FakeSerial()
        debug = []

        result = write_text_sms_pdu(
            serial_obj,
            "+1234",
            "A" * 100,
            push_debug=debug.append,
            sleep_func=lambda _seconds: None,
        )

        self.assertTrue(result)
        self.assertEqual(serial_obj.writes[0], b"AT+CMGF=0\r\n")
        cmgs_commands = [
            payload for payload in serial_obj.writes
            if payload.startswith(b"AT+CMGS=")
        ]
        pdu_payloads = [
            payload for payload in serial_obj.writes
            if payload.endswith(b"\x1a")
        ]
        self.assertEqual(len(cmgs_commands), 2)
        self.assertEqual(len(pdu_payloads), 2)
        self.assertTrue(any("(1/2)" in line for line in debug))
        self.assertTrue(any("(2/2)" in line for line in debug))


if __name__ == "__main__":
    unittest.main()
