import threading
import unittest

from sms_core.serial_sender import (
    SerialCommandResult,
    send_command_with_result_async,
    write_serial_command_sequence_locked,
    write_text_sms_pdu,
    write_text_sms_pdu_locked,
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


class TrackingLock:
    def __init__(self):
        self.depth = 0

    def __enter__(self):
        self.depth += 1
        return self

    def __exit__(self, exc_type, exc, tb):
        self.depth -= 1

    @property
    def locked(self):
        return self.depth > 0


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

    def test_write_serial_command_sequence_locked_releases_lock_while_waiting(self):
        serial_obj = FakeSerial()
        lock = TrackingLock()
        debug = []
        sleep_locked_states = []

        result = write_serial_command_sequence_locked(
            lock,
            lambda: serial_obj,
            ["AT", "ATI"],
            push_debug=debug.append,
            sleep_func=lambda _seconds: sleep_locked_states.append(lock.locked),
        )

        self.assertTrue(result)
        self.assertEqual(serial_obj.writes, [b"AT\r\n", b"ATI\r\n"])
        self.assertEqual(sleep_locked_states, [False])
        self.assertIn(">>> 发送: ATI\\r\\n", debug)

    def test_write_text_sms_pdu_locked_releases_lock_while_waiting(self):
        serial_obj = FakeSerial()
        lock = TrackingLock()
        debug = []
        ui_lines = []
        sleep_calls = []

        result = write_text_sms_pdu_locked(
            lock,
            lambda: serial_obj,
            "+1234",
            "A" * 100,
            push_debug=debug.append,
            port_ui=lambda *args: ui_lines.append(args),
            sleep_func=lambda seconds: sleep_calls.append((seconds, lock.locked)),
        )

        self.assertTrue(result)
        self.assertEqual([locked for _seconds, locked in sleep_calls], [False, False, False, False])
        self.assertEqual([seconds for seconds, _locked in sleep_calls], [0.3, 1.0, 1.5, 1.0])
        self.assertEqual(ui_lines[0], ("📤 发送短信至 +1234：", "normal"))
        self.assertIn(">>> 发送: AT+CMGF=0\\r\\n", debug)


if __name__ == "__main__":
    unittest.main()
