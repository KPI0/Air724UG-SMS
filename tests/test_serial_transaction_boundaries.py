import threading
import unittest
from unittest.mock import patch

from sms_core import serial_sender
from sms_core.cloud_message_namespace_runtime import send_cloud_serial_command_namespace_runtime


class SyntheticModem:
    def __init__(self, at, sms, *, command_reply="OK"):
        self.at = at
        self.sms = sms
        self.command_reply = command_reply
        self.writes = []
        self.is_open = True

    def write(self, payload):
        self.writes.append(payload)
        if payload.startswith(b"AT+CMGS="):
            self.sms.observe_line(">", connection=self)
        elif payload.endswith(b"\x1a"):
            self.sms.observe_line("+CMGS: 1", connection=self)
            self.sms.observe_line("OK", connection=self)
        else:
            self.at.observe_line(self.command_reply, connection=self)
        return len(payload)

    def flush(self):
        pass


class FinalWriteGate:
    def __init__(self, pause_write=1):
        self.lock = serial_sender.DEFAULT_SERIAL_WRITE_LOCK
        self.reached = threading.Event()
        self.resume = threading.Event()
        self.pause_write = pause_write
        self.count = 0

    def __enter__(self):
        self.count += 1
        if self.count == self.pause_write:
            self.reached.set()
            if not self.resume.wait(3):
                raise RuntimeError("Test did not release the final write lock")
        return self.lock.__enter__()

    def __exit__(self, *args):
        return self.lock.__exit__(*args)


class SerialTransactionBoundaryTests(unittest.TestCase):
    def run_paused_write(self, operation, while_waiting, pause_write=1):
        gate = FinalWriteGate(pause_write)
        results, errors = [], []

        def run():
            try:
                results.append(operation())
            except Exception as exc:
                errors.append(exc)

        with patch.object(serial_sender, "DEFAULT_SERIAL_WRITE_LOCK", gate):
            worker = threading.Thread(target=run)
            worker.start()
            try:
                self.assertTrue(gate.reached.wait(3), errors)
                while_waiting()
            finally:
                gate.resume.set()
                worker.join(3)
        self.assertFalse(worker.is_alive())
        self.assertEqual(errors, [])
        self.assertEqual(len(results), 1)
        return results[0]

    def test_previous_reply_before_current_write_cannot_complete_current_command(self):
        for previous, current in (("OK", "ERROR"), ("ERROR", "OK")):
            with self.subTest(previous=previous, current=current):
                at = serial_sender.AtCommandResponseCoordinator()
                modem = SyntheticModem(at, serial_sender.SmsPduSendCoordinator(), command_reply=current)

                def old_response():
                    self.assertEqual(modem.writes, [])
                    at.observe_line("[I]-[ril.proatc] " + previous, connection=modem)

                result = self.run_paused_write(
                    lambda: serial_sender.write_serial_command_confirmed_locked(
                        threading.RLock(), lambda: modem, "ATI", response_coordinator=at,
                        response_timeout=0.1,
                    ),
                    old_response,
                )
                self.assertEqual(result.ok, current == "OK")
                self.assertEqual(modem.writes, [b"ATI\r\n"])
                self.assertIsNone(at._active)

    def test_cancelled_waiter_does_not_write_after_final_lock_wait(self):
        at = serial_sender.AtCommandResponseCoordinator()
        modem = SyntheticModem(at, serial_sender.SmsPduSendCoordinator())
        result = self.run_paused_write(
            lambda: serial_sender.write_serial_command_confirmed_locked(
                threading.RLock(), lambda: modem, "ATI", response_coordinator=at,
                response_timeout=0.1,
            ),
            lambda: at.cancel_active("synthetic cancellation"),
        )
        self.assertFalse(result.ok)
        self.assertEqual(modem.writes, [])
        self.assertIsNone(at._active)

    def test_immediate_reader_reply_can_complete_while_serial_flush_waits(self):
        at = serial_sender.AtCommandResponseCoordinator()
        observed = threading.Event()
        readers = []

        class ConcurrentReplyModem(SyntheticModem):
            def write(self, payload):
                self.writes.append(payload)

                def read_reply():
                    at.observe_line("OK", connection=self)
                    observed.set()

                reader = threading.Thread(target=read_reply)
                readers.append(reader)
                reader.start()
                return len(payload)

            def flush(self):
                if not observed.wait(1):
                    raise RuntimeError("Reader was blocked by the writer")

        modem = ConcurrentReplyModem(at, serial_sender.SmsPduSendCoordinator())
        try:
            result = serial_sender.write_serial_command_confirmed_locked(
                threading.RLock(), lambda: modem, "ATI", response_coordinator=at,
                response_timeout=0.1,
            )
        finally:
            for reader in readers:
                reader.join(2)
                self.assertFalse(reader.is_alive())
        self.assertTrue(result.ok, result.error)
        self.assertEqual(modem.writes, [b"ATI\r\n"])

    def exercise_cloud_context_change(self, command, metadata, pause_write, change):
        at = serial_sender.AtCommandResponseCoordinator()
        sms = serial_sender.SmsPduSendCoordinator()
        modem = SyntheticModem(at, sms)
        namespace = {
            "serial_lock": threading.RLock(), "serial_obj": modem,
            "serial_connection_generation": 1,
            "CLOUD_DEVICE_SECRET": "synthetic-only",
            "CLOUD_WS_URL": "wss://example.invalid/ws/device",
            "CLOUD_CONTROL_ENABLED": True, "cloud_ws_conn": object(),
            "cloud_connected": True, "cloud_device_authorized": True,
            "_cloud_runtime_imei": lambda: "000000000000001",
            "SMS_SEND_COORDINATOR": sms, "SERIAL_COMMAND_RESPONSE_COORDINATOR": at,
            "CLOUD_SENSITIVE_COMMAND_PERMISSIONS": {"sms": True, "phone_number": True},
            "_cloud_log": lambda *args, **kwargs: None,
        }
        offset = []

        def change_context():
            offset.append(len(modem.writes))
            if change == "disabled":
                namespace["CLOUD_CONTROL_ENABLED"] = False
            elif change == "authorization":
                namespace["cloud_device_authorized"] = False
            elif change == "generation":
                namespace["serial_connection_generation"] += 1
            elif change == "websocket":
                namespace["cloud_ws_conn"] = object()
            elif change == "secret":
                namespace["CLOUD_DEVICE_SECRET"] = "changed-synthetic-only"

        result = self.run_paused_write(
            lambda: send_cloud_serial_command_namespace_runtime(namespace, command, metadata),
            change_context, pause_write,
        )
        if change == "unchanged":
            self.assertTrue(result[0], result[1])
            self.assertGreater(len(modem.writes), offset[0])
        else:
            self.assertFalse(result[0], "Changed cloud context reported success")
            self.assertEqual(modem.writes[offset[0]:], [], "Command was committed after context changed")
        self.assertIsNone(at._active)
        self.assertIsNone(sms._active)

    def test_cloud_context_rechecked_at_each_final_write(self):
        transactions = (
            ("ATI", {}, 1),
            ("CLOUD:SEND_SMS", {"sms_phone": "10086", "sms_message": "synthetic-only"}, 3),
            ("CLOUD:SET_OWN_NUMBER", {"own_number": "+19990000001"}, 3),
        )
        for command, metadata, write_count in transactions:
            for pause_write in range(1, write_count + 1):
                for change in ("disabled", "authorization", "generation", "websocket", "secret", "unchanged"):
                    with self.subTest(command=command, pause_write=pause_write, change=change):
                        self.exercise_cloud_context_change(command, metadata, pause_write, change)


if __name__ == "__main__":
    unittest.main()
