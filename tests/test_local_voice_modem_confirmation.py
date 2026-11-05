import threading
import unittest

from sms_core import serial_sender
from sms_core.call_send_namespace_runtime import send_local_serial_command_namespace_runtime
from tests.test_local_outgoing_calls import ModemSerial, namespace_fixture


class LocalVoiceModemConfirmationTests(unittest.TestCase):
    commands = ("ATA", "ATH", "ATD10086;")

    def make_namespace(self, response=None, stage="normal", error_message="synthetic-private-error-value"):
        namespace = namespace_fixture(response)
        namespace["ring_timeout_target"] = 0.0
        messages, debug = [], []
        namespace["port_ui"] = lambda message, _kind: messages.append(message)
        namespace["_push_serial_debug"] = debug.append
        modem = namespace["serial_obj"]
        coordinator = namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"]
        original_write = modem.write

        def write(payload):
            if stage == "write-error":
                modem.writes.append(payload)
                raise OSError(error_message)
            written = original_write(payload)
            if stage == "partial":
                return 1
            if stage == "cancel":
                coordinator.cancel_active("synthetic cancellation")
            return written

        def flush():
            if stage == "flush-error":
                raise OSError(error_message)

        modem.write, modem.flush = write, flush
        return namespace, modem, coordinator, messages, debug

    def run_voice(self, namespace, command, before_worker=None):
        def start(_name, target, **_kwargs):
            if before_worker:
                before_worker()
            return target()

        return send_local_serial_command_namespace_runtime(
            namespace, command, start_worker=start, response_timeout=0.002,
        )

    def test_timeout_cannot_let_late_terminal_complete_the_next_command(self):
        for command in self.commands:
            with self.subTest(command=command):
                namespace, modem, coordinator, messages, _debug = self.make_namespace()
                self.run_voice(namespace, command)
                modem.after_write = lambda: coordinator.observe_line("[I]-[ril.proatc] OK", connection=modem)
                second = serial_sender.write_serial_command_confirmed_locked(
                    namespace["serial_lock"], lambda: modem, "ATI",
                    response_coordinator=coordinator, response_timeout=0.002,
                )
                self.assertFalse(second.ok)
                self.assertEqual(modem.writes, [(command + "\r\n").encode("ascii")])
                self.assertTrue(modem.is_open, "Keep the reader alive for late results and recording")
                self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())
                self.assertIsNone(coordinator._active)
                self.assertTrue(any("重启设备" in message for message in messages))

    def test_unknown_write_results_block_further_writes_without_sending_escape(self):
        for command in self.commands:
            for stage in ("write-error", "partial", "flush-error", "cancel"):
                with self.subTest(command=command, stage=stage):
                    namespace, modem, coordinator, messages, _debug = self.make_namespace(stage=stage)
                    self.run_voice(namespace, command)
                    self.assertTrue(serial_sender.serial_sms_recovery_error(modem))
                    self.assertFalse(serial_sender.write_serial_command_result(modem, "ATI").ok)
                    self.assertEqual(modem.writes, [(command + "\r\n").encode("ascii")])
                    self.assertNotIn(b"\x1b", modem.writes)
                    self.assertIsNone(coordinator._active)
                    self.assertTrue(any("重启设备" in message for message in messages))
                    self.assertNotIn("synthetic-private-error-value", "\n".join(messages))

    def test_sensitive_transport_exception_is_hidden_in_local_voice_debug(self):
        sensitive = (
            "CLOUD_DEVICE_SECRET=synthetic-only-password; PIN=123456; "
            "wss://operator:synthetic-only-password@example.invalid/ws"
        )
        for command in self.commands:
            for stage in ("write-error", "flush-error"):
                with self.subTest(command=command, stage=stage):
                    namespace, modem, coordinator, messages, debug = self.make_namespace(
                        stage=stage, error_message=sensitive,
                    )
                    self.run_voice(namespace, command)
                    self.assertIn(
                        ">>> 发送失败: 通话指令未完成，请检查串口状态", debug,
                    )
                    combined = "\n".join(messages + debug)
                    for private in ("synthetic-only-password", "123456", "example.invalid"):
                        self.assertNotIn(private, combined)
                    self.assertTrue(serial_sender.serial_sms_recovery_error(modem))
                    self.assertIsNone(coordinator._active)

    def test_successful_local_voice_debug_keeps_original_command_and_line_ending(self):
        for command in self.commands:
            with self.subTest(command=command):
                namespace, modem, coordinator, messages, debug = self.make_namespace(response="OK")
                self.run_voice(namespace, command)
                self.assertEqual(debug, [">>> 发送: " + command + r"\r\n"])
                self.assertEqual(modem.writes, [(command + "\r\n").encode("ascii")])
                self.assertEqual(messages, [])
                self.assertIsNone(coordinator._active)

    def test_trusted_terminal_during_write_survives_flush_failure(self):
        for command in self.commands:
            for response in ("OK", "ERROR"):
                with self.subTest(command=command, response=response):
                    namespace, modem, coordinator, messages, _debug = self.make_namespace(
                        response, "flush-error",
                    )
                    self.run_voice(namespace, command)
                    self.assertFalse(serial_sender.serial_sms_recovery_error(modem))
                    expected_history = 1 if command.startswith("ATD") and response == "OK" else 0
                    self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].qsize(), expected_history)
                    self.assertIsNone(coordinator._active)
                    self.assertFalse(any("重启设备" in message for message in messages))

    def test_trusted_rejection_and_unanswered_dial_keep_same_serial_usable(self):
        for response in ("ERROR", "+CME ERROR: 3", "BUSY", "NO ANSWER", "NO CARRIER"):
            with self.subTest(response=response):
                namespace, modem, coordinator, _messages, _debug = self.make_namespace(response)
                self.run_voice(namespace, "ATD10086;")
                self.assertFalse(serial_sender.serial_sms_recovery_error(modem))
                expected_history = int(response in ("BUSY", "NO ANSWER", "NO CARRIER"))
                self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].qsize(), expected_history)
                modem.response = "OK"
                self.assertTrue(serial_sender.write_serial_command_confirmed_locked(
                    namespace["serial_lock"], lambda: modem, "ATI",
                    response_coordinator=coordinator, response_timeout=0.002,
                ).ok)

    def test_cancelled_before_write_does_not_isolate_original_serial(self):
        for command in self.commands:
            with self.subTest(command=command):
                namespace, modem, coordinator, _messages, _debug = self.make_namespace()
                self.run_voice(namespace, command, before_worker=lambda: namespace.update(
                    serial_connection_generation=2,
                ))
                self.assertEqual(modem.writes, [])
                self.assertFalse(serial_sender.serial_sms_recovery_error(modem))
                self.assertIsNone(coordinator._active)

    def test_timeout_isolates_only_original_connection(self):
        for command in self.commands:
            with self.subTest(command=command):
                namespace, modem, coordinator, _messages, _debug = self.make_namespace()
                replacement = ModemSerial(coordinator, response="OK")
                modem.after_write = lambda: namespace.update(serial_obj=replacement, serial_connection_generation=2)
                self.run_voice(namespace, command)
                self.assertTrue(serial_sender.serial_sms_recovery_error(modem))
                self.assertFalse(serial_sender.serial_sms_recovery_error(replacement))
                self.assertTrue(serial_sender.write_serial_command_confirmed_locked(
                    namespace["serial_lock"], lambda: replacement, "ATI",
                    response_coordinator=coordinator, response_timeout=0.002,
                ).ok)
                self.assertEqual(modem.writes, [(command + "\r\n").encode("ascii")])
                self.assertEqual(replacement.writes, [b"ATI\r\n"])

    def test_timeout_isolation_completes_before_queued_transaction_can_write(self):
        namespace, modem, coordinator, _messages, _debug = self.make_namespace()
        wrote = threading.Event()
        modem.after_write = wrote.set
        original = send_local_serial_command_namespace_runtime(
            namespace, "ATH", response_timeout=0.1,
        )
        following_results, following_done = [], threading.Event()

        def following():
            following_results.append(serial_sender.write_serial_command_result(modem, "ATI"))
            following_done.set()

        next_worker = threading.Thread(target=following)
        try:
            self.assertTrue(wrote.wait(1))
            next_worker.start()
            self.assertFalse(following_done.wait(0.01))
            original.join(2)
            next_worker.join(2)
            self.assertFalse(original.is_alive())
            self.assertFalse(next_worker.is_alive())
            self.assertFalse(following_results[0].ok)
            self.assertEqual(modem.writes, [b"ATH\r\n"])
            self.assertTrue(serial_sender.serial_sms_recovery_error(modem))
        finally:
            coordinator.cancel_active("test cleanup")
            original.join(2)
            if next_worker.ident is not None:
                next_worker.join(2)


if __name__ == "__main__":
    unittest.main()
