import threading
import unittest

from sms_core.call_command_context import SerialCallStateSnapshot, set_local_dial_number
from sms_core.call_session import IncomingCallSessionTracker
from sms_core.call_send_namespace_runtime import send_local_serial_command_namespace_runtime
from sms_core.serial_runtime import SerialLineDecoder, SerialReadChunk, SerialReceivedLine
from sms_core.serial_sender import SmsPduSendCoordinator, write_serial_command_confirmed_locked
from tests import test_serial_dial_generation


class SerialCallResponseGenerationTests(unittest.TestCase):
    def test_partial_old_reply_does_not_complete_new_local_dial(self):
        for prefix, suffix, actual in (
            (b"NO ", b"CARRIER", "OK"),
            (b"NO ", b"ANSWER", "OK"),
            (b"BU", b"SY", "OK"),
            (b"ER", b"ROR", "OK"),
            (b"+CME ", b"ERROR: 3", "OK"),
            (b"O", b"K", "ERROR"),
        ):
            with self.subTest(prefix=prefix, suffix=suffix, actual=actual):
                self.exercise_partial_dial_reply(prefix, suffix, actual)

    def exercise_partial_dial_reply(self, prefix, suffix, actual):
        fixture = test_serial_dial_generation.SerialDialGenerationTests()
        fixture.setUp()
        namespace = fixture.namespace
        wrote = threading.Event()
        namespace["serial_obj"].response = None
        namespace["serial_obj"].after_write = wrote.set

        def queue_command(command, **kwargs):
            def start(_name, worker, **_kwargs):
                fixture.workers.append(worker)
            return send_local_serial_command_namespace_runtime(
                namespace, command, **kwargs, start_worker=start, response_timeout=2,
            )
        fixture.send_local = queue_command

        def scenario(handle, dial):
            decoder = SerialLineDecoder()
            old = namespace["get_serial_call_state"]()
            self.assertEqual(decoder.feed(SerialReadChunk(prefix, old)), [])
            dial("10086")
            worker = threading.Thread(target=fixture.workers.pop())
            worker.start()
            try:
                self.assertTrue(wrote.wait(1))
                current = namespace["get_serial_call_state"]()
                self.assertFalse(old.is_current())
                lines = decoder.feed(SerialReadChunk(suffix + b"\r\n", current))
                self.assertIs(lines[0].call_state_snapshot, old)
                handle(lines[0])
                handle(SerialReceivedLine(actual, current))
                worker.join(1)
                self.assertFalse(worker.is_alive())
                fixture.drain_ui()
                self.assertEqual(namespace["serial_obj"].writes, [b"ATD10086;\r\n"])
                self.assertEqual(namespace["current_dial_num"], "10086" if actual == "OK" else "")
                self.assertEqual(fixture.popups[-1].markers, [] if actual == "OK" else ["ended"])
                self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].qsize(), 1 if actual == "OK" else 0)
            finally:
                namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"].cancel_active("test cleanup")
                worker.join(3)
                self.assertFalse(worker.is_alive())

        fixture.run_scenario(scenario)

    def test_partial_old_reply_does_not_complete_current_answer_or_hangup(self):
        for command in ("ATA", "ATH"):
            for incoming in (False, True):
                for prefix, suffix, actual in ((b"ER", b"ROR", "OK"), (b"O", b"K", "ERROR")):
                    with self.subTest(command=command, incoming=incoming, actual=actual):
                        self.exercise_partial_control_reply(command, incoming, prefix, suffix, actual)

    def exercise_partial_control_reply(self, command, incoming, prefix, suffix, actual):
        fixture = test_serial_dial_generation.SerialDialGenerationTests()
        fixture.setUp()
        namespace = fixture.namespace
        wrote, results = threading.Event(), []
        namespace["serial_obj"].response = None
        namespace["serial_obj"].after_write = wrote.set
        tracker = IncomingCallSessionTracker()
        namespace["INCOMING_CALL_SESSION"] = tracker

        def scenario(handle, _dial):
            decoder = SerialLineDecoder()
            if incoming:
                tracker.start("10086")
            old = SerialCallStateSnapshot(namespace)
            self.assertEqual(decoder.feed(SerialReadChunk(prefix, old)), [])
            if incoming:
                tracker.finish()
                tracker.start("10086")
            else:
                set_local_dial_number(namespace, "10086")
            current = SerialCallStateSnapshot(namespace)

            def send():
                results.append(write_serial_command_confirmed_locked(
                    namespace["serial_lock"], lambda: namespace["serial_obj"], command,
                    response_coordinator=namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"],
                    response_timeout=2,
                ))
            worker = threading.Thread(target=send)
            worker.start()
            try:
                self.assertTrue(wrote.wait(1))
                old_line = decoder.feed(SerialReadChunk(suffix + b"\r\n", current))[0]
                handle(old_line)
                handle(SerialReceivedLine(actual, current))
                worker.join(1)
                self.assertFalse(worker.is_alive())
                self.assertEqual(len(results), 1)
                self.assertEqual(results[0].ok, actual == "OK")
                self.assertEqual(namespace["serial_obj"].writes, [(command + "\r\n").encode()])
            finally:
                namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"].cancel_active("test cleanup")
                worker.join(3)
                self.assertFalse(worker.is_alive())

        fixture.run_scenario(scenario)

    def test_call_generation_change_keeps_normal_at_sms_and_log_observation(self):
        fixture = test_serial_dial_generation.SerialDialGenerationTests()
        fixture.setUp()
        namespace = fixture.namespace
        namespace["SMS_SEND_COORDINATOR"] = SmsPduSendCoordinator()
        seen = []
        fixture.line_hook = lambda line, *_args: seen.append(str(line))

        def scenario(handle, _dial):
            old = SerialCallStateSnapshot(namespace)
            set_local_dial_number(namespace, "10086")
            self.assertFalse(old.is_current())
            coordinator = namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"]
            for command in ("ATI", "ATD*100#;", "ATD**21*10086#;"):
                waiter = coordinator.begin(command, connection=namespace["serial_obj"])
                try:
                    self.assertTrue(waiter.mark_write_started())
                    handle(SerialReceivedLine("OK", old))
                    self.assertTrue(waiter.wait(0).ok, command)
                finally:
                    coordinator.finish(waiter)
            sms = namespace["SMS_SEND_COORDINATOR"]
            segment = sms.begin_segment(connection=namespace["serial_obj"], cmgs_len=1)
            try:
                handle(SerialReceivedLine(">", old))
                self.assertTrue(segment.wait_prompt(0).ok)
                self.assertTrue(segment.mark_pdu_send_started())
                handle(SerialReceivedLine("+CMGS: 1", old))
                handle(SerialReceivedLine("OK", old))
                self.assertTrue(segment.wait(0).ok)
            finally:
                sms.finish(segment)
            self.assertTrue(all(item in seen for item in ("OK", ">", "+CMGS: 1")))

        fixture.run_scenario(scenario)


if __name__ == "__main__":
    unittest.main()
