import threading
import time
from pathlib import Path
import tempfile
import unittest
from dataclasses import replace

from sms_core.call_session import IncomingCallSessionTracker
from sms_core.call_recordings import CallRecordingRepository, SerialCallRecordingReceiver
from sms_core.serial_runtime import SerialRuntimeState, handle_serial_runtime_line
from sms_core.serial_sender import SmsPduSendCoordinator
from sms_ui.call_popup_namespace_runtime import (
    finish_incoming_call_session_namespace_runtime,
    reset_incoming_call_session_namespace_runtime,
    start_serial_incoming_call_namespace_runtime,
)
from tests.test_serial_initialization import InitializationHarness
from tests.test_serial_runtime import parse_head, runtime_callbacks, runtime_config
from tests import test_serial_session_finish as session_finish_fixture
from tests.test_call_recordings import AMR_PAYLOAD, begin_frame, chunk_frame, end_frame


class SerialSmsControlIsolationTests(unittest.TestCase):
    def wait_for(self, predicate):
        deadline = time.monotonic() + 2
        while time.monotonic() < deadline:
            if predicate():
                return
            time.sleep(0.005)
        self.fail("Virtual reader did not process the expected input")

    def test_callback_body_cannot_complete_startup_waiter_but_real_terminal_can(self):
        for terminal in ("OK", "ERROR"):
            with self.subTest(terminal=terminal), InitializationHarness(hold="AT+CNUM") as harness:
                self.wait_for(lambda: harness.wrote["AT+CNUM"].is_set())
                session = harness.opened_session
                waiter = session.waiter
                observed = threading.Event()
                harness.on_line = lambda line: observed.set() if line == terminal else None
                harness.emit(
                    "[I]-[handler_sms.smsCallback] 10000 26/11/04,10:00:00+32 first line\r\n"
                    + terminal + "\r\n"
                )
                self.assertTrue(observed.wait(2))
                self.assertFalse(waiter.done(), "SMS body was accepted as a Modem terminal")
                self.assertTrue(session.pending)
                self.assertTrue(session.lock_held)
                self.assertEqual(session.index, 2)
                harness.terminal("AT+CNUM", response=terminal)
                self.wait_for(lambda: not session.pending)
                self.assertEqual(session.index, 3)
                self.assertFalse(session.lock_held)
                self.assertEqual(session.has_command_errors, terminal == "ERROR")
                self.assertEqual(harness.errors, [])

    def incoming_harness(self):
        harness = InitializationHarness(hold="NONE")
        namespace = harness.ns
        tracker = IncomingCallSessionTracker()
        terminals, missed, connected = [], [], []
        namespace.update(
            INCOMING_CALL_SESSION=tracker,
            start_serial_incoming_call=lambda *args: start_serial_incoming_call_namespace_runtime(namespace, *args),
            finish_incoming_call_session=lambda: finish_incoming_call_session_namespace_runtime(namespace),
            reset_incoming_call_session=lambda: reset_incoming_call_session_namespace_runtime(namespace),
            show_missed_call_popup=missed.append,
            mark_call_connected=lambda *args: connected.append(args),
            _cloud_send_call_state=lambda *args, **kwargs: terminals.append((args, kwargs)),
        )
        return harness, tracker, terminals, missed, connected

    def test_callback_prompt_and_error_do_not_change_sms_waiter_but_real_responses_do(self):
        harness = InitializationHarness(hold="NONE")
        coordinator = SmsPduSendCoordinator()
        harness.ns["SMS_SEND_COORDINATOR"] = coordinator
        with harness:
            self.wait_for(lambda: not harness.opened_session.pending)
            waiter = coordinator.begin_segment(connection=harness.modem, cmgs_len=1)
            try:
                harness.emit(
                    "[I]-[handler_sms.smsCallback] 10000 26/11/04,10:00:00+32 first line\r\n"
                    ">\r\nERROR\r\n[I]-[probe] BODY_END\r\n"
                )
                self.wait_for(lambda: "[I]-[probe] BODY_END" in harness.lines)
                self.assertFalse(waiter.wait_prompt(0).ok)
                self.assertFalse(waiter.done())
                harness.emit("[I]-[ril.proatc] >\r\n[I]-[probe] PROMPT_END\r\n")
                self.wait_for(lambda: "[I]-[probe] PROMPT_END" in harness.lines)
                self.assertTrue(waiter.wait_prompt(0).ok)
                self.assertTrue(waiter.mark_pdu_send_started())
                harness.emit("[I]-[ril.proatc] +CMGS: 1\r\n[I]-[ril.proatc] OK\r\n")
                self.wait_for(waiter.done)
                self.assertTrue(waiter.wait(0).ok)
            finally:
                coordinator.finish(waiter)
            self.assertEqual(harness.errors, [])

    def test_callback_body_does_not_change_live_call_but_prefixed_terminal_does(self):
        harness, tracker, terminals, missed, connected = self.incoming_harness()
        with harness:
            self.wait_for(lambda: not harness.opened_session.pending)
            harness.emit('[I]-[ril.proatc] +CLIP: "10000",129\r\n')
            self.wait_for(lambda: bool(tracker.snapshot().caller_num))
            token = tracker.snapshot().session_token
            deadline = harness.ns["ring_timeout_target"]
            harness.clock = 101.0
            harness.emit(
                "[I]-[handler_sms.smsCallback] 10000 26/11/04,10:00:00+32 first line\r\n"
                "NO CARRIER\r\nCONNECT\r\nRING\r\n"
                "[I]-[probe] BODY_END\r\n"
            )
            self.wait_for(lambda: "[I]-[probe] BODY_END" in harness.lines)
            self.assertIs(tracker.snapshot().session_token, token)
            self.assertEqual(harness.ns["ring_timeout_target"], deadline)
            self.assertEqual((terminals, missed, connected), ([], [], []))
            harness.emit("[I]-[ril.proatc] NO CARRIER\r\n[I]-[probe] TERMINAL_END\r\n")
            self.wait_for(lambda: "[I]-[probe] TERMINAL_END" in harness.lines)
            self.assertIsNone(tracker.snapshot().session_token)
            self.assertEqual([args[1] for args, _kwargs in terminals], ["ended"])
            self.assertEqual(len(missed), 1)
            self.assertEqual(harness.errors, [])

    def test_owned_callback_keeps_entire_body_and_logs_without_control_side_effects(self):
        state = SerialRuntimeState.create(parse_head)
        calls = []
        controls, recordings, responses = [], [], []
        callbacks = replace(
            runtime_callbacks(calls),
            observe_call_command_line=controls.append,
            observe_sms_send_line=responses.append,
            consume_call_recording_line=lambda line: recordings.append(line) or False,
        )
        body_lines = ["first line", "OK", "ERROR", "NO CARRIER", "CONNECT", "RING",
                      "000000000000001", "recording notes"]
        head = "10000 " + body_lines[0]
        raw_lines = ["[I]-[handler_sms.smsCallback] " + head] + body_lines[1:]
        for line in raw_lines:
            handle_serial_runtime_line(state, line, 10.0, "TEST", False,
                                       runtime_config(), callbacks, {})
        self.assertEqual((controls, recordings, responses), ([], [], []))
        self.assertEqual(state.call_state.ring_timeout_target, 0.0)
        self.assertEqual(state.sms_collector.raw_lines, body_lines[1:])
        self.assertEqual([args[0] for name, args, *rest in calls if name == "debug"], raw_lines)
        self.assertEqual([args[0] for name, args, *rest in calls if name == "cloud_log"], raw_lines)
        self.assertFalse(any(name in {"imei", "temp", "signal", "call_popup", "hangup"}
                             for name, *rest in calls))
        handle_serial_runtime_line(state, "", 12.0, "TEST", False,
                                   runtime_config(), callbacks, {})
        handle_serial_runtime_line(state, "", 20.0, "TEST", False,
                                   runtime_config(), callbacks, {})
        self.assertIn(("sms_popup", ("\n".join(body_lines),)), calls)
        self.assertIn(("cloud_sms", (head, "\n".join(body_lines))), calls)

    def test_long_callback_does_not_suspend_real_ring_timeout(self):
        state = SerialRuntimeState.create(parse_head)
        calls = []
        tracker = IncomingCallSessionTracker()
        tracker.start("10000")
        state.call_state.last_clip_num = "10000"
        state.call_state.call_session_id = "incoming:synthetic"
        state.call_state.ring_timeout_target = 11.0
        callbacks = replace(runtime_callbacks(calls), finish_incoming_call=tracker.finish)
        handle_serial_runtime_line(state, "[I]-[handler_sms.smsCallback] 10000 first line",
                                   10.0, "TEST", True, runtime_config(), callbacks, {})
        handle_serial_runtime_line(state, "RING", 12.0, "TEST", True,
                                   runtime_config(), callbacks, {})
        self.assertIsNone(tracker.snapshot().session_token)
        self.assertEqual(state.call_state.ring_timeout_target, 0.0)
        self.assertEqual(state.sms_collector.raw_lines, ["RING"])
        self.assertIn(("close_popup",), calls)

    def test_real_boundary_releases_callback_and_keeps_independent_recording(self):
        state = SerialRuntimeState.create(parse_head)
        calls, recordings, responses = [], [], []
        callbacks = replace(
            runtime_callbacks(calls),
            consume_call_recording_line=lambda line: recordings.append(line) or line.startswith("@@CALL_RECORD"),
            observe_sms_send_line=responses.append,
        )
        for line in ("[I]-[handler_sms.smsCallback] 10000 first line", "body line",
                     "[I]-[ril.proatc] OK", "@@CALL_RECORD_BEGIN|synthetic|frame"):
            handle_serial_runtime_line(state, line, 10.0, "TEST", False,
                                       runtime_config(), callbacks, {})
        self.assertEqual(recordings, ["[I]-[ril.proatc] OK", "@@CALL_RECORD_BEGIN|synthetic|frame"])
        self.assertEqual(responses, ["[I]-[ril.proatc] OK"])
        self.assertFalse(state.sms_collector.active)

    def test_reserved_recording_frames_complete_while_sms_callback_is_active(self):
        with tempfile.TemporaryDirectory() as temp:
            saved, aborted, controls, responses, calls = [], [], [], [], []
            receiver = SerialCallRecordingReceiver(
                CallRecordingRepository(temp), on_saved=saved.append,
                on_aborted=lambda _metadata, reason: aborted.append(reason),
            )
            self.addCleanup(receiver.abort)
            callbacks = replace(
                runtime_callbacks(calls),
                consume_call_recording_line=receiver.consume_line,
                abort_call_recording_transfer=receiver.abort,
                observe_call_command_line=controls.append,
                observe_sms_send_line=responses.append,
            )
            head = "10000 26/11/04,10:00:00+32 first line"
            frames = [
                "[I]-[handler_sms.smsCallback] " + head, "before recording",
                begin_frame("interleaved", phone="10000", imei="000000000000001"),
                chunk_frame("interleaved", 1, AMR_PAYLOAD), end_frame("interleaved", 1),
                "after recording", "[I]-[ril.proatc] OK",
            ]
            session_finish_fixture.SerialSessionFinishTests("runTest").run_recording_session(
                receiver, frames, callbacks=callbacks)
            self.assertEqual(aborted, [])
            self.assertEqual(len(saved), 1)
            self.assertEqual(Path(saved[0].path).read_bytes(), AMR_PAYLOAD)
            self.assertEqual(controls, ["[I]-[ril.proatc] OK", ""])
            self.assertEqual(responses, ["[I]-[ril.proatc] OK", ""])
            body = "first line\nbefore recording\nafter recording"
            self.assertIn(("sms_popup", (body,)), calls)
            self.assertIn(("cloud_sms", (head, body)), calls)


if __name__ == "__main__":
    unittest.main()
