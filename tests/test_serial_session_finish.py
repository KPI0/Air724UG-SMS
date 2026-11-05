from dataclasses import replace
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from sms_core.call_recordings import CallRecordingRepository, SerialCallRecordingReceiver
from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.long_sms_assembler import LongSmsAssembler
from sms_core.serial_app_runtime import build_serial_runtime_callbacks
from sms_core.serial_runtime import run_serial_runtime_thread
from sms_core.serial_sms import PendingSms
from sms_core.sms_pdu import ConcatSmsInfo
from tests.test_serial_runtime import runtime_callbacks, runtime_config
from tests.test_call_recordings import AMR_PAYLOAD, begin_frame, chunk_frame, end_frame
from tools.sms_replay import _incoming_ucs2_pdu, _wrap_hex


def complete_long_sms_frames(prefix="FIRST"):
    timestamps = ("26/07/22,19:52:00+32", "26/07/22,19:52:01+32")
    parts = (prefix + "-FIRST-SEGMENT-LONG-", prefix + "-SECOND-END")
    lines = []
    for index, (part, timestamp) in enumerate(zip(parts, timestamps), 1):
        pdu = _incoming_ucs2_pdu("10001", part, timestamp, reference=0x75, total=2, index=index)
        lines.extend([f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80",
                      *_wrap_hex(pdu), "[I]-[TP-PID : ] 0 dcs:  8"])
    body = "".join(parts)
    lines.extend([f"[I]-[handler_sms.smsCallback] 10001 {timestamps[0]} {body}", "[I]-[ril.proatc] OK"])
    return ("\r\n".join(lines) + "\r\n").encode(), body


class SerialSessionFinishTests(unittest.TestCase):
    def run_recording_session(self, receiver, frames, *, disconnect=False, callbacks=None):
        state = {"running": True, "reads": 0}
        closes = []
        callbacks = callbacks or build_serial_runtime_callbacks({**vars(runtime_callbacks([])),
            "consume_call_recording_line": receiver.consume_line,
            "abort_call_recording_transfer": receiver.abort})

        def read():
            state["reads"] += 1
            if state["reads"] == 1:
                return ("\r\n".join(frames) + "\r\n").encode()
            if disconnect:
                raise OSError("synthetic serial disconnected")
            state["running"] = False
            return b""

        def handle_disconnect(_error, _port):
            state["running"] = False
            receiver.abort("serial_disconnect")
            return False

        run_serial_runtime_thread(parse_callback_head=parse_sms_callback_head,
            get_runtime_config=runtime_config, callbacks=callbacks, get_call_state=lambda: (0.0, ""),
            set_call_state=lambda *_args: None, popup_active=lambda: False, ignore_repeat_state={},
            should_continue=lambda: state["running"], get_target_port=lambda: "TEST",
            resolve_target_port=lambda: "TEST", set_connecting_status=lambda *_args: None,
            open_and_initialize_serial=lambda *_args: None, on_connected_port=lambda *_args: None,
            read_serial_line=read, handle_disconnect=handle_disconnect, wait_before_retry=lambda: None,
            safe_close_serial=lambda: closes.append(True), settle_wait=lambda _seconds: None)
        return closes

    def test_normal_stop_releases_partial_recording_before_return(self):
        with tempfile.TemporaryDirectory() as temp:
            repository = CallRecordingRepository(temp)
            aborted = []
            receiver = SerialCallRecordingReceiver(repository, on_aborted=lambda _metadata, reason: aborted.append(reason))
            self.addCleanup(receiver.abort)
            closes = self.run_recording_session(receiver, [begin_frame("stopped-partial")])
            self.assertEqual(closes, [True])
            self.assertIsNone(receiver.current)
            self.assertEqual(list(Path(repository.incoming_dir).iterdir()), [])
            self.assertEqual(aborted, ["serial_stop"])

    def test_normal_stop_preserves_completed_recording(self):
        with tempfile.TemporaryDirectory() as temp:
            repository = CallRecordingRepository(temp)
            saved, aborted = [], []
            receiver = SerialCallRecordingReceiver(repository, on_saved=saved.append,
                on_aborted=lambda _metadata, reason: aborted.append(reason))
            self.addCleanup(receiver.abort)
            self.run_recording_session(receiver, [begin_frame("completed"),
                chunk_frame("completed", 1, AMR_PAYLOAD), end_frame("completed", 1)])
            self.assertEqual(aborted, [])
            self.assertEqual(len(saved), 1)
            self.assertEqual(Path(saved[0].path).read_bytes(), AMR_PAYLOAD)
            self.assertTrue(Path(saved[0].metadata_path).is_file())
            self.assertEqual(list(Path(repository.incoming_dir).iterdir()), [])

    def test_normal_stop_keeps_custom_callbacks_without_abort_compatible(self):
        with tempfile.TemporaryDirectory() as temp:
            receiver = SerialCallRecordingReceiver(CallRecordingRepository(temp))
            old_callbacks = vars(runtime_callbacks([])).copy()
            old_callbacks.pop("abort_call_recording_transfer")
            self.assertEqual(self.run_recording_session(receiver, [],
                callbacks=SimpleNamespace(**old_callbacks)), [True])

    def test_sms_finish_failure_still_releases_partial_recording(self):
        with tempfile.TemporaryDirectory() as temp:
            repository = CallRecordingRepository(temp)
            receiver = SerialCallRecordingReceiver(repository)
            self.addCleanup(receiver.abort)
            original = RuntimeError("synthetic SMS finish failure")
            with patch.object(LongSmsAssembler, "finish_session", side_effect=original):
                with self.assertRaises(RuntimeError) as raised:
                    self.run_recording_session(receiver, [begin_frame("sms-finish-failure")])
            self.assertIs(raised.exception, original)
            self.assertIsNone(receiver.current)
            self.assertEqual(list(Path(repository.incoming_dir).iterdir()), [])

    def test_abort_callback_failure_retains_sms_finish_exception_context(self):
        with tempfile.TemporaryDirectory() as temp:
            repository = CallRecordingRepository(temp)
            receiver = SerialCallRecordingReceiver(repository)
            self.addCleanup(receiver.abort)
            original = RuntimeError("synthetic SMS finish failure")
            cleanup_error = ValueError("synthetic recording cleanup failure")

            def abort(reason):
                receiver.abort(reason)
                raise cleanup_error

            callbacks = build_serial_runtime_callbacks({**vars(runtime_callbacks([])),
                "consume_call_recording_line": receiver.consume_line, "abort_call_recording_transfer": abort})
            with patch.object(LongSmsAssembler, "finish_session", side_effect=original):
                with self.assertRaises(ValueError) as raised:
                    self.run_recording_session(receiver, [begin_frame("dual-failure")], callbacks=callbacks)
            self.assertIs(raised.exception, cleanup_error)
            self.assertIs(raised.exception.__context__, original)
            self.assertIsNone(receiver.current)
            self.assertEqual(list(Path(repository.incoming_dir).iterdir()), [])

    def test_disconnect_then_stop_does_not_report_duplicate_recording_failure(self):
        with tempfile.TemporaryDirectory() as temp:
            repository = CallRecordingRepository(temp)
            aborted = []
            receiver = SerialCallRecordingReceiver(repository,
                on_aborted=lambda _metadata, reason: aborted.append(reason))
            self.addCleanup(receiver.abort)
            self.run_recording_session(receiver, [begin_frame("disconnect-partial")], disconnect=True)
            self.assertEqual(aborted, ["serial_disconnect"])
            self.assertIsNone(receiver.current)
            self.assertEqual(list(Path(repository.incoming_dir).iterdir()), [])

    def run_session(self, raw, mode, next_raw=None):
        state = {"running": True, "now": 10.0, "reads": 0, "device": 0}
        calls, deliveries, waits = [], [], []

        def opened(_port):
            state["device"] += 1

        def read():
            state["reads"] += 1
            if state["reads"] == 1:
                return raw
            if state["reads"] == 2:
                state["now"] = 10.1
                if mode == "disconnect":
                    raise OSError("synthetic disconnect")
                if mode == "shutdown":
                    state["running"] = False
                    return b""
                state["now"] = 20.0
                return b""
            if state["reads"] == 3 and next_raw is not None:
                return next_raw
            if state["reads"] == 4:
                state["now"] += 10.0
                return b""
            state["running"] = False
            return b""

        def wait(seconds):
            waits.append(seconds)
            state["now"] += seconds

        callbacks = replace(runtime_callbacks(calls), enqueue_third_push=lambda body, **_kwargs:
                            deliveries.append((state["device"], body, state["now"])))
        final = run_serial_runtime_thread(
            parse_callback_head=parse_sms_callback_head, get_runtime_config=runtime_config,
            callbacks=callbacks, get_call_state=lambda: (0.0, ""), set_call_state=lambda *a: None,
            popup_active=lambda: False, ignore_repeat_state={}, should_continue=lambda: state["running"],
            get_target_port=lambda: "TEST", resolve_target_port=lambda: "TEST", set_connecting_status=lambda *a: None,
            open_and_initialize_serial=opened, on_connected_port=lambda *a: None, read_serial_line=read,
            handle_disconnect=lambda *a: False, wait_before_retry=lambda: None, safe_close_serial=lambda: None,
            clock=lambda: state["now"], settle_wait=wait,
        )
        self.assertFalse(final.sms_collector.active)
        self.assertEqual(final.long_sms_assembler._deferred, {})
        return calls, deliveries, waits

    def test_complete_short_callback_survives_disconnect_and_shutdown(self):
        body = "SYNTHETIC-SHORT-SMS"
        raw = f"[I]-[handler_sms.smsCallback] 10001 26/07/22,19:52:00+32 {body}\r\n".encode()
        for mode in ("normal", "disconnect", "shutdown"):
            with self.subTest(mode=mode):
                calls, deliveries, waits = self.run_session(raw, mode)
                self.assertEqual([(device, text) for device, text, _time in deliveries], [(1, body)])
                self.assertEqual([item[1][1] for item in calls if item[0] == "cloud_sms"], [body])
                self.assertEqual([item[1][0] for item in calls if item[0] == "sms_popup"], [body])
                self.assertEqual(waits, [])

    def test_complete_deferred_long_sms_survives_disconnect_and_shutdown_after_grace(self):
        raw, body = complete_long_sms_frames()
        for mode in ("normal", "disconnect", "shutdown"):
            with self.subTest(mode=mode):
                calls, deliveries, waits = self.run_session(raw, mode)
                self.assertEqual([(device, text) for device, text, _time in deliveries], [(1, body)])
                self.assertGreaterEqual(deliveries[0][2], 15.0)
                self.assertEqual([item[1][1] for item in calls if item[0] == "cloud_sms"], [body])
                if mode != "normal":
                    self.assertAlmostEqual(sum(waits), 4.9)

    def test_old_session_finishes_before_new_device_with_same_reference(self):
        old_raw, old_body = complete_long_sms_frames("OLD")
        new_raw, new_body = complete_long_sms_frames("NEW")
        _calls, deliveries, _waits = self.run_session(old_raw, "disconnect", new_raw)
        self.assertEqual([(device, text) for device, text, _time in deliveries], [(1, old_body), (2, new_body)])

    def test_missing_parts_are_not_fabricated_as_a_complete_message(self):
        assembler = LongSmsAssembler(parse_sms_callback_head)
        partial = PendingSms("10001 26/07/22,19:52:00+32 A", "A", [], ConcatSmsInfo(7, 3, 1))
        self.assertIsNone(assembler.add_message(partial, now=1.0))
        waits = []
        self.assertIsNone(assembler.finish_session(now=2.0, wait=waits.append))
        self.assertEqual(waits, [])


if __name__ == "__main__":
    unittest.main()
