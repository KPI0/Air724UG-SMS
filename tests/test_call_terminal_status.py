"""Terminal call status follows the current serial phase and call lifetime."""
from contextlib import contextmanager
from functools import partial
import queue
import time
import unittest
from unittest.mock import patch

from sms_core.call_session import IncomingCallSessionTracker
from sms_core.cloud_message_runtime import handle_device_call_state_runtime
from sms_core.serial_app_runtime import (
    run_serial_app_from_wiring, run_serial_app_runtime, run_serial_reader_namespace_runtime,
)
from sms_core.serial_runtime import run_serial_runtime_thread
from sms_ui import call_popup_namespace_runtime as popup_runtime
from sms_ui.app_ui_namespace_runtime import set_cloud_status_namespace_runtime
from tests.test_connected_status_generation import install_status, drain
from tests import test_connected_status_generation as connected_fixture
from tests import test_serial_initialization as initialization_fixture
from tests import test_call_popup_command_isolation as popup_command_fixture
from tests.test_peer_terminal_ui_queue import PeerTerminalReplay


class CallTerminalStatusTests(unittest.TestCase):
    def wait_for(self, predicate):
        deadline = time.monotonic() + 3
        while time.monotonic() < deadline:
            if predicate():
                return
            time.sleep(0.005)
        self.fail("Virtual production reader did not reach expected state")

    @contextmanager
    def reader(self, *, event_hook=None, port_hook=None):
        harness = initialization_fixture.InitializationHarness(hold="AT+CNUM")
        namespace = harness.ns
        tasks = queue.Queue()
        value = install_status(namespace, tasks)
        tracker = IncomingCallSessionTracker()
        missed = []
        namespace.update(
            INCOMING_CALL_SESSION=tracker,
            start_serial_incoming_call=partial(popup_runtime.start_serial_incoming_call_namespace_runtime, namespace),
            finish_incoming_call_session=partial(popup_runtime.finish_incoming_call_session_namespace_runtime, namespace),
            reset_incoming_call_session=partial(popup_runtime.reset_incoming_call_session_namespace_runtime, namespace),
            show_missed_call_popup=missed.append,
        )
        if event_hook is not None:
            namespace["_cloud_send_call_event"] = lambda *args, **kwargs: event_hook(namespace, *args, **kwargs)
        if port_hook is not None:
            original_port_ui = namespace["port_ui"]
            def port_ui(text, *args):
                original_port_ui(text, *args)
                port_hook(namespace, text)
            namespace["port_ui"] = port_ui
        clocked_reader = partial(
            run_serial_reader_namespace_runtime,
            run_reader=partial(run_serial_app_from_wiring, run_app=partial(
                run_serial_app_runtime, run_runtime_thread=partial(
                    run_serial_runtime_thread, clock=lambda: harness.clock))),
        )
        with patch.object(initialization_fixture, "run_serial_reader_namespace_runtime", clocked_reader), harness:
            self.wait_for(lambda: harness.wrote["AT+CNUM"].is_set())
            namespace["restore_connected_status"]("TEST")
            drain(tasks)
            yield harness, namespace, tracker, tasks, value, missed
            self.assertEqual(harness.errors, [])

    def incoming(self, harness, tracker):
        harness.emit('RING\r\n[I]-[ril.proatc] +CLIP: "10086",129\r\n')
        self.wait_for(lambda: tracker.snapshot().session_token is not None)

    def finish(self, harness, namespace, tracker, tasks):
        harness.emit("[I]-[ril.proatc] NO CARRIER\r\n")
        self.wait_for(lambda: tracker.snapshot().session_token is None and namespace["ring_timeout_target"] == 0)
        self.wait_for(lambda: not tasks.empty())

    def peer(self, namespace, phase, local_id, peer_id="synthetic-peer"):
        payload = {
            "type": "device_call_state", "imei": "000000000000001",
            "direction": "incoming", "phase": phase, "phone": "10086",
            "call_session_id": peer_id,
        }
        if phase == "incoming":
            payload.update(local_call_session_id=local_id,
                           call_session_ids={"desktop": local_id, "firmware": peer_id})
        return handle_device_call_state_runtime(
            payload, runtime_imei=lambda: "000000000000001",
            register_incoming=partial(popup_runtime.register_peer_incoming_call_namespace_runtime, namespace),
            finish_call=partial(popup_runtime.finish_remote_incoming_call_namespace_runtime, namespace),
        )

    def test_real_reader_terminal_restores_connected_after_startup_ok_or_error(self):
        for response in ("OK", "ERROR"):
            with self.subTest(response=response), self.reader() as (harness, namespace, tracker, tasks, value, missed):
                harness.terminal("AT+CNUM", response=response)
                harness.emit("[I]-[probe] STARTUP_END\r\n")
                self.wait_for(lambda: "[I]-[probe] STARTUP_END" in harness.lines)
                self.assertFalse(getattr(harness.modem, "_sms_recovery_error", ""))
                drain(tasks)
                expected = value.text
                self.incoming(harness, tracker)
                drain(tasks)
                self.assertIn("响铃", value.text)
                self.finish(harness, namespace, tracker, tasks)
                drain(tasks)
                self.assertEqual(value.text, expected)
                self.assertEqual(len(missed), 1)
                self.assertEqual(harness.messages[-3:], [
                    "📞 收到来电：来自 10086", "📞 语音通话已结束", "📵 未接来电：10086",
                ])
                cloud_value = type(value)()
                namespace.update(cloud_var=cloud_value, cloud_label=cloud_value)
                for text in ("🌐 重连中", "🌐 连接中", "🌐 已授权") * 3:
                    set_cloud_status_namespace_runtime(namespace, text)
                self.assertEqual(value.text, expected)

    def test_real_reader_terminal_without_startup_ack_keeps_actual_transaction_phase(self):
        for quarantined in (False, True):
            with self.subTest(quarantined=quarantined), self.reader() as (harness, namespace, tracker, tasks, value, missed):
                if quarantined:
                    harness.modem._sms_recovery_error = "synthetic unresolved AT command"
                    namespace["restore_connected_status"]("TEST")
                drain(tasks)
                expected = value.text
                self.incoming(harness, tracker)
                drain(tasks)
                self.finish(harness, namespace, tracker, tasks)
                drain(tasks)
                self.assertEqual(value.text, expected)
                if not quarantined:
                    self.assertEqual(value.text, "🟢 已连接：TEST")
                else:
                    self.assertIn("串口指令结果未确认", value.text)
                    self.assertEqual(value.color, "red")
                self.assertEqual(len(missed), 1)

    def test_peer_terminal_during_incoming_event_does_not_restore_old_ringing_status(self):
        ended = []
        def terminal(namespace, _number, _message, **kwargs):
            local_id = kwargs["call_session_id"]
            self.peer(namespace, "incoming", local_id)
            self.peer(namespace, "ended", local_id)
            ended.append(True)
        with self.reader(event_hook=terminal) as (harness, namespace, tracker, tasks, value, missed):
            harness.terminal("AT+CNUM")
            self.assertTrue(harness.ready.is_set())
            drain(tasks)
            harness.emit('RING\r\n[I]-[ril.proatc] +CLIP: "10086",129\r\n')
            self.wait_for(lambda: bool(ended) and namespace["ring_timeout_target"] == 0)
            drain(tasks)
            self.assertIsNone(tracker.snapshot().session_token)
            self.assertEqual(value.text, "🟢 已连接：TEST")
            self.assertEqual(len(missed), 1)

    def test_terminal_cleanup_and_old_peer_terminal_preserve_same_number_redial(self):
        with self.reader() as (harness, namespace, tracker, tasks, value, missed):
            harness.terminal("AT+CNUM")
            self.assertTrue(harness.ready.is_set())
            self.incoming(harness, tracker)
            drain(tasks)
            original_token = tracker.snapshot().session_token
            local_id = namespace["call_popup_local_session"][1]
            self.peer(namespace, "incoming", local_id)
            self.finish(harness, namespace, tracker, tasks)
            # Old terminal cleanup is still waiting in the UI queue.
            self.incoming(harness, tracker)
            self.assertIsNot(tracker.snapshot().session_token, original_token)
            self.peer(namespace, "ended", local_id)
            drain(tasks)
            self.assertEqual(tracker.snapshot().caller_num, "10086")
            self.assertIn("响铃", value.text)
            self.assertGreater(namespace["ring_timeout_target"], harness.clock)
            self.assertEqual(len(missed), 1)

    def test_peer_terminal_between_connected_log_and_status_does_not_restore_old_connected_status(self):
        ended = []
        def port_hook(namespace, text):
            if "对方已接听" in text:
                local_id = namespace["call_popup_local_session"][1]
                self.peer(namespace, "ended", local_id)
                ended.append(True)
        with self.reader(port_hook=port_hook) as (harness, namespace, tracker, tasks, value, _missed):
            harness.terminal("AT+CNUM")
            self.assertTrue(harness.ready.is_set())
            drain(tasks)
            self.incoming(harness, tracker)
            drain(tasks)
            local_id = namespace["call_popup_local_session"][1]
            self.peer(namespace, "incoming", local_id)
            harness.emit("[I]-[ril.proatc] CONNECT\r\n")
            self.wait_for(lambda: bool(ended) and namespace["ring_timeout_target"] == 0)
            drain(tasks)
            self.assertIsNone(tracker.snapshot().session_token)
            self.assertEqual(value.text, "🟢 已连接：TEST")

    def test_queued_original_ringing_status_cannot_return_after_terminal_or_redial(self):
        for redial in (False, True):
            with self.subTest(redial=redial), self.reader() as (harness, namespace, tracker, tasks, value, _missed):
                harness.terminal("AT+CNUM")
                self.assertTrue(harness.ready.is_set())
                drain(tasks)
                self.incoming(harness, tracker)
                self.wait_for(lambda: not tasks.empty())
                callback, args, kwargs = tasks.get_nowait()
                tasks.task_done()
                self.finish(harness, namespace, tracker, tasks)
                if redial:
                    self.incoming(harness, tracker)
                drain(tasks)
                expected = value.text
                callback(*args, **kwargs)
                self.assertEqual(value.text, expected)
                self.assertEqual("响铃" in value.text, redial)

    def test_peer_connected_status_boundary_preserves_terminal_redial_and_changed_connection(self):
        for transition in ("ended", "redial", "reconnected"):
            with self.subTest(transition=transition):
                fixture = connected_fixture.ConnectedStatusGenerationTests("runTest")
                _owner, replay, namespace, value, serial = fixture.peer_fixture()
                local_id = replay.start_old()
                replay.drain()
                replay.cloud("incoming", local_id)
                original_status = namespace["set_status"]
                boundary = []

                def status(text, color="black", **kwargs):
                    if "通话中" in text and not boundary:
                        replay.cloud("ended")
                        boundary.append(True)
                        if transition == "redial":
                            replay.redial()
                        elif transition == "reconnected":
                            fixture.change_connection(namespace, serial, "reconnected")
                            original_status("new serial state", "orange")
                    return original_status(text, color, **kwargs)

                namespace["set_status"] = status
                replay.cloud("connected")
                replay.drain()
                self.assertEqual(boundary, [True])
                self.assertEqual(replay.errors, [])
                if transition == "redial":
                    self.assertEqual(namespace["INCOMING_CALL_SESSION"].snapshot().caller_num, "10086")
                    self.assertIn("响铃", value.text)
                    self.assertGreater(namespace["ring_timeout_target"], 101)
                else:
                    self.assertIsNone(namespace["INCOMING_CALL_SESSION"].snapshot().session_token)
                    expected = "new serial state" if transition == "reconnected" else "🟢 已连接：TEST"
                    self.assertEqual(value.text, expected)

    def test_confirmed_local_answer_status_after_peer_terminal_stays_finished(self):
        fixture = popup_command_fixture.CallPopupCommandIsolationTests("runTest")
        fixture.setUp()
        try:
            namespace = fixture.namespace
            replay = PeerTerminalReplay(namespace)
            value = install_status(namespace, replay.tasks)
            boundary = []

            def port_ui(text, *args):
                fixture.messages.append((text, *args))
                if "已发送接听指令" in text:
                    replay.cloud("ended")
                    boundary.append(True)

            namespace["port_ui"] = port_ui
            popup = fixture.start_popup()
            replay.cloud("incoming", "local-old")
            fixture.click(popup, "answer")
            for worker in fixture.threads:
                worker.join(3)
                self.assertFalse(worker.is_alive())
            drain(replay.tasks)
            self.assertEqual(boundary, [True])
            self.assertEqual(fixture.serial.writes, [b"ATA\r\n"])
            self.assertTrue(fixture.results[0].ok)
            self.assertIsNone(fixture.tracker.snapshot().session_token)
            self.assertEqual(namespace["ring_timeout_target"], 0)
            self.assertEqual(value.text, "🟢 已连接：TEST")
            self.assertEqual(fixture.errors, [])
        finally:
            fixture.tearDown()

    def test_queued_peer_connected_keeps_original_serial_before_ui_dispatch(self):
        for change in ("closed", "reconnected", "same_object_generation"):
            with self.subTest(change=change):
                fixture = connected_fixture.ConnectedStatusGenerationTests("runTest")
                _owner, replay, namespace, value, serial = fixture.peer_fixture()
                local_id = replay.start_old()
                replay.drain()
                replay.cloud("incoming", local_id)
                replay.cloud("connected")
                fixture.change_connection(namespace, serial, change)
                namespace["set_status"]("new serial state", "orange")
                replay.drain()
                self.assertEqual((value.text, value.color), ("new serial state", "orange"))
                self.assertEqual(replay.errors, [])


if __name__ == "__main__":
    unittest.main()
