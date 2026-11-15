"""Local dial launch retains the request made before logging and Tk work."""
from contextlib import ExitStack
from dataclasses import replace
from functools import partial
import unittest
import threading
from unittest.mock import patch

from sms_core.call_session import IncomingCallSessionTracker
from sms_core.call_send_namespace_runtime import send_local_serial_command_namespace_runtime
from tests.test_connected_status_generation import install_status
from tests import test_local_call_command_isolation as local_fixture
from tests.test_local_outgoing_calls import namespace_fixture
from tests.test_peer_terminal_ui_queue import PeerTerminalReplay
from tests import test_serial_dial_generation as dial_generation_fixture
from sms_ui import serial_debug_window
from sms_ui import serial_debug_namespace_runtime


class LocalDialLaunchStatusTests(unittest.TestCase):
    def fixture(self):
        fixture = local_fixture.LocalDialButtonIsolationTests("runTest")
        fixture.setUp()
        namespace = fixture.ns
        namespace["INCOMING_CALL_SESSION"] = IncomingCallSessionTracker()
        replay = PeerTerminalReplay(namespace, show_popup=lambda *_args, **_kwargs: None)
        value = install_status(namespace, replay.tasks)
        replay.callbacks = replace(replay.callbacks, set_status=namespace["set_status"])
        return fixture, namespace, replay, value

    def incoming(self, replay, number):
        replay.serial("RING", 100.0)
        replay.serial(f'+CLIP: "{number}",129', 100.1)
        replay.drain()

    def test_log_and_popup_construction_boundaries_cancel_old_dial_for_new_incoming(self):
        for boundary in ("log", "popup", "submit"):
            for caller in ("10086", "10010"):
                with self.subTest(boundary=boundary, caller=caller):
                    fixture, namespace, replay, value = self.fixture()
                    original_serial = namespace["serial_obj"]
                    reached = []

                    def transition():
                        self.incoming(replay, caller)
                        reached.append(True)

                    if boundary == "log":
                        old_port_ui = namespace["port_ui"]
                        def port_ui(text, *args):
                            old_port_ui(text, *args)
                            if "主动呼叫" in text:
                                transition()
                        namespace["port_ui"] = port_ui
                    with ExitStack() as stack:
                        stack.enter_context(patch("sms_ui.serial_debug_namespace_runtime.send_local_serial_command_namespace_runtime",
                                                  partial(send_local_serial_command_namespace_runtime, response_timeout=0.05)))
                        fixture.open_window(stack)
                        if boundary == "popup":
                            original_popup = serial_debug_window.open_dial_call_popup
                            def open_popup(*args):
                                popup = original_popup(*args)
                                transition()
                                return popup
                            stack.enter_context(patch.object(serial_debug_window, "open_dial_call_popup", open_popup))
                        elif boundary == "submit":
                            original_submit = serial_debug_namespace_runtime.send_local_serial_command_namespace_runtime
                            def submit(*args, **kwargs):
                                transition()
                                return original_submit(*args, **kwargs)
                            stack.enter_context(patch.object(serial_debug_namespace_runtime,
                                                            "send_local_serial_command_namespace_runtime", submit))
                        fixture.actions["dial"]("10086")
                        fixture.join_workers()
                        replay.drain()
                        self.assertEqual(reached, [True])
                        self.assertEqual(original_serial.writes, [])
                        self.assertEqual(namespace["INCOMING_CALL_SESSION"].snapshot().caller_num, caller)
                        self.assertEqual(value.text, f"🔔 响铃中：{caller}")
                        self.assertEqual(namespace["current_dial_num"], "")
                        self.assertIsNone(namespace["current_dial_popup"])
                        self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())
                        self.assertEqual(replay.errors, [])

    def test_log_boundary_preserves_changed_serial_and_explicit_pending_cancel(self):
        for transition in ("serial_object", "serial_generation", "cancel"):
            with self.subTest(transition=transition):
                fixture, namespace, replay, value = self.fixture()
                original_serial = namespace["serial_obj"]
                replacement = namespace_fixture()["serial_obj"]
                reached = []
                def port_ui(text, *args):
                    if "主动呼叫" in text:
                        reached.append(True)
                        if transition == "cancel":
                            send_local_serial_command_namespace_runtime(namespace, "ATH")
                        else:
                            namespace["serial_connection_generation"] += 1
                            if transition == "serial_object":
                                namespace["serial_obj"] = replacement
                        namespace["set_status"]("new serial state", "orange")
                namespace["port_ui"] = port_ui
                with ExitStack() as stack:
                    stack.enter_context(patch("sms_ui.serial_debug_namespace_runtime.send_local_serial_command_namespace_runtime",
                                              partial(send_local_serial_command_namespace_runtime, response_timeout=0.05)))
                    fixture.open_window(stack)
                    fixture.actions["dial"]("10086")
                    fixture.join_workers()
                    replay.drain()
                    self.assertEqual(reached, [True])
                    self.assertEqual(original_serial.writes, [])
                    self.assertEqual(replacement.writes, [])
                    self.assertEqual(value.text, "new serial state")
                    self.assertIsNone(namespace["current_dial_popup"])
                    self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())

    def test_current_launch_still_sends_once_and_keeps_popup_and_history(self):
        fixture, namespace, replay, value = self.fixture()
        with ExitStack() as stack:
            fixture.open_window(stack)
            fixture.actions["dial"]("10086")
            fixture.join_workers()
            replay.drain()
            self.assertEqual(namespace["serial_obj"].writes, [b"ATD10086;\r\n"])
            self.assertEqual(value.text, "📞 呼叫中：10086")
            self.assertIs(namespace["current_dial_popup"], fixture.popups[-1])
            self.assertTrue(fixture.popups[-1].exists)
            self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].qsize(), 1)

    def test_current_dial_failure_keeps_its_popup_for_immediate_and_async_end_marker(self):
        for deferred in (False, True):
            for response in ("ERROR", "+CME ERROR: 3", "write_failure"):
                with self.subTest(deferred=deferred, response=response):
                    fixture = dial_generation_fixture.SerialDialGenerationTests("runTest")
                    fixture.setUp()
                    fixture.pending = deferred

                    def scenario(_handle, dial):
                        serial = fixture.namespace["serial_obj"]
                        serial.response = response
                        if response == "write_failure":
                            def fail_write(_payload):
                                raise OSError("synthetic write failure")
                            serial.write = fail_write
                        dial("10086")
                        if deferred:
                            worker = threading.Thread(target=fixture.workers.pop())
                            worker.start()
                            worker.join(3)
                            self.assertFalse(worker.is_alive())
                        popup = fixture.popups[-1]
                        self.assertIs(fixture.namespace["current_dial_popup"], popup)
                        self.assertTrue(popup.exists)
                        self.assertEqual(popup.markers, [])
                        fixture.drain_ui()
                        self.assertEqual(popup.markers, ["ended"])
                        self.assertEqual(fixture.namespace["current_dial_num"], "")

                    fixture.run_scenario(scenario)


if __name__ == "__main__":
    unittest.main()
