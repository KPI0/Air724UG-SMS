import queue
import threading
import unittest
from contextlib import ExitStack
from types import SimpleNamespace
from unittest.mock import Mock, patch

from sms_core.call_send_namespace_runtime import send_local_serial_command_namespace_runtime
from sms_core.cloud_sms_event_runtime import CloudSmsEventDrainState
from sms_core.serial_sender import (
    AtCommandResponseCoordinator, DEFAULT_SERIAL_TRANSACTION_LOCK, send_command_async,
)
from sms_core.threading_runtime import WorkerThreadRegistry
from sms_ui.serial_debug_namespace_runtime import open_serial_debug_window_namespace_runtime
from tests.test_serial_debug_window import (
    FakeFinder, FakePauseController, FakeVar, FakeWidget, FakeWindow,
)


IMEI = "000000000000001"
OTHER_IMEI = "000000000000002"


class ModemSerial:
    def __init__(self, coordinator, response="OK"):
        self.coordinator = coordinator
        self.response = response
        self.is_open = True
        self.writes = []
        self.after_write = None

    def write(self, payload):
        self.writes.append(payload)
        if self.response:
            lines = [self.response] if isinstance(self.response, str) else self.response
            for line in lines:
                self.coordinator.observe_line(line, connection=self)
        if self.after_write:
            self.after_write()
        return len(payload)

    def flush(self):
        pass


def namespace_fixture(response="OK"):
    coordinator = AtCommandResponseCoordinator()
    namespace = {
        "root": "root", "serial_debug_win": None, "serial_debug_text": None,
        "SERIAL_DEBUG_ENABLED": True, "serial_debug_drop_count": 0,
        "serial_debug_queue": queue.Queue(), "current_dial_num": "", "PORT": "TEST",
        "serial_lock": threading.RLock(), "serial_connection_generation": 1,
        "serial_obj": ModemSerial(coordinator, response),
        "SERIAL_COMMAND_RESPONSE_COORDINATOR": coordinator,
        "SERIAL_COMMAND_THREAD_REGISTRY": WorkerThreadRegistry(),
        "CLOUD_CONTROL_ENABLED": True, "cloud_connected": False,
        "cloud_device_authorized": False, "cloud_ws_loop": None, "cloud_ws_conn": None,
        "CLOUD_SMS_EVENT_Q": queue.Queue(maxsize=50),
        "CLOUD_SMS_EVENT_DRAIN_STATE": CloudSmsEventDrainState(),
        "_cloud_now_ts": lambda: 1789171200,
        "_push_serial_debug": lambda *args: None, "port_ui": lambda *args: None,
        "set_status": lambda *args: None, "format_connected_status": lambda port: port,
        "center_window": lambda *args: None, "log_file_only": lambda *args: None,
        "_schedule_cloud_sms_event_drain": lambda *args: None, "imei": IMEI,
    }
    namespace["_cloud_runtime_imei"] = lambda: namespace["imei"]
    namespace["_cloud_identity_payload"] = lambda: {"imei": namespace["imei"], "device_imei": namespace["imei"]}
    return namespace


class LocalOutgoingCallRuntimeTests(unittest.TestCase):
    def run_dial(self, namespace, command="ATD10086;", append_crlf=True, before_worker=None):
        def start(_name, target, **_kwargs):
            if before_worker:
                before_worker()
            return target()

        return send_local_serial_command_namespace_runtime(
            namespace, command, append_crlf=append_crlf,
            start_worker=start, response_timeout=0.001,
        )

    def test_modem_acceptance_and_unanswered_terminals_each_create_one_attempt(self):
        for prefix in ("", "[I]-[ril.proatc] "):
            for terminal in ("OK", "NO ANSWER", "BUSY", "NO CARRIER"):
                with self.subTest(prefix=prefix, terminal=terminal):
                    namespace = namespace_fixture([prefix + terminal] * 2)
                    self.run_dial(namespace)
                    event = namespace["CLOUD_SMS_EVENT_Q"].get_nowait()
                    self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())
                    self.assertEqual(event["direction"], "outgoing")
                    self.assertEqual(event["phase"], "outgoing")
                    self.assertEqual(event["phone"], "10086")
                    self.assertEqual(event["timestamp"], namespace["_cloud_now_ts"]())
                    self.assertEqual(event["source_event_id"], event["call_session_id"])
                    self.assertTrue(event["ack_required"])
                    self.assertEqual(namespace["serial_obj"].writes, [b"ATD10086;\r\n"])

    def test_rejected_timed_out_and_untrusted_responses_do_not_create_history(self):
        for response in (None, "ERROR", "+CME ERROR: 3", "+CMS ERROR: 500",
                         "[I]-[app.note] OK", "[I]-[app.note] NO CARRIER",
                         ">>> 云端发送: ATD10086; OK", "CONNECT"):
            with self.subTest(response=response):
                namespace = namespace_fixture(response)
                self.run_dial(namespace)
                self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())
                self.assertEqual(namespace["serial_obj"].writes, [b"ATD10086;\r\n"])
                self.assertIsNone(namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active)

    def test_write_exception_text_is_not_modem_evidence_and_never_retries_dial(self):
        namespace = namespace_fixture()
        namespace["serial_obj"].write = Mock(side_effect=OSError("NO CARRIER"))
        self.run_dial(namespace)
        namespace["serial_obj"].write.assert_called_once_with(b"ATD10086;\r\n")
        self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())
        self.assertIsNone(namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active)

    def test_response_from_another_serial_cannot_confirm_local_dial(self):
        namespace = namespace_fixture(None)
        coordinator = namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"]
        namespace["serial_obj"].after_write = lambda: coordinator.observe_line("OK", connection=object())
        self.run_dial(namespace)
        self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())

    def test_queued_dial_is_cancelled_when_connection_generation_or_identity_changes(self):
        for field in ("serial_obj", "serial_connection_generation", "imei"):
            with self.subTest(field=field):
                namespace = namespace_fixture()
                original_serial = namespace["serial_obj"]
                replacement = ModemSerial(namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"])
                value = {"serial_obj": replacement, "serial_connection_generation": 2,
                         "imei": OTHER_IMEI}[field]
                self.run_dial(namespace, before_worker=lambda: namespace.update({field: value}))
                self.assertEqual(original_serial.writes, [])
                self.assertEqual(replacement.writes, [])
                self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())

    def test_context_is_rechecked_immediately_before_write_and_waiter_is_cleaned(self):
        namespace = namespace_fixture()
        coordinator = namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"]
        original_begin = coordinator.begin

        def change_after_begin(**kwargs):
            waiter = original_begin(**kwargs)
            namespace["serial_connection_generation"] += 1
            return waiter

        with patch.object(coordinator, "begin", side_effect=change_after_begin):
            self.run_dial(namespace)
        self.assertEqual(namespace["serial_obj"].writes, [])
        self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())
        self.assertIsNone(coordinator._active)

    def test_context_getter_exception_cleans_waiter_without_sending_or_recording(self):
        namespace = namespace_fixture()
        namespace["_cloud_runtime_imei"] = Mock(side_effect=[IMEI, RuntimeError("synthetic identity failure")])
        self.run_dial(namespace)
        self.assertEqual(namespace["serial_obj"].writes, [])
        self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())
        self.assertIsNone(namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active)

    def test_shutdown_and_stopped_serial_cancel_queued_dial(self):
        for field in ("serial_stop_event", "TK_SHUTDOWN", "serial_running"):
            with self.subTest(field=field):
                namespace = namespace_fixture()
                stopped = threading.Event()
                stopped.set()
                namespace[field] = False if field == "serial_running" else stopped
                self.run_dial(namespace)
                self.assertEqual(namespace["serial_obj"].writes, [])
                self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())

    def test_closed_serial_does_not_send_or_create_history(self):
        namespace = namespace_fixture()
        namespace["serial_obj"].is_open = False
        self.run_dial(namespace)
        self.assertEqual(namespace["serial_obj"].writes, [])
        self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())

    def test_switch_after_modem_confirmation_keeps_original_device_and_queues_offline(self):
        namespace = namespace_fixture()
        serial = namespace["serial_obj"]
        namespace.update(cloud_connected=True, cloud_device_authorized=True,
                         cloud_ws_conn=object(), cloud_ws_loop=SimpleNamespace(is_running=lambda: True))
        namespace["_schedule_cloud_sms_event_drain"] = Mock()
        serial.after_write = lambda: namespace.update(imei=OTHER_IMEI, serial_obj=object(),
                                                      serial_connection_generation=2)
        self.run_dial(namespace)
        event = namespace["CLOUD_SMS_EVENT_Q"].get_nowait()
        self.assertEqual(event["imei"], IMEI)
        self.assertEqual(event["device_imei"], IMEI)
        namespace["_schedule_cloud_sms_event_drain"].assert_not_called()
        self.assertEqual(serial.writes, [b"ATD10086;\r\n"])

    def test_online_authorized_device_schedules_existing_reliable_event_queue(self):
        namespace = namespace_fixture()
        namespace.update(cloud_connected=True, cloud_device_authorized=True,
                         cloud_ws_conn=object(), cloud_ws_loop=SimpleNamespace(is_running=lambda: True))
        namespace["_schedule_cloud_sms_event_drain"] = Mock()
        self.run_dial(namespace)
        namespace["_schedule_cloud_sms_event_drain"].assert_called_once_with(
            namespace["cloud_ws_loop"], namespace["cloud_ws_conn"],
        )
        self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].qsize(), 1)

    def test_cloud_disabled_or_missing_identity_does_not_prevent_local_dial(self):
        for overrides in ({"CLOUD_CONTROL_ENABLED": False}, {"imei": ""}):
            with self.subTest(overrides=overrides):
                namespace = namespace_fixture()
                namespace.update(overrides)
                self.run_dial(namespace)
                self.assertEqual(namespace["serial_obj"].writes, [b"ATD10086;\r\n"])
                self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())

    def test_same_number_redials_in_same_second_have_distinct_stable_history_ids(self):
        namespace = namespace_fixture()
        self.run_dial(namespace)
        self.run_dial(namespace)
        first, second = list(namespace["CLOUD_SMS_EVENT_Q"].queue)
        self.assertEqual(first["timestamp"], second["timestamp"])
        self.assertNotEqual(first["source_event_id"], second["source_event_id"])
        self.assertNotEqual(first["call_session_id"], second["call_session_id"])

    def test_upload_failure_does_not_repeat_the_successful_dial(self):
        namespace = namespace_fixture()
        namespace["port_ui"] = Mock()
        with patch("sms_core.call_send_namespace_runtime.enqueue_cloud_sms_event_runtime",
                   side_effect=RuntimeError("synthetic queue failure")):
            self.run_dial(namespace)
        self.assertEqual(namespace["serial_obj"].writes, [b"ATD10086;\r\n"])
        self.assertIn("请勿为补记录重复拨号", namespace["port_ui"].call_args.args[0])

    def test_voice_dial_preserves_editor_bytes_and_explicit_line_ending(self):
        for command, append, expected in (("  atd+10086;\t", True, b"  atd+10086;\t\r\n"),
                                          ("ATD10086;\r", False, b"ATD10086;\r"),
                                          ("ATD10086;\r\n", False, b"ATD10086;\r\n")):
            with self.subTest(command=command, append=append):
                namespace = namespace_fixture()
                self.run_dial(namespace, command, append_crlf=append)
                self.assertEqual(namespace["serial_obj"].writes, [expected])
                self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].qsize(), 1)

    def test_other_commands_and_unterminated_atd_keep_the_existing_raw_sender(self):
        for command, append in (("AT", True), ("ATD*100#;", True), ("ATD**21*10086#;", True),
                                ("ATD10086;", False), ("ATD10086;\r\nATH", True)):
            with self.subTest(command=command, append=append):
                namespace = namespace_fixture()
                raw = Mock(return_value="raw-worker")
                result = send_local_serial_command_namespace_runtime(
                    namespace, command, append_crlf=append, send_raw=raw,
                )
                self.assertEqual(result, "raw-worker")
                self.assertEqual(raw.call_args.args[2], command)
                self.assertEqual(raw.call_args.kwargs["append_crlf"], append)
                self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())

    def test_real_transaction_queue_uses_execution_time_and_unregisters_worker(self):
        namespace = namespace_fixture()
        with DEFAULT_SERIAL_TRANSACTION_LOCK:
            worker = send_local_serial_command_namespace_runtime(namespace, "ATD10086;")
            self.assertIn(worker, namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot())
            self.assertEqual(namespace["serial_obj"].writes, [])
            namespace["_cloud_now_ts"] = lambda: 1789171500
        worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].get_nowait()["timestamp"], 1789171500)
        self.assertEqual(namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(), ())

    def test_disconnect_cancels_response_wait_and_unregisters_worker(self):
        namespace = namespace_fixture(None)
        wrote = threading.Event()
        namespace["serial_obj"].after_write = wrote.set
        worker = send_local_serial_command_namespace_runtime(namespace, "ATD10086;")
        try:
            self.assertTrue(wrote.wait(2))
            self.assertIn(worker, namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot())
        finally:
            namespace["SERIAL_COMMAND_RESPONSE_COORDINATOR"].cancel_active("serial disconnected")
            worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())
        self.assertEqual(namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(), ())


class LocalOutgoingCallUiTests(unittest.TestCase):
    def test_dial_button_and_raw_atd_each_enqueue_one_confirmed_outgoing_event(self):
        namespace = namespace_fixture()
        captured, workers = {}, []

        def capture_actions(*args):
            captured.update(quick_send=args[5], dial=args[8])

        def original_sender(*args, **kwargs):
            worker = send_command_async(*args, **kwargs)
            workers.append(worker)
            return worker

        with ExitStack() as stack:
            overrides = {
                "tk.Toplevel": lambda *args: FakeWindow(),
                "tk.BooleanVar": FakeVar, "tk.StringVar": FakeVar,
                "ttk.Frame": FakeWidget, "ttk.Checkbutton": FakeWidget,
                "ttk.Label": FakeWidget, "ttk.Entry": FakeWidget, "ttk.Button": FakeWidget,
                "create_serial_debug_body": lambda *args: (FakeWidget(), FakeWidget(), FakeWidget()),
                "create_serial_debug_quick_actions": capture_actions,
                "SerialDebugPauseController": FakePauseController, "SerialDebugFinder": FakeFinder,
                "start_serial_debug_append_loop": lambda *args: None,
                "open_dial_call_popup": lambda *args: None,
                "send_command_async": original_sender,
            }
            for name, value in overrides.items():
                stack.enter_context(patch("sms_ui.serial_debug_window." + name, value))
            open_serial_debug_window_namespace_runtime(namespace)
            for action, argument in (("dial", "10086"), ("quick_send", "ATD10086;")):
                captured[action](argument)
                for worker in workers + list(namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot()):
                    worker.join(2)
                    self.assertFalse(worker.is_alive())

        self.assertEqual(namespace["serial_obj"].writes, [b"ATD10086;\r\n"] * 2)
        queued = list(namespace["CLOUD_SMS_EVENT_Q"].queue)
        self.assertEqual(len(queued), 2, "Local modem-confirmed dials have no independent history event")
        self.assertEqual(len({payload["source_event_id"] for payload in queued}), 2)
        for payload in queued:
            self.assertEqual(payload["type"], "call_event")
            self.assertEqual(payload["direction"], "outgoing")
            self.assertEqual(payload["phase"], "outgoing")
            self.assertEqual(payload["imei"], IMEI)
            self.assertTrue(payload["ack_required"])


if __name__ == "__main__":
    unittest.main()
