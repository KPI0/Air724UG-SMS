import queue
import threading
import unittest
from unittest.mock import patch

from sms_core.cloud_payloads import build_sms_event_payload, identity_payload
from sms_core.cloud_sms_event_runtime import CloudSmsEventDrainState
from sms_core.serial_sender import SmsPduSendCoordinator, send_text_sms_pdu_async
from sms_core.sms_send_namespace_runtime import send_manual_sms_namespace_runtime
from sms_core.threading_runtime import WorkerThreadRegistry


IMEI = "000000000000001"
OTHER_IMEI = "000000000000002"


class ModemSerial:
    def __init__(self, coordinator, results):
        self.coordinator, self.results = coordinator, list(results)
        self.is_open, self.writes = True, []

    def write(self, payload):
        self.writes.append(payload)
        if payload.startswith(b"AT+CMGS="):
            self.coordinator.observe_line(">")
        elif payload.endswith(b"\x1a"):
            result = self.results.pop(0)
            if result is True:
                self.coordinator.observe_line("+CMGS: 17")
                self.coordinator.observe_line("OK")
            elif result is False:
                self.coordinator.observe_line("+CMS ERROR: 500")

    def flush(self):
        pass


class SmsSentRuntimeTests(unittest.TestCase):
    def namespace(self):
        state = {
            "serial_lock": threading.RLock(), "serial_obj": object(),
            "serial_connection_generation": 1, "imei": IMEI,
            "CLOUD_CONTROL_ENABLED": True, "cloud_connected": False,
            "cloud_device_authorized": False, "cloud_ws_loop": None, "cloud_ws_conn": None,
            "CLOUD_SMS_EVENT_Q": queue.Queue(maxsize=10),
            "CLOUD_SMS_EVENT_DRAIN_STATE": CloudSmsEventDrainState(),
            "_cloud_now_ts": lambda: 1789171200,
            "_push_serial_debug": lambda *_args: None, "port_ui": lambda *_args: None,
            "_schedule_cloud_sms_event_drain": lambda *_args: None,
        }
        state["_cloud_runtime_imei"] = lambda: state["imei"]
        state["_cloud_identity_payload"] = lambda: identity_payload(state["imei"], "test", "test-device")
        return state

    def prepare(self, namespace, message="测试发送正文"):
        captured = {}

        def capture(lock, get_serial, phone, body, **kwargs):
            captured.update(get_serial=get_serial, phone=phone, body=body, **kwargs)
            return "worker"

        self.assertEqual(send_manual_sms_namespace_runtime(
            namespace, "10086", message, send_runtime=capture,
        ), "worker")
        return captured

    def test_receive_payload_explicitly_marks_incoming(self):
        payload = build_sms_event_payload("", "接收正文", 1, {"imei": IMEI})
        self.assertEqual(payload["direction"], "incoming")

    def test_submission_is_not_recorded_until_success_callback(self):
        namespace = self.namespace()
        send = self.prepare(namespace)
        self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())
        self.assertEqual(send["on_sent"](), "queued_offline")
        payload = namespace["CLOUD_SMS_EVENT_Q"].get_nowait()
        self.assertEqual(payload["direction"], "outgoing")
        self.assertEqual(payload["imei"], IMEI)
        self.assertEqual(payload["phone"], "10086")
        self.assertEqual(payload["content"], "测试发送正文")

    def test_sent_body_is_preserved_including_empty_text(self):
        for body in ("", "  第一行\n26/09/12,08:00:00+32 原文  "):
            with self.subTest(body=body):
                namespace = self.namespace()
                send = self.prepare(namespace, body)
                send["on_sent"]()
                self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].get_nowait()["content"], body)

    def test_success_after_device_switch_keeps_original_imei(self):
        namespace = self.namespace()
        send = self.prepare(namespace)
        namespace.update(imei=OTHER_IMEI, serial_obj=object(), serial_connection_generation=2)
        send["on_sent"]()
        self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].get_nowait()["imei"], IMEI)

    def test_switched_online_device_cannot_drain_old_device_send(self):
        namespace = self.namespace()
        scheduled = []
        namespace.update(cloud_connected=True, cloud_device_authorized=True,
                         cloud_ws_conn=object(), cloud_ws_loop=type("Loop", (), {"is_running": lambda _self: True})())
        namespace["_schedule_cloud_sms_event_drain"] = lambda *args: scheduled.append(args)
        send = self.prepare(namespace)
        namespace["imei"] = OTHER_IMEI
        self.assertEqual(send["on_sent"](), "queued_offline")
        self.assertEqual(scheduled, [])
        self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].get_nowait()["imei"], IMEI)

    def test_unverified_identity_does_not_use_later_devices_imei(self):
        namespace = self.namespace()
        namespace["imei"] = ""
        send = self.prepare(namespace)
        self.assertIs(send["get_serial"](), namespace["serial_obj"])
        namespace["imei"] = OTHER_IMEI
        self.assertEqual(send["on_sent"](), "missing_imei")
        self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())

    def test_queued_send_cannot_target_replacement_serial_or_identity(self):
        for change in ({"serial_obj": object()}, {"serial_connection_generation": 2}, {"imei": OTHER_IMEI}):
            with self.subTest(change=list(change)):
                namespace = self.namespace()
                send = self.prepare(namespace)
                self.assertIs(send["get_serial"](), namespace["serial_obj"])
                namespace.update(change)
                with self.assertRaises(RuntimeError):
                    send["get_serial"]()
                self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())

    def test_independent_manual_sends_have_different_ids(self):
        namespace = self.namespace()
        for _ in range(2):
            self.prepare(namespace)["on_sent"]()
        first = namespace["CLOUD_SMS_EVENT_Q"].get_nowait()
        second = namespace["CLOUD_SMS_EVENT_Q"].get_nowait()
        self.assertNotEqual(first["source_event_id"], second["source_event_id"])

    def test_cloud_disabled_does_not_enqueue(self):
        namespace = self.namespace()
        send = self.prepare(namespace)
        namespace["CLOUD_CONTROL_ENABLED"] = False
        self.assertEqual(send["on_sent"](), "disabled")
        self.assertTrue(namespace["CLOUD_SMS_EVENT_Q"].empty())

    def run_modem_send(self, results, on_sent, body="短" * 100, log_error=None):
        coordinator = SmsPduSendCoordinator()
        serial = ModemSerial(coordinator, results)
        registry = WorkerThreadRegistry()
        thread = send_text_sms_pdu_async(
            threading.RLock(), lambda: serial, "10086", body,
            response_coordinator=coordinator, sleep_func=lambda _seconds: None,
            segment_timeout=0.01, prompt_timeout=0.01, thread_registry=registry,
            on_sent=on_sent, log_error=log_error,
        )
        thread.join(2)
        self.assertFalse(thread.is_alive())
        self.assertEqual(registry.snapshot(), ())
        return serial

    def test_all_segments_confirmed_produce_exactly_one_sent_callback(self):
        sent = []
        serial = self.run_modem_send([True, True], lambda: sent.append(True))
        self.assertEqual(sent, [True])
        self.assertEqual(sum(payload.endswith(b"\x1a") for payload in serial.writes), 2)

    def test_partial_failure_and_timeout_produce_no_sent_record(self):
        for results in ([True, False], [True, None]):
            with self.subTest(results=results):
                sent = []
                self.run_modem_send(results, lambda: sent.append(True))
                self.assertEqual(sent, [])

    def test_failed_record_callback_does_not_resend_or_leak_error_text(self):
        logs = []

        def fail_record():
            raise RuntimeError("private callback data")

        serial = self.run_modem_send([True], fail_record, body="短信", log_error=logs.append)
        self.assertEqual(sum(payload.endswith(b"\x1a") for payload in serial.writes), 1)
        self.assertEqual(logs, ["短信已发送，但发送记录更新失败"])

    def test_serial_window_wires_manual_send_handler(self):
        from sms_ui.serial_debug_namespace_runtime import open_serial_debug_window_namespace_runtime
        from tests.test_serial_debug_namespace_runtime import SerialDebugNamespaceRuntimeTests
        namespace = SerialDebugNamespaceRuntimeTests().base_namespace()
        forwarded = {}

        def open_window(_root, **kwargs):
            forwarded.update(kwargs)
            return None, None

        open_serial_debug_window_namespace_runtime(namespace, open_window_runtime=open_window)
        with patch("sms_ui.serial_debug_namespace_runtime.send_manual_sms_namespace_runtime") as send:
            forwarded["send_sms"]("10086", "测试正文")
            send.assert_called_once_with(namespace, "10086", "测试正文")


if __name__ == "__main__":
    unittest.main()
