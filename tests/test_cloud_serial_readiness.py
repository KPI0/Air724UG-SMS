import asyncio
import threading
import time
import unittest
from concurrent.futures import Future
from types import SimpleNamespace
from unittest.mock import patch

from sms_core.cloud_payloads import build_register_payload, build_status_payload
from sms_core.cloud_connection_runtime import send_cloud_payload_runtime
from sms_core.cloud_state_namespace_runtime import (
    cloud_status_payload_namespace_runtime,
    notify_cloud_channel_status_namespace_runtime,
)
from sms_core.cloud_ws_namespace_runtime import send_cloud_register_namespace_runtime
from sms_core.serial_startup_runtime import open_and_initialize_serial_runtime, SerialInitializationSession
from tests.test_serial_initialization import InitializationHarness


def channel_fixture(modem=None):
    sent = []
    namespace = dict(
        serial_lock=threading.RLock(),
        serial_obj=modem or SimpleNamespace(is_open=True),
        serial_connection_generation=7,
        cloud_ws_conn=object(), cloud_ws_loop=SimpleNamespace(is_running=lambda: True),
        cloud_connected=True, cloud_device_authorized=True,
        cloud_stop_event=threading.Event(),
        CLOUD_AUTO_UPLOAD=True, CLOUD_DEVICE_SECRET="synthetic-secret",
        CLOUD_WS_URL="wss://example.invalid/ws/device", PORT="TEST", BAUD=115200, MODE="Manual",
        _cloud_now_ts=lambda: 1,
        _cloud_identity_payload=lambda: {"imei": "000000000000001"},
        _cloud_runtime_imei=lambda: "000000000000001",
        _cloud_log=lambda *args, **kwargs: None,
        _cloud_build_register_payload=build_register_payload,
        _cloud_build_status_payload=build_status_payload,
        _cloud_send_payload=lambda ws, payload: sent.append(payload) or ("send", payload),
        asyncio=SimpleNamespace(run_coroutine_threadsafe=lambda *args: None),
    )
    return namespace, sent


class CloudSerialReadinessTests(unittest.TestCase):
    def test_registration_distinguishes_open_serial_from_control_readiness(self):
        for state in ("pending", "quarantined", "recovery", "ready"):
            with self.subTest(state=state):
                ns, _ = channel_fixture()
                modem = ns["serial_obj"]
                if state in ("pending", "quarantined"):
                    modem._serial_initialization_session = SimpleNamespace(
                        pending=state == "pending", quarantined=state == "quarantined",
                    )
                if state == "recovery":
                    modem._sms_recovery_error = "synthetic unconfirmed transaction"
                sent = []

                async def register(ws, **kwargs):
                    sent.append(kwargs)
                    return "sent"

                asyncio.run(send_cloud_register_namespace_runtime(ns, ns["cloud_ws_conn"], send_runtime=register))
                self.assertTrue(sent[0]["serial_connected"])
                self.assertEqual(sent[0]["control_available"], state == "ready")
                self.assertIsInstance(sent[0]["channel_status_revision"], int)

    def test_same_connection_readiness_changes_are_sent_with_ordered_revisions(self):
        ns, sent = channel_fixture()
        modem = ns["serial_obj"]
        modem._serial_initialization_session = SimpleNamespace(pending=True, quarantined=False)
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        self.assertFalse(notify_cloud_channel_status_namespace_runtime(ns))
        modem._serial_initialization_session = None
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        modem._sms_recovery_error = "synthetic unconfirmed transaction"
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        self.assertEqual([item["serial_connected"] for item in sent], [True, True, True])
        self.assertEqual([item["control_available"] for item in sent], [False, True, False])
        revisions = [item["channel_status_revision"] for item in sent]
        self.assertEqual(revisions, sorted(set(revisions)))
        self.assertEqual([item["serial_connection_generation"] for item in sent], [7, 7, 7])
        self.assertEqual(ns["serial_connection_generation"], 7)

    def test_status_response_keeps_logs_connected_but_marks_quarantine_unavailable(self):
        ns, _ = channel_fixture()
        ns["serial_obj"]._sms_recovery_error = "synthetic unconfirmed transaction"
        payload = cloud_status_payload_namespace_runtime(ns)
        self.assertTrue(payload["serial_connected"])
        self.assertFalse(payload["control_available"])
        self.assertEqual(payload["serial_connection_generation"], 7)
        self.assertIsInstance(payload["channel_status_revision"], int)

    def test_registration_cannot_hide_a_newer_channel_change_after_await(self):
        ns, sent = channel_fixture()

        async def register(ws, **kwargs):
            ns["serial_obj"]._sms_recovery_error = "synthetic unconfirmed transaction"
            notify_cloud_channel_status_namespace_runtime(ns)
            return "sent"

        asyncio.run(send_cloud_register_namespace_runtime(ns, ns["cloud_ws_conn"], send_runtime=register))
        self.assertFalse(notify_cloud_channel_status_namespace_runtime(ns))
        self.assertEqual(len(sent), 1)
        self.assertFalse(sent[0]["control_available"])

    def test_new_websocket_republishes_current_readiness(self):
        ns, sent = channel_fixture()
        ns["serial_obj"]._sms_recovery_error = "synthetic unconfirmed transaction"
        notify_cloud_channel_status_namespace_runtime(ns)
        ns["cloud_ws_conn"] = object()
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        self.assertEqual(len(sent), 2)
        self.assertFalse(sent[-1]["control_available"])

    def test_registration_cache_does_not_suppress_a_return_to_its_original_state(self):
        ns, sent = channel_fixture()
        modem = ns["serial_obj"]

        async def register(ws, **kwargs):
            modem._sms_recovery_error = "synthetic unconfirmed transaction"
            notify_cloud_channel_status_namespace_runtime(ns)
            modem._sms_recovery_error = ""
            cloud_status_payload_namespace_runtime(ns)
            return "sent"

        asyncio.run(send_cloud_register_namespace_runtime(ns, ns["cloud_ws_conn"], send_runtime=register))
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        self.assertEqual([item["control_available"] for item in sent], [False, True])
        self.assertGreater(sent[-1]["channel_status_revision"], sent[0]["channel_status_revision"])

    def test_stopping_serial_cannot_advertise_control_even_before_handle_closes(self):
        ns, sent = channel_fixture()
        ns["serial_stop_event"] = threading.Event()
        ns["serial_stop_event"].set()
        notify_cloud_channel_status_namespace_runtime(ns)
        self.assertTrue(sent[-1]["serial_connected"])
        self.assertFalse(sent[-1]["control_available"])

    def test_new_serial_is_not_observable_as_ready_before_its_initialization_owner_exists(self):
        ns, _ = channel_fixture()
        ns["serial_obj"] = None
        modem = SimpleNamespace(is_open=True)
        constructing, release, finished, cleanup = (threading.Event() for _ in range(4))
        snapshots, errors = [], []

        def initialization(*args, **kwargs):
            constructing.set()
            if not release.wait(2):
                raise RuntimeError("synthetic construction gate timed out")
            return SerialInitializationSession(*args, **kwargs)

        def opener():
            try:
                open_and_initialize_serial_runtime(
                    target_port="TEST", baud=115200, mode="Manual", serial_lock=ns["serial_lock"],
                    open_serial=lambda *args: modem,
                    set_serial_obj=lambda serial: ns.__setitem__("serial_obj", serial),
                    set_port=lambda *args: None, lock_port_mutex=lambda *args: None,
                    set_cloud_imei_query_deadline=lambda *args: None,
                    serial_error_ui=lambda *args: None, set_status=lambda *args: None,
                    response_coordinator=SimpleNamespace(), get_serial_obj=lambda: ns["serial_obj"],
                    get_generation=lambda: ns["serial_connection_generation"],
                )
                cleanup.wait(2)
                modem._serial_initialization_session.close()
            except Exception as exc:
                errors.append(type(exc).__name__)

        def observer():
            try:
                snapshots.append(cloud_status_payload_namespace_runtime(ns))
            finally:
                finished.set()

        with patch("sms_core.serial_startup_runtime.SerialInitializationSession", initialization):
            open_thread = threading.Thread(target=opener)
            open_thread.start()
            self.assertTrue(constructing.wait(2))
            observe_thread = threading.Thread(target=observer)
            observe_thread.start()
            try:
                self.assertFalse(finished.wait(0.05), "registration saw an unowned open serial as ready")
            finally:
                release.set()
                observe_thread.join(2)
                cleanup.set()
                open_thread.join(2)
        self.assertEqual(errors, [])
        self.assertTrue(snapshots[0]["serial_connected"])
        self.assertFalse(snapshots[0]["control_available"])

    def test_immediately_completed_send_failure_allows_next_read_boundary_to_retry(self):
        ns, _ = channel_fixture()
        future = Future()
        future.set_result("error")
        ns["asyncio"] = SimpleNamespace(run_coroutine_threadsafe=lambda *args: future)
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))


class ChannelPublicationFailureTests(unittest.IsolatedAsyncioTestCase):
    def async_channel_fixture(self):
        ns, _ = channel_fixture()
        futures = []

        def schedule(coro, loop):
            future = asyncio.run_coroutine_threadsafe(coro, loop)
            futures.append(future)
            return future

        ns.update(cloud_ws_loop=asyncio.get_running_loop(),
            asyncio=SimpleNamespace(run_coroutine_threadsafe=schedule))
        return ns, futures

    async def test_actual_send_failure_is_retried_without_reopening_the_websocket(self):
        ns, futures = self.async_channel_fixture()

        class FirstSendFails:
            calls = 0

            async def send(self, payload):
                self.calls += 1
                if self.calls == 1:
                    raise OSError("synthetic transient send failure")

        ws = FirstSendFails()
        ns.update(cloud_ws_conn=ws, _cloud_send_payload=send_cloud_payload_runtime)
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        self.assertEqual(await asyncio.wrap_future(futures[-1]), "error")
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        self.assertEqual(await asyncio.wrap_future(futures[-1]), "sent")
        self.assertFalse(notify_cloud_channel_status_namespace_runtime(ns))
        self.assertEqual(ws.calls, 2)
        self.assertTrue(ns["cloud_connected"])
        self.assertTrue(ns["cloud_device_authorized"])

    async def test_old_send_failure_does_not_clear_a_newer_successful_state(self):
        ns, futures = self.async_channel_fixture()
        entered, release = asyncio.Event(), asyncio.Event()

        async def send(ws, payload):
            if payload["channel_status_revision"] == 1:
                entered.set()
                await release.wait()
                return "error"
            return "sent"

        ns["_cloud_send_payload"] = send
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        old_future = futures[-1]
        await asyncio.wait_for(entered.wait(), 1)
        ns["serial_obj"]._sms_recovery_error = "synthetic unconfirmed transaction"
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        self.assertEqual(await asyncio.wrap_future(futures[-1]), "sent")
        release.set()
        self.assertEqual(await asyncio.wrap_future(old_future), "error")
        self.assertFalse(notify_cloud_channel_status_namespace_runtime(ns))

    async def test_old_send_failure_does_not_clear_a_successful_registration_of_the_same_state(self):
        ns, futures = self.async_channel_fixture()
        entered, release = asyncio.Event(), asyncio.Event()

        async def send(ws, payload):
            entered.set()
            await release.wait()
            return "error"

        async def register(ws, **kwargs):
            return "sent"

        ns["_cloud_send_payload"] = send
        self.assertTrue(notify_cloud_channel_status_namespace_runtime(ns))
        old_future = futures[-1]
        await asyncio.wait_for(entered.wait(), 1)
        self.assertEqual(await send_cloud_register_namespace_runtime(
            ns, ns["cloud_ws_conn"], send_runtime=register), "sent")
        release.set()
        self.assertEqual(await asyncio.wrap_future(old_future), "error")
        self.assertFalse(notify_cloud_channel_status_namespace_runtime(ns))


class ReaderChannelReadinessTests(unittest.TestCase):
    def wait_for(self, predicate):
        deadline = time.monotonic() + 2
        while time.monotonic() < deadline:
            if predicate():
                return
            threading.Event().wait(0.005)
        self.fail("serial reader did not publish the readiness transition")

    def test_normal_initialization_publishes_ready_without_waiting_for_remote_command(self):
        harness = InitializationHarness(hold="AT+CNUM")
        channel_ns, sent = channel_fixture(harness.modem)
        harness.ns.update({key: value for key, value in channel_ns.items() if key not in (
            "serial_obj", "serial_lock", "serial_connection_generation", "PORT", "_cloud_runtime_imei",
            "_cloud_identity_payload",
        )})
        harness.ns["notify_cloud_channel_status"] = lambda connected=None: notify_cloud_channel_status_namespace_runtime(
            harness.ns, connected,
        )
        with harness:
            self.assertTrue(harness.wrote["AT+CNUM"].wait(1))
            self.assertFalse(sent[0]["control_available"])
            self.wait_for(lambda: harness.ns["cloud_imei_verified"])
            harness.ns["cloud_device_authorized"] = True
            harness.terminal("AT+CNUM")
            self.wait_for(lambda: any(item["control_available"] for item in sent))
            self.assertEqual(harness.errors, [])

    def test_reader_republishes_local_quarantine_while_logs_still_arrive(self):
        harness = InitializationHarness(hold="unused")
        channel_ns, sent = channel_fixture(harness.modem)
        harness.ns.update({key: value for key, value in channel_ns.items() if key not in (
            "serial_obj", "serial_lock", "serial_connection_generation", "PORT", "_cloud_runtime_imei",
            "_cloud_identity_payload",
        )})
        harness.ns["notify_cloud_channel_status"] = lambda connected=None: notify_cloud_channel_status_namespace_runtime(
            harness.ns, connected,
        )
        with harness:
            self.wait_for(lambda: harness.ns["cloud_imei_verified"])
            harness.ns["cloud_device_authorized"] = True
            self.wait_for(lambda: any(item["control_available"] for item in sent))
            harness.modem._sms_recovery_error = "synthetic unconfirmed transaction"
            harness.emit("[I]-[ril.sendat] AT+CSQ\r\n")
            self.wait_for(lambda: sent[-1]["control_available"] is False)
            self.assertTrue(sent[-1]["serial_connected"])
            self.wait_for(lambda: any("AT+CSQ" in line for line in harness.lines))
            self.assertEqual(harness.errors, [])


if __name__ == "__main__":
    unittest.main()
