import asyncio
import inspect
import json
import threading
import unittest
from contextlib import suppress
from unittest.mock import patch

from sms_core.cloud_command_context import capture_cloud_serial_command_context, set_cloud_command_authorized
from sms_core.cloud_device_config import (
    CloudDeviceConfigRelay, UsbDeviceConfigTransport, handle_device_config,
    safe_config_reply, stop_device_config, validate_config_request,
)
from sms_core.cloud_firmware_ota import RelayError
from sms_core.cloud_message_runtime import handle_cloud_message_runtime
from sms_core.cloud_message_namespace_runtime import handle_cloud_message_namespace_runtime
from sms_core.cloud_payloads import build_register_payload, build_channel_status_payload
from sms_core import serial_sender
from test_cloud_firmware_ota import Clock, IMEI, Usb, WebSocket, namespace, port


def request(action="info", **fields):
    data = dict(type="device_config", protocol=1, action=action, request_id="1" * 32,
                transport="desktop", imei=IMEI, serial_connection_generation=1)
    if action != "info":
        data["expected_revision"] = 2
    if action == "set":
        data["configs"] = {"CONSOLE_CONNECTION_ENABLE": False}
    data.update(fields)
    return data


def config_reply(req, reply):
    reply.update(type="device_config_reply", protocol=1, action=req["action"], version="1.0.18",
                 supported=True, revision=2, values={"CONSOLE_CONNECTION_ENABLE": False},
                 configured={"DEVICE_SECRET": True})
    reply.pop("ota", None)
    if req["action"] == "set":
        reply.update(saved=True, applied=True, revision=3, restart_required=True)
    if req["action"] == "restart":
        reply["phase"] = "accepted"
    return reply


class UsbDeviceConfigTests(unittest.TestCase):
    def setUp(self):
        self.ns = namespace()
        self.context = capture_cloud_serial_command_context(self.ns)
        self.opened = []
        self.transports = []

    def tearDown(self):
        for transport in self.transports:
            transport.close()

    def transport(self, handler=config_reply, *, imei=IMEI, partial=False, current=None):
        usb = Usb(imei=imei, handler=handler, partial=partial)
        def open_port(name, **kwargs):
            self.opened.append(name)
            self.assertEqual(kwargs["write_timeout"], 2)
            return usb
        transport = UsbDeviceConfigTransport(self.context,
            current or (lambda: not self.context.rejection_reason(self.ns)),
            serial_factory=open_port, list_ports=lambda: [
                port("COM1", "LUAT USB Device 0 Modem"), port("COM2", "LUAT USB Device 1 AT")],
            clock=Clock(), response_timeout=0.3)
        self.transports.append(transport)
        return transport, usb

    def test_info_reads_disabled_direct_connection_without_any_secret(self):
        transport, usb = self.transport(partial=True)
        reply = transport.exchange(request())
        self.assertFalse(reply["values"]["CONSOLE_CONNECTION_ENABLE"])
        self.assertEqual(reply["revision"], 2)
        self.assertEqual(self.opened, ["COM2"])
        self.assertTrue(usb.closed)
        self.assertEqual(len(usb.requests), 1)
        self.assertEqual(usb.requests[0]["type"], "device_config")
        self.assertNotIn("secret", usb.requests[0])
        self.assertEqual(usb.requests[0]["imei"], IMEI)
        self.assertEqual(len(usb.requests[0]["session"]), 32)

    def test_info_passes_device_support_and_effective_settings_over_verified_usb(self):
        supported = ["ALLOW_NVM_SERVER_CONFIG", "CELLULAR_DATA_FORCE_DETACH",
                     "CELLULAR_DATA_DEACTIVATE_PDP", "PIN_CODE"]
        def handler(req, reply):
            reply = config_reply(req, reply)
            reply.update(supported_fields=supported, file_only={},
                         effective_values={"ALLOW_NVM_SERVER_CONFIG": False,
                                           "PIN_CODE": "synthetic-private"})
            return reply
        transport, usb = self.transport(handler, partial=True)
        reply = transport.exchange(request())
        self.assertEqual(reply["supported_fields"], supported)
        self.assertEqual(reply["file_only"], [])
        self.assertEqual(reply["effective_values"], {"ALLOW_NVM_SERVER_CONFIG": False})
        self.assertNotIn("synthetic-private", json.dumps(reply))
        self.assertEqual([item["action"] for item in usb.requests], ["info"])
        self.assertTrue(usb.closed)

    def test_info_passes_server_batch_capability_from_device_and_ignores_browser_claims(self):
        for supported in (True, False):
            with self.subTest(supported=supported):
                def handler(req, reply):
                    reply = config_reply(req, reply)
                    reply["server_config_batch"] = supported
                    return reply
                transport, usb = self.transport(handler, partial=True)
                reply = transport.exchange(request(server_config_batch=not supported))
                self.assertIs(reply["server_config_batch"], supported)
                self.assertTrue(all("server_config_batch" not in item for item in usb.requests))
                self.assertEqual([item["action"] for item in usb.requests], ["info"])
                self.assertTrue(usb.closed)

    def test_bad_server_batch_capability_stops_before_submitting_config(self):
        def handler(req, reply):
            reply = config_reply(req, reply)
            reply["server_config_batch"] = "true"
            return reply
        transport, usb = self.transport(handler, partial=True)
        with self.assertRaisesRegex(RelayError, "usb_unavailable"):
            transport.exchange(request("set", server_config_batch=True,
                configs={"ALLOW_NVM_SERVER_CONFIG": True, "DEVICE_SECRET": "synthetic-new"}))
        self.assertEqual([item["action"] for item in usb.requests], ["info"])
        self.assertNotIn("synthetic-new", json.dumps(usb.requests))
        self.assertTrue(usb.closed)

    def test_set_probes_identity_then_transmits_explicit_false_empty_and_template_once(self):
        transport, usb = self.transport(partial=True)
        configs = {"CONSOLE_CONNECTION_ENABLE": False, "CELLULAR_DATA_ENABLE": False,
                   "DEVICE_SECRET": "", "CUSTOM_POST_BODY_TABLE": {"text": "${msg}"}}
        reply = transport.exchange(request("set", configs=configs))
        self.assertTrue(reply["saved"])
        self.assertEqual(reply["revision"], 3)
        self.assertEqual([item["action"] for item in usb.requests], ["info", "set"])
        self.assertEqual(usb.requests[1]["configs"], configs)
        self.assertNotIn("secret", usb.requests[0])
        self.assertNotIn("secret", usb.requests[1])
        self.assertTrue(usb.closed)

    def test_wrong_imei_or_session_never_submits_config_or_a_secret(self):
        for broken in ("imei", "session", "request_id", "action", "protocol"):
            with self.subTest(broken=broken):
                def handler(req, reply):
                    reply = config_reply(req, reply)
                    reply[broken] = {"imei": "000000000000002", "session": "0" * 32,
                                     "request_id": "0" * 32, "action": "set", "protocol": True}[broken]
                    return reply
                transport, usb = self.transport(handler)
                with self.assertRaises(RelayError):
                    transport.exchange(request("set", configs={"DEVICE_SECRET": "synthetic-new"}))
                self.assertTrue(usb.closed)
                self.assertTrue(all(item["action"] == "info" for item in usb.requests))
                self.assertNotIn("synthetic-new", json.dumps(usb.requests))

    def test_config_conflict_is_preserved_without_second_set(self):
        def handler(req, reply):
            reply = config_reply(req, reply)
            if req["action"] == "set":
                reply.update(ok=False, reason="config_conflict", revision=4, saved=False)
            return reply
        transport, usb = self.transport(handler)
        reply = transport.exchange(request("set"))
        self.assertFalse(reply["ok"])
        self.assertEqual(reply["reason"], "config_conflict")
        self.assertEqual(sum(item["action"] == "set" for item in usb.requests), 1)

    def test_set_or_restart_timeout_is_uncertain_and_never_retries_write(self):
        for action in ("set", "restart"):
            with self.subTest(action=action):
                def handler(req, reply):
                    return config_reply(req, reply) if req["action"] == "info" else None
                transport, usb = self.transport(handler)
                with self.assertRaisesRegex(RelayError, "config_unconfirmed" if action == "set" else "restart_unconfirmed"):
                    transport.exchange(request(action))
                self.assertEqual([item["action"] for item in usb.requests], ["info", action])
                self.assertTrue(usb.closed)

    def test_restart_requires_explicit_accepted_receipt(self):
        transport, usb = self.transport()
        self.assertEqual(transport.exchange(request("restart"))["phase"], "accepted")
        self.assertEqual([item["action"] for item in usb.requests], ["info", "restart"])
        def handler(req, reply):
            reply = config_reply(req, reply)
            reply.pop("phase", None)
            return reply
        transport, _usb = self.transport(handler)
        with self.assertRaisesRegex(RelayError, "restart_unconfirmed"):
            transport.exchange(request("restart"))

    def test_set_success_requires_confirmed_save_and_revision(self):
        for changes in ({"saved": False}, {"revision": True}, {"revision": -1}):
            def handler(req, reply):
                reply = config_reply(req, reply)
                if req["action"] == "set":
                    reply.update(changes)
                return reply
            transport, _usb = self.transport(handler)
            with self.assertRaisesRegex(RelayError, "config_unconfirmed"):
                transport.exchange(request("set"))

    def test_authorization_change_after_probe_stops_set_and_releases_lease(self):
        def handler(req, reply):
            set_cloud_command_authorized(self.ns, False)
            return config_reply(req, reply)
        transport, usb = self.transport(handler)
        with self.assertRaisesRegex(RelayError, "connection_changed"):
            transport.exchange(request("set"))
        self.assertEqual([item["action"] for item in usb.requests], ["info"])
        self.assertTrue(serial_sender.write_serial_command_result(self.context.serial, "ATI").ok)

    def test_config_lease_blocks_modem_commands_and_releases_after_save(self):
        attempts = []
        def handler(req, reply):
            attempts.append(serial_sender.write_serial_command_result(self.context.serial, "ATI").ok)
            return config_reply(req, reply)
        transport, _usb = self.transport(handler)
        transport.exchange(request("set"))
        self.assertEqual(attempts, [False, False])
        self.assertEqual(self.context.serial.writes, [])
        self.assertTrue(serial_sender.write_serial_command_result(self.context.serial, "ATI").ok)

    def test_active_ota_owner_prevents_opening_config_usb(self):
        owner = object()
        self.assertTrue(serial_sender.reserve_serial_for_ota(self.context.serial, owner))
        try:
            transport, _usb = self.transport()
            with self.assertRaisesRegex(RelayError, "serial_busy"):
                transport.exchange(request("set"))
            self.assertEqual(self.opened, [])
            self.assertFalse(serial_sender.write_serial_command_result(self.context.serial, "ATI").ok)
        finally:
            serial_sender.release_serial_from_ota(self.context.serial, owner)

    def test_bad_revision_or_values_are_rejected_before_usb(self):
        cases = [request("set", expected_revision=value) for value in (True, -1, "2", 10 ** 1000)]
        cases += [request("set", configs=value) for value in ({}, [], {"DEVICE_SECRET": float("nan")},
                                                               {"bad_key": "value"}, {"DEVICE_SECRET": "x" * 6200})]
        for data in cases:
            with self.subTest(kind=type(data.get("configs")).__name__):
                transport, _usb = self.transport()
                with self.assertRaises(RelayError):
                    transport.exchange(data)
        self.assertEqual(self.opened, [])

    def test_reply_hides_credentials_templates_and_arbitrary_diagnostics(self):
        reply = safe_config_reply(dict(ok=True, protocol=1, values={"DEVICE_SECRET": "synthetic-private",
            "WEBSOCKET_URL": "wss://example.invalid/private", "PIN_CODE": "1234",
            "CUSTOM_POST_BODY_TABLE": {"password": "synthetic-private"}, "CALL_MAX_TIME": 60},
            configured={"DEVICE_SECRET": True, "PIN_CODE": "synthetic-private"},
            apply_errors={"DEVICE_SECRET": "synthetic-private"}, reason="synthetic-private"))
        self.assertEqual(reply["values"], {"CALL_MAX_TIME": 60})
        self.assertEqual(reply["configured"], {"DEVICE_SECRET": True})
        self.assertEqual(reply["apply_errors"], {"DEVICE_SECRET": "config_apply"})
        self.assertNotIn("synthetic-private", json.dumps(reply))


class ControlledConfigTransport:
    def __init__(self, context, current):
        self.context, self.current = context, current
        self.stopped = threading.Event()
        self.entered, self.release, self.finished = threading.Event(), threading.Event(), threading.Event()
        self.closed = False
        self.requests = []

    def exchange(self, req):
        self.entered.set()
        self.release.wait(3)
        try:
            if not self.current() or self.stopped.is_set():
                raise RelayError("connection_changed")
            self.requests.append(req)
            return config_reply(req, dict(ok=True))
        finally:
            self.finished.set()

    def close(self):
        self.closed = True


class CloudDeviceConfigTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.ws = WebSocket()
        self.ns = namespace(self.ws)
        self.created = []
        def factory(context, current):
            transport = ControlledConfigTransport(context, current)
            self.created.append(transport)
            return transport
        self.relay = CloudDeviceConfigRelay(self.ns, self.ws, transport_factory=factory)

    async def asyncTearDown(self):
        for transport in self.created:
            transport.release.set()
        await self.relay.stop()

    async def start(self, data):
        await self.relay.submit(data)
        await asyncio.sleep(0)
        self.assertTrue(await asyncio.to_thread(self.created[-1].entered.wait, 1))
        return self.created[-1]

    async def test_invalid_context_does_not_start_config_worker(self):
        for fields in ({"imei": "000000000000002"}, {"transport": "firmware"},
                       {"serial_connection_generation": 2}, {"serial_connection_generation": True},
                       {"request_id": "invalid"}):
            await self.relay.submit(request("set", **fields))
        set_cloud_command_authorized(self.ns, False)
        await self.relay.submit(request("set"))
        self.assertEqual(self.created, [])

    async def test_invalid_protocol_action_or_patch_returns_safe_failure(self):
        for fields in ({"protocol": True}, {"action": {"secret": "synthetic-private"}},
                       {"configs": {"bad": "synthetic-private"}}, {"expected_revision": True}):
            data = request("set")
            data.update(fields)
            await self.relay.submit(data)
            self.assertEqual(self.ws.replies[-1]["reason"], "invalid_request")
            self.assertNotIn("synthetic-private", json.dumps(self.ws.replies[-1]))
        self.assertEqual(self.created, [])

    async def test_duplicate_queued_request_runs_once_and_closes_before_reply(self):
        transport = await self.start(request("set"))
        await self.relay.submit(request("set"))
        original = self.ws.send
        async def send(raw):
            self.assertTrue(transport.closed)
            await original(raw)
        self.ws.send = send
        transport.release.set()
        await self.ws.received()
        await self.relay.task
        self.assertEqual(len(transport.requests), 1)
        self.assertEqual(self.ws.replies[0]["request_id"], "1" * 32)
        self.assertEqual(self.ws.replies[0]["serial_connection_generation"], 1)

    async def test_duplicate_during_reply_backpressure_never_repeats_set(self):
        transport = await self.start(request("set"))
        entered, release = asyncio.Event(), asyncio.Event()
        async def send(_raw):
            self.assertTrue(transport.closed)
            entered.set()
            await release.wait()
        self.ws.send = send
        transport.release.set()
        await asyncio.wait_for(entered.wait(), 1)
        await self.relay.submit(request("set"))
        self.assertEqual(self.relay.queue.qsize(), 0)
        release.set()
        await self.relay.task
        self.assertEqual(len(transport.requests), 1)

    async def test_source_patch_is_detached_after_admission(self):
        data = request("set", configs={"CUSTOM_POST_BODY_TABLE": {"text": "original"}})
        transport = await self.start(data)
        data["configs"]["CUSTOM_POST_BODY_TABLE"]["text"] = "changed"
        transport.release.set()
        await self.relay.task
        self.assertEqual(transport.requests[0]["configs"]["CUSTOM_POST_BODY_TABLE"]["text"], "original")

    async def test_revocation_and_reauthorization_cannot_resume_old_config(self):
        transport = await self.start(request("set"))
        set_cloud_command_authorized(self.ns, False)
        set_cloud_command_authorized(self.ns, True)
        transport.release.set()
        await self.relay.task
        self.assertEqual(transport.requests, [])
        self.assertEqual(self.ws.replies, [])
        self.assertTrue(transport.closed)

    async def test_stop_waits_for_original_worker_before_closing_usb(self):
        transport = await self.start(request("set"))
        stopping = asyncio.create_task(self.relay.stop())
        await asyncio.sleep(0.02)
        self.assertFalse(stopping.done())
        self.assertFalse(transport.closed)
        transport.release.set()
        await asyncio.wait_for(stopping, 1)
        self.assertTrue(transport.finished.is_set())
        self.assertTrue(transport.closed)

    async def test_repeated_cancel_waits_for_worker_and_suppresses_raw_error(self):
        transport = await self.start(request("set"))
        self.relay.task.cancel()
        await asyncio.sleep(0)
        self.relay.task.cancel()
        await asyncio.sleep(0.02)
        self.assertFalse(transport.closed)
        transport.release.set()
        with self.assertRaises(asyncio.CancelledError):
            await self.relay.task
        self.assertTrue(transport.finished.is_set())
        self.assertTrue(transport.closed)

    async def test_io_failure_is_fixed_uncertain_reply_after_handle_close(self):
        def failing_factory(context, current):
            transport = ControlledConfigTransport(context, current)
            def fail(_req):
                raise OSError("synthetic-private-adapter-error")
            transport.exchange = fail
            self.created.append(transport)
            return transport
        relay = CloudDeviceConfigRelay(self.ns, self.ws, transport_factory=failing_factory)
        await relay.submit(request("set"))
        await relay.task
        self.assertEqual(self.ws.replies[0]["reason"], "config_unconfirmed")
        self.assertNotIn("saved", self.ws.replies[0])
        self.assertNotIn("synthetic-private", json.dumps(self.ws.replies))
        self.assertTrue(self.created[-1].closed)

    async def test_configuration_messages_bypass_ordinary_logging(self):
        logged, handled = [], []
        async def handler(data):
            handled.append(data["type"])
        kwargs = {name: lambda *args, **kw: None for name, item in inspect.signature(handle_cloud_message_runtime).parameters.items()
                  if name != "message" and item.default is inspect.Parameter.empty}
        kwargs.update(log=lambda *args, **kw: logged.append(args), handle_device_config_message=handler)
        self.assertEqual(await handle_cloud_message_runtime(json.dumps(request("set", configs={"DEVICE_SECRET": "synthetic-private"})), **kwargs), "device_config")
        self.assertEqual(logged, [])
        self.assertEqual(handled, ["device_config"])

    async def test_namespace_routes_config_without_command_logging(self):
        self.ns.update(_cloud_log=lambda *_a, **_kw: self.fail("config logged"),
            set_cloud_auth_status_from_ack=lambda *_a: None, set_cloud_status=lambda *_a: None,
            _cloud_check_replay_window=lambda *_a, **_kw: True, _cloud_auth_matches=lambda *_a: True,
            _cloud_now_ts=lambda: 1, _cloud_send_status_payload=lambda: {}, show_window=lambda: None, hide_window=lambda: None)
        async def reply(*_args):
            return None
        self.ns["_cloud_reply"] = reply
        with patch("sms_core.cloud_device_config.CloudDeviceConfigRelay", return_value=self.relay):
            self.assertEqual(await handle_cloud_message_namespace_runtime(self.ns, self.ws, json.dumps(request())), "device_config")
        await asyncio.sleep(0)
        self.created[0].release.set()
        await self.relay.task
        self.assertTrue(self.ws.replies[0]["ok"])
        await stop_device_config(self.ns)
        self.assertNotIn("_device_config_relay", self.ns)

    async def test_new_request_waits_for_closing_old_relay(self):
        transport = await self.start(request("set"))
        self.relay.closing = True
        self.ns["_device_config_relay"] = self.relay
        with patch("sms_core.cloud_device_config.CloudDeviceConfigRelay", return_value=self.relay) as create:
            pending = asyncio.create_task(handle_device_config(self.ns, self.ws, request()))
            await asyncio.sleep(0.02)
            create.assert_not_called()
            transport.release.set()
            await asyncio.wait_for(pending, 1)
            self.assertTrue(transport.closed)
        with suppress(asyncio.CancelledError):
            await self.relay.task


class ConfigCapabilitiesTests(unittest.TestCase):
    def test_register_and_channel_status_advertise_config_relay(self):
        self.assertEqual(build_register_payload(True, 1, {}, "synthetic", "TEST", 115200, "Manual")["config_relay"], 1)
        self.assertEqual(build_channel_status_payload(1, {}, True, True, 2)["config_relay"], 1)

    def test_validate_does_not_strip_or_change_explicit_empty_credentials(self):
        data = validate_config_request(request("set", configs={"DEVICE_SECRET": "", "CALL_TTS_TEXT": "  value  "}))
        self.assertEqual(data["configs"], {"DEVICE_SECRET": "", "CALL_TTS_TEXT": "  value  "})

    def test_saved_console_target_is_distinct_from_actual_started_connection(self):
        data = safe_config_reply(dict(ok=True, saved=True, applied=False, console_connection_requested=True,
            console_connection_enabled=False, restart_required=True, values={"CONSOLE_CONNECTION_ENABLE": True}))
        self.assertTrue(data["values"]["CONSOLE_CONNECTION_ENABLE"])
        self.assertTrue(data["console_connection_requested"])
        self.assertFalse(data["console_connection_enabled"])
        self.assertFalse(data["applied"])

    def test_restart_acceptance_does_not_claim_already_applied(self):
        data = safe_config_reply(dict(ok=True, phase="accepted", accepted=True, applied=False))
        self.assertTrue(data["accepted"])
        self.assertEqual(data["phase"], "accepted")
        self.assertFalse(data["applied"])

    def test_device_support_and_effective_values_preserve_only_safe_metadata(self):
        supported = ["PIN_CODE", "ALLOW_NVM_SERVER_CONFIG", "CELLULAR_DATA_FORCE_DETACH"]
        data = safe_config_reply(dict(ok=True, supported_fields=supported,
            file_only=["CELLULAR_DATA_ENABLE"],
            effective_values={"ALLOW_NVM_SERVER_CONFIG": False, "LED_ENABLE": True,
                "PIN_CODE": "synthetic-private", "DEVICE_SECRET": "synthetic-private",
                "WEBSOCKET_URL": "synthetic-private", "CUSTOM_POST_BODY_TABLE": {"value": "synthetic-private"},
                "CALL_MAX_TIME": float("nan"), "bad_key": "synthetic-private"}))
        self.assertEqual(data["supported_fields"], supported)
        self.assertEqual(data["file_only"], ["CELLULAR_DATA_ENABLE"])
        self.assertEqual(data["effective_values"], {"ALLOW_NVM_SERVER_CONFIG": False, "LED_ENABLE": True})
        self.assertNotIn("synthetic-private", json.dumps(data))

    def test_empty_lua_support_lists_are_normalized_and_bad_metadata_is_rejected(self):
        for key in ("supported_fields", "file_only"):
            for value in ({}, []):
                with self.subTest(key=key, empty=type(value).__name__):
                    self.assertEqual(safe_config_reply(dict(ok=True, **{key: value}))[key], [])
            for value in (None, False, "", {"1": "PIN_CODE"}, [False], ["bad_key"], ["A" * 81], ["PIN_CODE"] * 129):
                with self.subTest(key=key, invalid=type(value).__name__):
                    with self.assertRaisesRegex(RelayError, "config_invalid"):
                        safe_config_reply(dict(ok=True, **{key: value}))

    def test_server_batch_capability_rejects_non_boolean_device_values(self):
        for value in (None, 0, 1, "true", "false", [], {}, [True]):
            with self.subTest(kind=type(value).__name__):
                with self.assertRaisesRegex(RelayError, "config_invalid"):
                    safe_config_reply(dict(ok=True, server_config_batch=value))

    def test_old_device_reply_without_metadata_keeps_its_original_shape(self):
        self.assertEqual(safe_config_reply(dict(ok=True, protocol=1, revision=2,
            values={"CONSOLE_CONNECTION_ENABLE": False}, configured={"DEVICE_SECRET": True})),
            dict(ok=True, protocol=1, revision=2, values={"CONSOLE_CONNECTION_ENABLE": False},
                 configured={"DEVICE_SECRET": True}))


if __name__ == "__main__":
    unittest.main()
