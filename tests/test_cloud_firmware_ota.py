import asyncio
import inspect
import json
import threading
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from sms_core.cloud_command_context import capture_cloud_serial_command_context
from sms_core.cloud_firmware_ota import CloudFirmwareOtaRelay, RelayError, UsbOtaTransport, usb_at_candidates
from sms_core.cloud_message_runtime import handle_cloud_message_runtime
from sms_core import serial_sender

IMEI = "000000000000001"
CAP = dict(protocol=1, supported=True, core="synthetic-core", capacity=458752,
           version="1.0.9", config_md5="a" * 32, main_md5="b" * 32, last_job="")


class Modem:
    port, is_open = "COM1", True

    def __init__(self):
        self.writes = []

    def write(self, data):
        self.writes.append(data)
        return len(data)

    def flush(self):
        pass


class WebSocket:
    def __init__(self):
        self.replies = []
        self.event = asyncio.Event()

    async def send(self, raw):
        self.replies.append(json.loads(raw))
        self.event.set()

    async def received(self, count=1):
        async def wait():
            while len(self.replies) < count:
                self.event.clear()
                await self.event.wait()
        await asyncio.wait_for(wait(), 2)


def namespace(ws=None):
    ws = ws or object()
    result = dict(serial_lock=threading.RLock(), serial_obj=Modem(), serial_connection_generation=1,
                  cloud_connected=True, cloud_device_authorized=True, cloud_ws_conn=ws,
                  CLOUD_DEVICE_SECRET="synthetic-secret", CLOUD_WS_URL="wss://example.invalid/ws/device",
                  CLOUD_CONTROL_ENABLED=True, imei=IMEI)
    result["_cloud_runtime_imei"] = lambda: result["imei"]
    return result


def port(device, description, vid=0x1782, pid=0x4E00):
    return SimpleNamespace(device=device, description=description, vid=vid, pid=pid)


class Clock:
    def __init__(self):
        self.now = 0

    def __call__(self):
        self.now += 0.005
        return self.now


class Usb:
    def __init__(self, imei=IMEI, handler=None, partial=False):
        self.imei, self.handler, self.partial = imei, handler, partial
        self.incoming, self.outgoing, self.requests = bytearray(), bytearray(), []
        self.closed = False

    def reset_input_buffer(self):
        self.outgoing.clear()

    def write(self, data):
        count = min(len(data), 17) if self.partial else len(data)
        self.incoming.extend(data[:count])
        while b"\n" in self.incoming:
            raw, _, rest = self.incoming.partition(b"\n")
            self.incoming = bytearray(rest)
            request = json.loads(raw)
            self.requests.append(request)
            reply = dict(type="firmware_ota_reply", request_id=request["request_id"],
                         job_id=request.get("job_id"), session=request["session"], imei=self.imei, ok=True)
            if request["action"] == "info":
                reply["ota"] = dict(CAP)
            if self.handler:
                reply = self.handler(request, reply)
            if reply is not None:
                self.outgoing.extend(json.dumps(reply).encode() + b"\n")
        return count

    def read(self, size):
        size = min(size, 31)  # Exercise fragmented replies.
        result = bytes(self.outgoing[:size])
        del self.outgoing[:size]
        return result

    def close(self):
        self.closed = True


class UsbOtaTransportTests(unittest.TestCase):
    def setUp(self):
        self.ns = namespace()
        self.context = capture_cloud_serial_command_context(self.ns)
        self.opened, self.transports = [], []

    def tearDown(self):
        for transport in self.transports:
            transport.commit_sent = False
            transport.close()

    def make(self, devices=None):
        devices = devices or {"COM2": Usb()}
        descriptors = [port("COM1", "LUAT USB Device 0 Modem")]
        descriptors += [port(name, "LUAT USB Device 1 AT") for name in devices]
        def open_port(name, **kwargs):
            self.assertEqual(kwargs["timeout"], 0.1)
            self.assertEqual(kwargs["write_timeout"], 2)
            self.opened.append(name)
            return devices[name]
        def current():
            with self.ns["serial_lock"]:
                return not self.context.rejection_reason(self.ns)
        transport = UsbOtaTransport(self.context, current, serial_factory=open_port,
                                    list_ports=lambda: descriptors, clock=Clock(), response_timeout=1)
        self.transports.append(transport)
        return transport

    def request(self, action, **fields):
        return dict(type="firmware_ota", action=action, request_id="1" * 32, job_id="2" * 32, **fields)

    def test_only_known_usb_at_interfaces_are_candidates(self):
        ports = [port("COM1", "LUAT USB Device 0 Modem"), port("COM2", "LUAT USB Device 1 AT"),
                 port("COM3", "LUAT USB Device 2 AP"), port("COM4", "Other USB Serial"),
                 port("COM5", "LUAT USB Device 1 AT", vid=1)]
        self.assertEqual(usb_at_candidates(ports, "COM1"), ["COM2"])
        self.assertEqual(usb_at_candidates(ports, "COM3"), [])

    def test_wrong_device_is_closed_without_ever_receiving_a_secret(self):
        wrong, correct = Usb("000000000000002"), Usb(partial=True)
        transport = self.make({"COM2": wrong, "COM3": correct})
        reply = transport.exchange(self.request("begin", secret="synthetic-secret"))
        self.assertTrue(reply["ok"])
        self.assertTrue(wrong.closed)
        self.assertEqual(self.opened, ["COM2", "COM3"])
        self.assertTrue(all(r["action"] == "info" and "secret" not in r for r in wrong.requests))
        self.assertEqual(correct.requests[-1]["imei"], IMEI)
        self.assertEqual(self.ns["serial_obj"].writes, [])

    def test_nonce_or_session_mismatch_cannot_select_device(self):
        for field in ("request_id", "session"):
            usb = Usb(handler=lambda req, reply: dict(reply, **{field: "0" * 32}))
            transport = self.make({"COM2": usb})
            with self.assertRaisesRegex(RelayError, "usb_unavailable"):
                transport.connect()
            self.assertTrue(usb.closed)
            self.assertTrue(all(r["action"] == "info" for r in usb.requests))

    def test_bounded_receive_rejects_unterminated_frame(self):
        usb = Usb()
        def oversized(req, reply):
            usb.outgoing.extend(b"X" * 7000)
            return None
        usb.handler = oversized
        transport = self.make({"COM2": usb})
        with self.assertRaises(RelayError):
            transport.connect()
        self.assertTrue(usb.closed)
        self.assertEqual(len(transport.buffer), 0)

    def test_changed_serial_or_cloud_context_prevents_new_write(self):
        for key, value in (("serial_obj", Modem()), ("serial_connection_generation", 2),
                           ("imei", "000000000000002"), ("cloud_ws_conn", object()),
                           ("cloud_device_authorized", False), ("CLOUD_DEVICE_SECRET", "changed"),
                           ("CLOUD_WS_URL", "wss://other.invalid/ws/device")):
            with self.subTest(key=key):
                usb = Usb()
                transport = self.make({"COM2": usb})
                transport.connect()
                count = len(usb.requests)
                before = self.ns[key]
                self.ns[key] = value
                with self.assertRaisesRegex(RelayError, "connection_changed"):
                    transport.exchange(self.request("begin"))
                self.assertEqual(len(usb.requests), count)
                self.ns[key] = before

    def test_active_transfer_does_not_reopen_or_follow_a_different_job(self):
        transport = self.make()
        transport.exchange(self.request("begin"))
        with self.assertRaisesRegex(RelayError, "busy"):
            transport.exchange(dict(self.request("chunk"), job_id="3" * 32))
        self.assertEqual(self.opened, ["COM2"])

    def test_ota_reservation_rejects_at_and_sms_bytes_and_releases_on_abort(self):
        transport = self.make()
        transport.exchange(self.request("begin"))
        for payload in (b"ATD10086;\r\n", b"AT+CMGS=20\r\n", b"001122\x1a"):
            result = serial_sender._write_serial_obj_bytes(self.context.serial, payload)
            self.assertFalse(result.ok)
        self.assertEqual(self.context.serial.writes, [])
        transport.exchange(self.request("abort"))
        transport.close()
        self.assertTrue(serial_sender.write_serial_command_result(self.context.serial, "ATI").ok)

    def test_existing_modem_transaction_blocks_new_ota_reservation(self):
        results = []
        with serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK:
            thread = threading.Thread(target=lambda: results.append(serial_sender.reserve_serial_for_ota(self.context.serial, self)))
            thread.start()
            thread.join(2)
        self.assertEqual(results, [False])

    def test_commit_timeout_keeps_local_commands_blocked_until_lease_expires(self):
        usb = Usb(handler=lambda req, reply: None if req["action"] == "commit" else reply)
        transport = self.make({"COM2": usb})
        with patch.object(serial_sender, "time", SimpleNamespace(monotonic=lambda: 100)):
            transport.exchange(self.request("begin"))
            with self.assertRaisesRegex(RelayError, "usb_timeout"):
                transport.exchange(self.request("commit"))
            transport.close()
            self.assertFalse(serial_sender.write_serial_command_result(self.context.serial, "ATI").ok)
        with patch.object(serial_sender, "time", SimpleNamespace(monotonic=lambda: 251)):
            self.assertTrue(serial_sender.write_serial_command_result(self.context.serial, "ATI").ok)


class ControlledTransport:
    def __init__(self, context, current):
        self.context, self.current = context, current
        self.stopped, self.entered, self.release = threading.Event(), threading.Event(), threading.Event()
        self.release.set()
        self.job_id, self.commit_sent, self.closed = None, False, False
        self.actions, self.fail_commit, self.native_failure = [], False, False

    def exchange(self, request):
        self.entered.set()
        while not self.release.wait(0.005):
            if self.stopped.is_set() or not self.current():
                raise RelayError("connection_changed")
        if not self.current():
            raise RelayError("connection_changed")
        self.actions.append(request["action"])
        if request["action"] == "begin":
            self.job_id = request["job_id"]
        if request["action"] == "commit":
            self.commit_sent = True
            if self.fail_commit:
                raise RelayError("usb_timeout")
        return dict(ok=True, ota=CAP, phase="accepted")

    def wait_install_result(self, request):
        if self.native_failure:
            self.commit_sent = False
            return dict(ok=False, phase="failed", reason="native_start")
        raise RelayError("usb_timeout")

    def abort(self):
        pass

    def close(self):
        self.closed = True


class CloudFirmwareOtaRelayTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.ws = WebSocket()
        self.ns = namespace(self.ws)
        self.created = []
        self.setup_transport = lambda _: None
        def factory(context, current):
            transport = ControlledTransport(context, current)
            self.setup_transport(transport)
            self.created.append(transport)
            return transport
        self.relay = CloudFirmwareOtaRelay(self.ns, self.ws, transport_factory=factory)

    async def asyncTearDown(self):
        await self.relay.stop()

    def request(self, action="info", **values):
        fields = dict(type="firmware_ota", transport="desktop", imei=IMEI, serial_connection_generation=1,
                      action=action, request_id="1" * 32, job_id="2" * 32, secret="synthetic-secret")
        fields.update(values)
        return fields

    async def test_invalid_authorization_identity_generation_or_secret_never_opens_usb(self):
        for changes in ({"imei": "000000000000002"}, {"transport": "firmware"},
                        {"serial_connection_generation": 2}, {"serial_connection_generation": True},
                        {"request_id": "invalid"}, {"secret": "wrong"}):
            await self.relay.submit(self.request("begin", **changes))
        self.assertEqual(self.created, [])
        self.assertEqual(self.ws.replies[-1]["reason"], "auth")
        self.ns["cloud_device_authorized"] = False
        await self.relay.submit(self.request())
        self.assertEqual(self.created, [])

    async def test_background_exchange_does_not_block_admission_or_duplicate_work(self):
        self.setup_transport = lambda t: t.release.clear()
        request = self.request("begin")
        await asyncio.wait_for(self.relay.submit(request), 0.5)
        for _ in range(100):
            if self.created:
                break
            await asyncio.sleep(0.001)
        transport = self.created[0]
        self.assertTrue(await asyncio.to_thread(transport.entered.wait, 1))
        await self.relay.submit(request)
        self.assertEqual(len(self.relay.pending), 1)
        transport.release.set()
        await self.ws.received()
        self.assertEqual(transport.actions, ["begin"])

    async def test_cloud_stop_request_rejects_ota_before_socket_cleanup(self):
        self.ns["cloud_stop_event"] = threading.Event()
        self.ns["cloud_stop_event"].set()
        await self.relay.submit(self.request("begin"))
        self.assertTrue(self.ns["cloud_connected"])
        self.assertEqual(self.created, [])
        self.assertEqual(self.ws.replies, [])

    async def test_changed_context_while_worker_waits_cancels_without_writing(self):
        self.setup_transport = lambda t: t.release.clear()
        await self.relay.submit(self.request("begin"))
        for _ in range(100):
            if self.created:
                break
            await asyncio.sleep(0.001)
        transport = self.created[0]
        self.assertTrue(await asyncio.to_thread(transport.entered.wait, 1))
        self.ns["serial_connection_generation"] = 2
        transport.release.set()
        await self.relay.task
        self.assertEqual(transport.actions, [])
        self.assertEqual(self.ws.replies, [])
        self.assertTrue(transport.closed)

    async def test_commit_timeout_does_not_send_a_definite_failure(self):
        self.setup_transport = lambda t: setattr(t, "fail_commit", True)
        await self.relay.submit(self.request("commit"))
        await self.relay.task
        self.assertTrue(self.created[0].commit_sent)
        self.assertEqual(self.ws.replies, [])

    async def test_late_native_failure_is_forwarded_after_commit_acceptance(self):
        self.setup_transport = lambda t: setattr(t, "native_failure", True)
        await self.relay.submit(self.request("commit"))
        await self.relay.task
        self.assertEqual([r["ok"] for r in self.ws.replies], [True, False])
        self.assertEqual(self.ws.replies[-1]["phase"], "failed")
        self.assertTrue(self.created[0].closed)

    async def test_shutdown_joins_serial_worker_and_closes_handle(self):
        self.setup_transport = lambda t: t.release.clear()
        await self.relay.submit(self.request())
        for _ in range(100):
            if self.created:
                break
            await asyncio.sleep(0.001)
        transport = self.created[0]
        self.assertTrue(await asyncio.to_thread(transport.entered.wait, 1))
        await asyncio.wait_for(self.relay.stop(), 2)
        self.assertTrue(self.relay.task.done())
        self.assertTrue(transport.closed)
        self.assertFalse(self.relay.pending)

    async def test_ota_frames_bypass_ordinary_cloud_logging(self):
        logged, handled = [], []
        async def handler(data):
            handled.append(data["type"])
        kwargs = {name: lambda *a, **k: None for name, item in inspect.signature(handle_cloud_message_runtime).parameters.items()
                  if name != "message" and item.default is inspect.Parameter.empty}
        kwargs.update(log=lambda *a, **k: logged.append(a), handle_firmware_ota_message=handler)
        await handle_cloud_message_runtime(json.dumps(self.request("begin", data="synthetic-private-package")), **kwargs)
        self.assertEqual(logged, [])
        self.assertEqual(handled, ["firmware_ota"])


if __name__ == "__main__":
    unittest.main()
