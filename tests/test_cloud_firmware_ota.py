import asyncio
import inspect
import json
import threading
import time
import unittest
from contextlib import contextmanager, suppress
from types import SimpleNamespace
from unittest.mock import patch

from sms_core.cloud_command_context import capture_cloud_serial_command_context, set_cloud_command_authorized
from sms_core.cloud_firmware_ota import CloudFirmwareOtaRelay, RelayError, UsbOtaTransport, usb_at_candidates, handle_firmware_ota
from sms_core.cloud_message_runtime import handle_cloud_message_runtime
from sms_core.cloud_message_namespace_runtime import handle_cloud_message_namespace_runtime
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


class ObservedTransactionLock:
    def __init__(self, lock):
        self.lock = lock
        self.waiting = threading.Event()
        self.timeouts = []

    def acquire(self, blocking=True, timeout=-1):
        if timeout > 0:
            self.timeouts.append(timeout)
            self.waiting.set()
        return self.lock.acquire(blocking, timeout)

    def release(self):
        self.lock.release()

    def __enter__(self):
        self.lock.acquire()
        return self

    def __exit__(self, *_args):
        self.lock.release()


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

    @contextmanager
    def pending_ordinary_at(self):
        coordinator = serial_sender.AtCommandResponseCoordinator()
        written, results = threading.Event(), []
        modem = self.context.serial
        original_write = modem.write

        def write(data):
            result = original_write(data)
            written.set()
            return result

        def execute():
            results.append(serial_sender.write_serial_command_confirmed_locked(
                self.ns["serial_lock"], lambda: modem, "ATI",
                response_coordinator=coordinator, response_timeout=5,
            ))

        with patch.object(modem, "write", write):
            worker = threading.Thread(target=execute)
            worker.start()
            try:
                self.assertTrue(written.wait(1))
                yield coordinator, results
            finally:
                coordinator.observe_line("OK", connection=modem)
                worker.join(2)
                self.assertFalse(worker.is_alive())

    def start_exchange(self, transport, request):
        replies, errors = [], []

        def exchange():
            try:
                replies.append(transport.exchange(request))
            except Exception as exc:
                errors.append(exc)

        worker = threading.Thread(target=exchange)
        worker.start()
        return worker, replies, errors

    def test_only_known_usb_at_interfaces_are_candidates(self):
        ports = [port("COM1", "LUAT USB Device 0 Modem"), port("COM2", "LUAT USB Device 1 AT"),
                 port("COM3", "LUAT USB Device 2 AP"), port("COM4", "Other USB Serial"),
                 port("COM5", "LUAT USB Device 1 AT", vid=1)]
        self.assertEqual(usb_at_candidates(ports, "COM1"), ["COM2"])
        self.assertEqual(usb_at_candidates(ports, "COM3"), [])

    def test_first_info_uses_identity_probe_but_next_info_reads_fresh_version(self):
        usb = Usb()
        transport = self.make({"COM2": usb})
        first = transport.exchange(self.request("info"))
        self.assertEqual(first["ota"]["version"], CAP["version"])
        self.assertEqual(len(usb.requests), 1)
        usb.handler = lambda req, reply: dict(reply, ota=dict(CAP, version="1.0.10"))
        second = transport.exchange(self.request("info"))
        self.assertEqual(second["ota"]["version"], "1.0.10")
        self.assertEqual(len(usb.requests), 2)
        self.assertNotEqual(usb.requests[0]["request_id"], usb.requests[1]["request_id"])

    def test_info_accepts_delayed_usb_reply_within_response_timeout(self):
        usb = Usb()
        transport = self.make({"COM2": usb})
        transport.response_timeout = 5
        original_read = usb.read
        def delayed_read(size):
            return b"" if transport.clock.now < 1.5 else original_read(size)
        usb.read = delayed_read
        self.assertTrue(transport.exchange(self.request("info"))["ok"])
        self.assertGreater(transport.clock.now, 1.5)

    def test_first_lost_info_is_resynchronized_without_reopening_usb(self):
        usb = Usb(handler=lambda req, reply: None if len(usb.requests) == 1 else reply)
        transport = self.make({"COM2": usb})
        transport.response_timeout = 5
        self.assertTrue(transport.exchange(self.request("info"))["ok"])
        self.assertEqual(self.opened, ["COM2"])
        self.assertEqual(len(usb.requests), 2)
        self.assertEqual(usb.requests[0], usb.requests[1])
        self.assertTrue(all(req["action"] == "info" and "secret" not in req for req in usb.requests))
        self.assertEqual(self.context.serial.writes, [])
        self.assertLess(transport.clock.now, 5.1)

    def test_resynchronization_accepts_delayed_first_reply_with_same_nonce(self):
        usb = Usb(handler=lambda req, reply: reply if len(usb.requests) == 1 else None)
        transport = self.make({"COM2": usb})
        transport.response_timeout = 5
        original_read = usb.read
        usb.read = lambda size: b"" if transport.clock.now < 1.5 else original_read(size)
        self.assertTrue(transport.exchange(self.request("info"))["ok"])
        self.assertEqual(len(usb.requests), 2)
        self.assertEqual(usb.requests[0]["request_id"], usb.requests[1]["request_id"])
        self.assertLess(transport.clock.now, 5.1)

    def test_existing_usb_info_timeout_does_not_resynchronize_again(self):
        usb = Usb()
        transport = self.make({"COM2": usb})
        transport.response_timeout = 5
        transport.exchange(self.request("info"))
        usb.handler = lambda req, reply: None
        with self.assertRaisesRegex(RelayError, "usb_timeout"):
            transport.exchange(self.request("info"))
        self.assertEqual(len(usb.requests), 2)
        self.assertEqual(self.opened, ["COM2"])

    def test_changed_context_before_resynchronization_prevents_second_write(self):
        usb = Usb(handler=lambda req, reply: None)
        transport = self.make({"COM2": usb})
        transport.response_timeout = 5
        original_read = transport._read

        def read_then_revoke(*args, **kwargs):
            try:
                return original_read(*args, **kwargs)
            except RelayError:
                self.ns["serial_connection_generation"] += 1
                raise

        transport._read = read_then_revoke
        with self.assertRaisesRegex(RelayError, "connection_changed"):
            transport.connect()
        self.assertEqual(len(usb.requests), 1)
        self.assertTrue(usb.closed)

    def test_unsupported_reply_does_not_trigger_resynchronization(self):
        usb = Usb(handler=lambda req, reply: dict(reply, ok=False))
        transport = self.make({"COM2": usb})
        transport.response_timeout = 5
        with self.assertRaisesRegex(RelayError, "usb_unavailable"):
            transport.connect()
        self.assertEqual(len(usb.requests), 1)
        self.assertTrue(usb.closed)

    def test_resynchronization_short_writes_stop_within_original_budget(self):
        usb = Usb(partial=True, handler=lambda req, reply: None)
        transport = self.make({"COM2": usb})
        transport.response_timeout = 5
        original_write = usb.write

        def slow_retry_write(data):
            if usb.requests:
                transport.clock.now += 1
            return original_write(data)

        usb.write = slow_retry_write
        with self.assertRaisesRegex(RelayError, "usb_unavailable"):
            transport.connect()
        self.assertEqual(len(usb.requests), 1)
        self.assertTrue(usb.closed)
        self.assertLess(transport.clock.now, 5.1)

    def test_first_probe_slow_short_writes_stop_within_scan_budget(self):
        usb = Usb(partial=True)
        transport = self.make({"COM2": usb})
        transport.response_timeout = 5
        original_write = usb.write

        def slow_write(data):
            transport.clock.now += 1.5
            return original_write(data)

        usb.write = slow_write
        with self.assertRaisesRegex(RelayError, "usb_unavailable"):
            transport.connect()
        self.assertLess(transport.clock.now, 8.1)
        self.assertEqual(usb.requests, [])
        self.assertTrue(usb.closed)
        self.assertEqual(self.opened, ["COM2"])

    def test_first_probe_short_writes_preserve_remaining_reply_budget(self):
        usb = Usb(partial=True)
        transport = self.make({"COM2": usb})
        transport.response_timeout = 5
        original_write = usb.write

        def slower_write(data):
            transport.clock.now += 0.6
            return original_write(data)

        usb.write = slower_write
        reply = transport.exchange(self.request("info"))
        self.assertTrue(reply["ok"])
        self.assertEqual(reply["ota"], CAP)
        self.assertGreater(transport.clock.now, 5)
        self.assertLess(transport.clock.now, 8.1)
        self.assertEqual([request["action"] for request in usb.requests], ["info"])
        self.assertEqual(self.opened, ["COM2"])

    def test_late_candidate_reply_keeps_remaining_time_without_extra_write(self):
        unresponsive, correct = Usb(handler=lambda req, reply: None), Usb()
        transport = self.make({"COM2": unresponsive, "COM3": correct})
        transport.response_timeout = 5
        original_read = correct.read
        correct.read = lambda size: b"" if transport.clock.now < 6.6 else original_read(size)
        self.assertTrue(transport.exchange(self.request("info"))["ok"])
        self.assertEqual(self.opened, ["COM2", "COM3"])
        self.assertEqual(len(correct.requests), 1)
        self.assertLess(transport.clock.now, 8.1)

    def test_unresponsive_candidates_have_one_bounded_scan_budget(self):
        devices = {f"COM{n}": Usb(handler=lambda req, reply: None) for n in range(2, 10)}
        transport = self.make(devices)
        transport.response_timeout = 5
        with self.assertRaisesRegex(RelayError, "usb_unavailable"):
            transport.exchange(self.request("info"))
        self.assertLess(transport.clock.now, 10)
        self.assertTrue(all(devices[name].closed for name in self.opened))

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

    def test_first_begin_waits_for_ordinary_at_ok_and_submits_once(self):
        usb = Usb()
        transport = self.make({"COM2": usb})
        transport.connect()
        observed = ObservedTransactionLock(serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK)
        with patch.object(serial_sender, "DEFAULT_SERIAL_TRANSACTION_LOCK", observed):
            with self.pending_ordinary_at() as (coordinator, at_results):
                worker, replies, errors = self.start_exchange(transport, self.request("begin"))
                try:
                    self.assertTrue(observed.waiting.wait(1), "OTA must wait for the existing AT transaction")
                    coordinator.observe_line("OK", connection=self.context.serial)
                finally:
                    coordinator.observe_line("OK", connection=self.context.serial)
                    worker.join(3)
                    self.assertFalse(worker.is_alive())
            self.assertEqual(errors, [])
            self.assertTrue(replies[0]["ok"])
            self.assertTrue(at_results[0].ok)
        self.assertEqual(observed.timeouts, [2])
        self.assertEqual(self.context.serial.writes, [b"ATI\r\n"])
        self.assertEqual([req["action"] for req in usb.requests], ["info", "begin"])

    def test_first_begin_continuing_transaction_times_out_without_usb_submission(self):
        usb = Usb()
        transport = self.make({"COM2": usb})
        transport.connect()
        observed = ObservedTransactionLock(serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK)
        with patch.object(serial_sender, "DEFAULT_SERIAL_TRANSACTION_LOCK", observed), observed:
            started = time.monotonic()
            worker, replies, errors = self.start_exchange(transport, self.request("begin"))
            worker.join(3)
            self.assertFalse(worker.is_alive())
            elapsed = time.monotonic() - started
        self.assertEqual(replies, [])
        self.assertEqual(len(errors), 1)
        self.assertIsInstance(errors[0], RelayError)
        self.assertEqual(str(errors[0]), "serial_busy")
        self.assertEqual(observed.timeouts, [2])
        self.assertGreaterEqual(elapsed, 1.7)
        self.assertLess(elapsed, 3)
        self.assertEqual([req["action"] for req in usb.requests], ["info"])
        self.assertNotIn(id(self.context.serial), serial_sender._OTA_RESERVATIONS)

    def test_first_begin_rechecks_revocation_generation_and_stop_after_wait(self):
        for change in ("authorization", "authorization_restored", "generation", "stop"):
            with self.subTest(change=change):
                usb = Usb()
                transport = self.make({"COM2": usb})
                transport.connect()
                observed = ObservedTransactionLock(serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK)
                with patch.object(serial_sender, "DEFAULT_SERIAL_TRANSACTION_LOCK", observed):
                    with self.pending_ordinary_at() as (coordinator, _results):
                        worker, replies, errors = self.start_exchange(transport, self.request("begin"))
                        try:
                            self.assertTrue(observed.waiting.wait(1))
                            if change == "authorization":
                                self.ns["cloud_device_authorized"] = False
                            elif change == "authorization_restored":
                                set_cloud_command_authorized(self.ns, False)
                                set_cloud_command_authorized(self.ns, True)
                            elif change == "generation":
                                self.ns["serial_connection_generation"] += 1
                            else:
                                transport.stopped.set()
                        finally:
                            coordinator.observe_line("OK", connection=self.context.serial)
                            worker.join(3)
                            self.assertFalse(worker.is_alive())
                self.assertEqual(replies, [])
                self.assertEqual(len(errors), 1)
                self.assertIsInstance(errors[0], RelayError)
                self.assertEqual(str(errors[0]), "connection_changed")
                self.assertIsNone(transport.job_id)
                self.assertFalse(transport.commit_sent)
                self.assertNotIn(id(self.context.serial), serial_sender._OTA_RESERVATIONS)
                self.assertEqual([req["action"] for req in usb.requests], ["info"])
                self.ns["cloud_device_authorized"] = True
                self.ns["serial_connection_generation"] = self.context.generation
                self.ns["_cloud_authorization_epoch"] = self.context.authorization_epoch

    def test_context_change_after_reservation_releases_lease_before_begin_write(self):
        usb = Usb()
        transport = self.make({"COM2": usb})
        transport.connect()
        reserve = serial_sender.reserve_serial_for_ota

        def reserve_then_change(*args, **kwargs):
            result = reserve(*args, **kwargs)
            self.ns["serial_connection_generation"] += 1
            return result

        with patch("sms_core.cloud_firmware_ota.reserve_serial_for_ota", reserve_then_change):
            with self.assertRaisesRegex(RelayError, "connection_changed"):
                transport.exchange(self.request("begin"))
        self.assertNotIn(id(self.context.serial), serial_sender._OTA_RESERVATIONS)
        self.assertIsNone(transport.job_id)
        self.assertFalse(transport.commit_sent)
        self.assertEqual([req["action"] for req in usb.requests], ["info"])

    def test_failed_reservation_rechecks_context_before_reporting_serial_busy(self):
        usb = Usb()
        transport = self.make({"COM2": usb})
        transport.connect()

        def timeout_then_revoke(*_args, **_kwargs):
            self.ns["cloud_device_authorized"] = False
            return False

        with patch("sms_core.cloud_firmware_ota.reserve_serial_for_ota", timeout_then_revoke):
            with self.assertRaisesRegex(RelayError, "connection_changed"):
                transport.exchange(self.request("begin"))
        self.assertNotIn(id(self.context.serial), serial_sender._OTA_RESERVATIONS)
        self.assertEqual([req["action"] for req in usb.requests], ["info"])

    def test_existing_lease_renewal_and_later_actions_do_not_wait_for_transactions(self):
        usb = Usb()
        transport = self.make({"COM2": usb})
        reserve = serial_sender.reserve_serial_for_ota
        with patch("sms_core.cloud_firmware_ota.reserve_serial_for_ota", wraps=reserve) as reservation:
            transport.exchange(self.request("begin"))
            with serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK:
                replies, errors = [], []

                def exchange_remaining():
                    try:
                        for action in ("begin", "chunk", "verify", "commit", "abort"):
                            replies.append(transport.exchange(self.request(action)))
                    except Exception as exc:
                        errors.append(exc)

                worker = threading.Thread(target=exchange_remaining)
                worker.start()
                worker.join(1)
                self.assertFalse(worker.is_alive())
            self.assertEqual(errors, [])
            self.assertTrue(all(reply["ok"] for reply in replies))
            self.assertEqual(len(replies), 5)
            self.assertEqual([call.kwargs.get("wait_timeout", 0) for call in reservation.call_args_list], [2, 0, 0, 0, 0, 0])
        self.assertFalse(transport.commit_sent)
        self.assertIs(serial_sender._OTA_RESERVATIONS[id(self.context.serial)][1], transport)

    def test_competing_ota_owner_remains_reserved_without_new_begin_submission(self):
        usb = Usb()
        transport = self.make({"COM2": usb})
        transport.connect()
        other_owner = object()
        self.assertTrue(serial_sender.reserve_serial_for_ota(self.context.serial, other_owner))
        try:
            with self.assertRaisesRegex(RelayError, "serial_busy"):
                transport.exchange(self.request("begin"))
            self.assertIs(serial_sender._OTA_RESERVATIONS[id(self.context.serial)][1], other_owner)
            self.assertEqual([req["action"] for req in usb.requests], ["info"])
        finally:
            serial_sender.release_serial_from_ota(self.context.serial, other_owner)

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

    def test_invalid_commit_receipts_keep_lease_and_never_abort(self):
        for fields in ({"ok": False, "phase": "accepted"}, {"ok": True, "phase": "failed"},
                       {"phase": "failed"}, {"ok": 0, "phase": "failed"},
                       {"ok": "false", "phase": "failed"}):
            with self.subTest(fields=fields):
                self.ns = namespace()
                self.context = capture_cloud_serial_command_context(self.ns)
                def handler(req, reply):
                    if req["action"] == "commit":
                        reply.pop("ok", None)
                        reply.update(fields)
                    return reply
                usb = Usb(handler=handler)
                transport = self.make({"COM2": usb})
                transport.exchange(self.request("begin"))
                with self.assertRaisesRegex(RelayError, "commit_reply"):
                    transport.exchange(dict(self.request("commit"), request_id="3" * 32))
                self.assertTrue(transport.commit_sent)
                transport.abort()
                transport.close()
                self.assertFalse(any(req["action"] == "abort" for req in usb.requests))
                self.assertFalse(serial_sender.write_serial_command_result(self.context.serial, "ATI").ok)
                self.assertIs(serial_sender._OTA_RESERVATIONS[id(self.context.serial)][1], transport)
                transport.commit_sent = False
                transport.close()

    def test_consistent_commit_rejections_release_lease(self):
        for fields in ({"ok": False, "phase": "failed"}, {"ok": False, "reason": "busy"}):
            with self.subTest(fields=fields):
                self.ns = namespace()
                self.context = capture_cloud_serial_command_context(self.ns)
                usb = Usb(handler=lambda req, reply: dict(reply, **fields) if req["action"] == "commit" else reply)
                transport = self.make({"COM2": usb})
                transport.exchange(self.request("begin"))
                self.assertFalse(transport.exchange(dict(self.request("commit"), request_id="3" * 32))["ok"])
                self.assertFalse(transport.commit_sent)
                transport.close()
                self.assertTrue(serial_sender.write_serial_command_result(self.context.serial, "ATI").ok)

    def test_invalid_late_install_receipts_cannot_release_lease(self):
        for fields in ({"ok": True, "phase": "failed"}, {"ok": False, "phase": "rebooted"},
                       {"phase": "failed"}, {"ok": 0, "phase": "failed"}):
            with self.subTest(fields=fields):
                self.ns = namespace()
                self.context = capture_cloud_serial_command_context(self.ns)
                usb = Usb()
                transport = self.make({"COM2": usb})
                transport.exchange(self.request("begin"))
                request = dict(self.request("commit"), request_id="3" * 32)
                transport.exchange(request)
                reply = dict(type="firmware_ota_reply", request_id=request["request_id"],
                             job_id=request["job_id"], session=transport.session, imei=IMEI, **fields)
                usb.outgoing.extend(json.dumps(reply).encode() + b"\n")
                with self.assertRaisesRegex(RelayError, "usb_timeout"):
                    transport.wait_install_result(request)
                self.assertTrue(transport.commit_sent)
                transport.abort()
                transport.close()
                self.assertFalse(any(req["action"] == "abort" for req in usb.requests))
                self.assertIs(serial_sender._OTA_RESERVATIONS[id(self.context.serial)][1], transport)
                transport.commit_sent = False
                transport.close()

    def test_failed_fresh_probe_cannot_prove_installation_or_release_lease(self):
        for fields in ({"ok": False, "phase": "failed"}, {"ok": False}, {"ok": 0}, {"ok": "true"}, {}):
            with self.subTest(fields=fields):
                self.ns = namespace()
                self.context = capture_cloud_serial_command_context(self.ns)
                usb = Usb()
                transport = self.make({"COM2": usb})
                target = dict(version="1.0.10", main_md5="c" * 32)
                transport.exchange(self.request("begin", manifest=target))
                request = dict(self.request("commit"), request_id="3" * 32)
                transport.exchange(request)
                def failed_probe(req, reply):
                    reply.pop("ok", None)
                    return dict(reply, ota=dict(CAP, **target, last_job=transport.job_id), **fields)
                usb.handler = failed_probe
                with self.assertRaisesRegex(RelayError, "usb_timeout"):
                    transport.wait_install_result(request)
                self.assertTrue(transport.commit_sent)
                transport.close()
                self.assertFalse(serial_sender.write_serial_command_result(self.context.serial, "ATI").ok)
                transport.commit_sent = False
                transport.close()

    def test_fresh_probe_and_native_failure_keep_confirmed_outcomes(self):
        for outcome in ("rebooted", "failed"):
            with self.subTest(outcome=outcome):
                self.ns = namespace()
                self.context = capture_cloud_serial_command_context(self.ns)
                usb = Usb()
                transport = self.make({"COM2": usb})
                target = dict(version="1.0.10", main_md5="c" * 32)
                transport.exchange(self.request("begin", manifest=target))
                request = dict(self.request("commit"), request_id="3" * 32)
                transport.exchange(request)
                if outcome == "failed":
                    reply = dict(type="firmware_ota_reply", request_id=request["request_id"],
                                 job_id=request["job_id"], session=transport.session, imei=IMEI,
                                 ok=False, phase="failed", reason="native_start")
                    usb.outgoing.extend(json.dumps(reply).encode() + b"\n")
                else:
                    usb.handler = lambda req, reply: dict(reply, ota=dict(CAP, **target, last_job=transport.job_id))
                result = transport.wait_install_result(request)
                self.assertEqual(result["phase"], outcome)
                self.assertIs(result["ok"], outcome == "rebooted")
                self.assertFalse(transport.commit_sent)
                transport.close()
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

    async def test_malformed_action_returns_safe_failure_and_keeps_connection_for_next_info(self):
        self.ns.update(_firmware_ota_relay=self.relay,
                       _cloud_log=lambda *_a, **_k: None,
                       set_cloud_auth_status_from_ack=lambda *_a: None,
                       set_cloud_status=lambda *_a: None,
                       _cloud_check_replay_window=lambda *_a, **_k: True,
                       _cloud_auth_matches=lambda *_a: True,
                       _cloud_now_ts=lambda: 1,
                       _cloud_send_status_payload=lambda: {},
                       show_window=lambda: None, hide_window=lambda: None)
        async def reply(*_args):
            return None
        self.ns['_cloud_reply'] = reply
        invalid_actions = ([], {}, ['synthetic-private'], {'secret': 'synthetic-private'}, None, 1, False)
        for action in invalid_actions:
            with self.subTest(action_type=type(action).__name__):
                result = await handle_cloud_message_namespace_runtime(
                    self.ns, self.ws, json.dumps(self.request(action)))
                self.assertEqual(result, 'firmware_ota')
                self.assertEqual(self.ws.replies[-1]['reason'], 'invalid_request')
                self.assertFalse(self.ws.replies[-1]['ok'])
                self.assertEqual(self.ws.replies[-1]['action'], '')
                self.assertNotIn('synthetic-private', json.dumps(self.ws.replies[-1]))
                self.assertTrue(self.ns['cloud_connected'])
                self.assertTrue(self.ns['cloud_device_authorized'])
        self.assertEqual(self.created, [])
        self.assertEqual(self.ns['serial_obj'].writes, [])
        await handle_cloud_message_namespace_runtime(self.ns, self.ws, json.dumps(self.request()))
        await self.ws.received(len(invalid_actions) + 1)
        self.assertTrue(self.ws.replies[-1]['ok'])
        self.assertEqual(self.created[0].actions, ['info'])

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

    async def test_read_only_info_closes_usb_before_cloud_reply(self):
        await self.relay.submit(self.request())
        await self.ws.received()
        self.assertTrue(self.created[0].closed)
        self.assertIsNone(self.relay.transport)
        self.assertFalse(self.relay.closed)
        self.assertEqual(self.created[0].actions, ["info"])

    async def test_info_reply_backpressure_does_not_hold_usb(self):
        entered, release = asyncio.Event(), asyncio.Event()
        original_send = self.ws.send

        async def delayed_send(raw):
            entered.set()
            await release.wait()
            await original_send(raw)

        self.ws.send = delayed_send
        try:
            await self.relay.submit(self.request())
            await asyncio.wait_for(entered.wait(), 2)
            self.assertTrue(self.created[0].closed)
            self.assertIsNone(self.relay.transport)
            self.assertEqual(self.ws.replies, [])
        finally:
            release.set()
        await self.ws.received()

    async def test_begin_after_info_uses_new_transport_and_keeps_job(self):
        await self.relay.submit(self.request())
        await self.ws.received()
        await self.relay.submit(self.request("begin", request_id="3" * 32))
        await self.ws.received(2)
        self.assertEqual(len(self.created), 2)
        self.assertTrue(self.created[0].closed)
        self.assertIs(self.relay.transport, self.created[1])
        self.assertEqual(self.created[1].job_id, "2" * 32)
        self.assertFalse(self.created[1].closed)

    async def test_info_during_active_job_keeps_original_transport(self):
        await self.relay.submit(self.request("begin"))
        await self.ws.received()
        await self.relay.submit(self.request(request_id="3" * 32))
        await self.ws.received(2)
        self.assertEqual(len(self.created), 1)
        self.assertIs(self.relay.transport, self.created[0])
        self.assertFalse(self.created[0].closed)
        self.assertEqual(self.created[0].actions, ["begin", "info"])

    async def test_queued_info_probes_close_each_idle_transport(self):
        self.setup_transport = lambda transport: transport.release.clear() if not self.created else None
        await self.relay.submit(self.request())
        for _ in range(100):
            if self.created:
                break
            await asyncio.sleep(0.001)
        self.assertTrue(await asyncio.to_thread(self.created[0].entered.wait, 1))
        await self.relay.submit(self.request(request_id="3" * 32))
        self.created[0].release.set()
        await self.ws.received(2)
        self.assertEqual(len(self.created), 2)
        self.assertTrue(all(transport.closed for transport in self.created))
        self.assertIsNone(self.relay.transport)
        self.assertFalse(self.relay.pending)

    async def test_failed_cloud_reply_does_not_retain_read_only_usb(self):
        async def failed_send(raw):
            raise OSError("synthetic reply failure")

        self.ws.send = failed_send
        await self.relay.submit(self.request())
        await asyncio.wait_for(self.relay.task, 2)
        self.assertTrue(self.created[0].closed)
        self.assertIsNone(self.relay.transport)
        self.assertTrue(self.relay.closed)
        self.assertFalse(self.relay.pending)

    async def test_new_info_during_cleanup_waits_for_old_usb_handle_to_close(self):
        entered, release = threading.Event(), threading.Event()
        cleanup_current = []
        def setup(transport):
            def abort():
                entered.set()
                release.wait(2)
                cleanup_current.append(transport.current())
            transport.abort = abort
        self.setup_transport = setup
        self.ns["_firmware_ota_relay"] = self.relay
        await self.relay.submit(self.request("abort"))
        self.assertTrue(await asyncio.to_thread(entered.wait, 1))
        replacement = CloudFirmwareOtaRelay(self.ns, self.ws, transport_factory=ControlledTransport)
        try:
            with patch("sms_core.cloud_firmware_ota.CloudFirmwareOtaRelay", return_value=replacement):
                pending = asyncio.create_task(handle_firmware_ota(
                    self.ns, self.ws, self.request(request_id="3" * 32)))
                await asyncio.sleep(0.02)
                release.set()
                await pending
                await self.relay.task
                await self.ws.received(2)
            self.assertTrue(self.created[0].closed)
            self.assertEqual(cleanup_current, [True])
            self.assertEqual(self.ws.replies[-1]["request_id"], "3" * 32)
            self.assertTrue(self.ws.replies[-1]["ok"])
        finally:
            release.set()
            await replacement.stop()

    async def test_request_queued_before_abort_gets_failure_instead_of_being_dropped(self):
        self.setup_transport = lambda t: t.release.clear()
        await self.relay.submit(self.request("abort"))
        for _ in range(100):
            if self.created:
                break
            await asyncio.sleep(0.001)
        transport = self.created[0]
        self.assertTrue(await asyncio.to_thread(transport.entered.wait, 1))
        await self.relay.submit(self.request(request_id="3" * 32))
        transport.release.set()
        await self.relay.task
        self.assertEqual(len(self.ws.replies), 2)
        self.assertEqual(self.ws.replies[-1]["request_id"], "3" * 32)
        self.assertFalse(self.ws.replies[-1]["ok"])
        self.assertTrue(transport.closed)
        self.assertFalse(self.relay.pending)

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

    async def test_real_transport_invalid_commit_never_forwards_definite_failure(self):
        usb = Usb(handler=lambda req, reply: dict(reply, ok=False, phase="accepted")
                  if req["action"] == "commit" else reply)
        transports = []
        def factory(context, current):
            transport = UsbOtaTransport(context, current, serial_factory=lambda *_args, **_kwargs: usb,
                list_ports=lambda: [port("COM1", "LUAT USB Device 0 Modem"), port("COM2", "LUAT USB Device 1 AT")],
                clock=Clock(), response_timeout=1)
            transports.append(transport)
            return transport
        self.relay.transport_factory = factory
        try:
            await self.relay.submit(self.request("begin"))
            await self.ws.received()
            await self.relay.submit(self.request("commit", request_id="3" * 32))
            await self.relay.task
            self.assertEqual([reply["action"] for reply in self.ws.replies], ["begin"])
            self.assertTrue(transports[0].commit_sent)
            self.assertTrue(usb.closed)
            self.assertFalse(any(req["action"] == "abort" for req in usb.requests))
            self.assertIs(serial_sender._OTA_RESERVATIONS[id(self.ns["serial_obj"])][1], transports[0])
        finally:
            for transport in transports:
                transport.commit_sent = False
                transport.close()

    async def test_real_transport_invalid_late_reply_only_forwards_acceptance(self):
        usb = Usb()
        original_write = usb.write
        def write(data):
            count = original_write(data)
            request = usb.requests[-1]
            if request["action"] == "commit":
                reply = dict(type="firmware_ota_reply", request_id=request["request_id"],
                             job_id=request["job_id"], session=request["session"], imei=IMEI,
                             ok=True, phase="failed")
                usb.outgoing.extend(json.dumps(reply).encode() + b"\n")
            return count
        usb.write = write
        transports = []
        def factory(context, current):
            transport = UsbOtaTransport(context, current, serial_factory=lambda *_args, **_kwargs: usb,
                list_ports=lambda: [port("COM1", "LUAT USB Device 0 Modem"), port("COM2", "LUAT USB Device 1 AT")],
                clock=Clock(), response_timeout=1)
            transports.append(transport)
            return transport
        self.relay.transport_factory = factory
        try:
            await self.relay.submit(self.request("begin"))
            await self.ws.received()
            await self.relay.submit(self.request("commit", request_id="3" * 32))
            await self.relay.task
            self.assertEqual([reply["action"] for reply in self.ws.replies], ["begin", "commit"])
            self.assertTrue(all(reply["ok"] is True for reply in self.ws.replies))
            self.assertTrue(transports[0].commit_sent)
            self.assertTrue(usb.closed)
            self.assertFalse(any(req["action"] == "abort" for req in usb.requests))
            self.assertIs(serial_sender._OTA_RESERVATIONS[id(self.ns["serial_obj"])][1], transports[0])
        finally:
            for transport in transports:
                transport.commit_sent = False
                transport.close()

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

    async def test_repeated_cancel_joins_blocked_usb_before_close_and_replacement(self):
        entered, release, finished = threading.Event(), threading.Event(), threading.Event()
        usb = Usb()
        original_write, original_close = usb.write, usb.close
        closed_while_writing, transports = [], []

        def write(data):
            entered.set()
            try:
                release.wait(5)
                return original_write(data)
            finally:
                finished.set()

        def close():
            closed_while_writing.append(not finished.is_set())
            original_close()

        def factory(context, current):
            transport = UsbOtaTransport(
                context, current, serial_factory=lambda *_args, **_kwargs: usb,
                list_ports=lambda: [port("COM1", "LUAT USB Device 0 Modem"), port("COM2", "LUAT USB Device 1 AT")],
                clock=Clock(), response_timeout=1,
            )
            transports.append(transport)
            return transport

        usb.write, usb.close = write, close
        self.relay.transport_factory = factory
        self.ns["_firmware_ota_relay"] = self.relay
        replacement = CloudFirmwareOtaRelay(self.ns, self.ws, transport_factory=ControlledTransport)
        stopping = pending = None
        try:
            await self.relay.submit(self.request())
            self.assertTrue(await asyncio.to_thread(entered.wait, 1))
            stopping = asyncio.create_task(self.relay.stop())
            await asyncio.sleep(0)
            self.relay.task.cancel()
            self.assertTrue(await asyncio.to_thread(transports[0].stopped.wait, 1))
            self.relay.task.cancel()
            with patch("sms_core.cloud_firmware_ota.CloudFirmwareOtaRelay", return_value=replacement) as create:
                pending = asyncio.create_task(handle_firmware_ota(
                    self.ns, self.ws, self.request("begin", request_id="3" * 32)))
                await asyncio.sleep(0.02)
                self.assertFalse(finished.is_set())
                self.assertFalse(self.relay.task.done())
                self.assertFalse(usb.closed)
                create.assert_not_called()
                release.set()
                await asyncio.wait_for(stopping, 2)
                with self.assertRaises(asyncio.CancelledError):
                    await self.relay.task
                await asyncio.wait_for(pending, 2)
                await self.ws.received()
                create.assert_called_once()
            self.assertTrue(finished.is_set())
            self.assertTrue(usb.closed)
            self.assertFalse(any(closed_while_writing))
            self.assertTrue(self.ws.replies[-1]["ok"])
            self.assertEqual(self.ws.replies[-1]["request_id"], "3" * 32)
            self.assertTrue(all(request["action"] == "info" for request in usb.requests))
            self.assertNotIn(id(self.ns["serial_obj"]), serial_sender._OTA_RESERVATIONS)
        finally:
            release.set()
            if stopping is not None:
                await stopping
            if pending is not None:
                with suppress(asyncio.CancelledError):
                    await pending
            await replacement.stop()

    async def test_terminal_reply_backpressure_waits_for_cleanup_before_next_begin(self):
        for action, reason in (("begin", "busy"), ("abort", "session"), ("begin", "serial_busy")):
            with self.subTest(action=action, reason=reason):
                ws = WebSocket()
                ns = namespace(ws)
                send_release = asyncio.Event()
                cleanup_entered, cleanup_release = threading.Event(), threading.Event()
                created = []

                def factory(context, current):
                    transport = ControlledTransport(context, current)

                    def exchange(request):
                        transport.actions.append(request["action"])
                        if reason == "serial_busy":
                            raise RelayError(reason)
                        return dict(ok=False, reason=reason)

                    def abort():
                        cleanup_entered.set()
                        cleanup_release.wait(5)

                    transport.exchange, transport.abort = exchange, abort
                    created.append(transport)
                    return transport

                original_send = ws.send

                async def send(raw):
                    await original_send(raw)
                    if json.loads(raw)["request_id"] == "1" * 32:
                        await send_release.wait()

                ws.send = send
                relay = CloudFirmwareOtaRelay(ns, ws, transport_factory=factory)
                ns["_firmware_ota_relay"] = relay
                replacement = CloudFirmwareOtaRelay(ns, ws, transport_factory=ControlledTransport)
                pending = None
                try:
                    await relay.submit(self.request(action))
                    await ws.received()
                    with patch("sms_core.cloud_firmware_ota.CloudFirmwareOtaRelay", return_value=replacement) as create:
                        pending = asyncio.create_task(handle_firmware_ota(
                            ns, ws, self.request("begin", request_id="3" * 32, job_id="4" * 32)))
                        await asyncio.sleep(0.02)
                        self.assertTrue(relay.closing)
                        self.assertFalse(pending.done())
                        create.assert_not_called()
                        send_release.set()
                        self.assertTrue(await asyncio.to_thread(cleanup_entered.wait, 1))
                        self.assertFalse(created[0].closed)
                        self.assertFalse(pending.done())
                        cleanup_release.set()
                        await asyncio.wait_for(pending, 2)
                        await ws.received(2)
                        create.assert_called_once()
                    self.assertTrue(created[0].closed)
                    self.assertTrue(ws.replies[-1]["ok"])
                    self.assertEqual(ws.replies[-1]["request_id"], "3" * 32)
                    self.assertEqual(replacement.transport.actions, ["begin"])
                    self.assertEqual(len(ws.replies), 2)
                    self.assertFalse(relay.pending)
                finally:
                    send_release.set()
                    cleanup_release.set()
                    await relay.stop()
                    if pending is not None:
                        with suppress(asyncio.CancelledError):
                            await pending
                    await replacement.stop()

    async def test_repeated_cancel_settles_worker_error_without_asyncio_diagnostic(self):
        entered, release, finished = threading.Event(), threading.Event(), threading.Event()
        loop = asyncio.get_running_loop()
        previous_handler = loop.get_exception_handler()
        diagnostics = []
        loop.set_exception_handler(lambda _loop, context: diagnostics.append(context))

        def failing_io():
            entered.set()
            try:
                release.wait(5)
                raise OSError("synthetic-private-adapter-error")
            finally:
                finished.set()

        task = asyncio.create_task(self.relay._thread(failing_io))
        try:
            self.assertTrue(await asyncio.to_thread(entered.wait, 1))
            task.cancel()
            await asyncio.sleep(0)
            task.cancel()
            await asyncio.sleep(0.01)
            self.assertFalse(task.done())
            release.set()
            with self.assertRaises(asyncio.CancelledError):
                await asyncio.wait_for(task, 2)
            await asyncio.sleep(0)
            self.assertTrue(finished.is_set())
            self.assertEqual(diagnostics, [])
        finally:
            release.set()
            with suppress(asyncio.CancelledError):
                await task
            loop.set_exception_handler(previous_handler)

    async def test_worker_error_maps_to_fixed_usb_failure_without_raw_details(self):
        def setup(transport):
            def exchange(request):
                raise OSError("synthetic-private-adapter-error")
            transport.exchange = exchange

        self.setup_transport = setup
        await self.relay.submit(self.request("begin"))
        await asyncio.wait_for(self.relay.task, 2)
        self.assertEqual(len(self.ws.replies), 1)
        self.assertFalse(self.ws.replies[0]["ok"])
        self.assertEqual(self.ws.replies[0]["reason"], "usb_io")
        self.assertNotIn("synthetic-private-adapter-error", json.dumps(self.ws.replies))
        self.assertTrue(self.created[0].closed)

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
