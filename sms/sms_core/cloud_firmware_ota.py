"""Bounded OTA relay over Air724UG's independent USB AT/data interface.

No modem reader is replaced. Every open is followed by a nonce/IMEI probe;
an active transfer never reopens a port or follows a changed cloud/serial link.
"""
import asyncio
import json
import re
import threading
import time
import uuid
from contextlib import suppress

from sms_core.cloud_command_context import capture_cloud_serial_command_context
from sms_core.serial_sender import reserve_serial_for_ota, release_serial_from_ota

HEX_ID = re.compile(r"[0-9a-f]{32}\Z")
MAX_FRAME = 6144
CAP_FIELDS = ("protocol", "supported", "core", "capacity", "version", "config_md5", "main_md5", "last_job")


class RelayError(Exception):
    """Only fixed protocol reasons, never raw serial data or exception text."""


def usb_at_candidates(ports, modem_port):
    ports = list(ports)
    modem = next((p for p in ports if p.device == modem_port), None)
    if modem is None or (modem.vid, modem.pid) != (0x1782, 0x4E00):
        return []
    if not re.search(r"LUAT USB Device 0 Modem\b", str(modem.description), re.I):
        return []
    return [p.device for p in ports if p.device != modem_port
            and (p.vid, p.pid) == (modem.vid, modem.pid)
            and re.search(r"LUAT USB Device 1 AT\b", str(p.description), re.I)][:8]


class UsbOtaTransport:
    def __init__(self, context, current, *, serial_factory=None, list_ports=None,
                 clock=time.monotonic, response_timeout=5):
        self.context, self.current = context, current
        self.serial_factory, self.list_ports = serial_factory, list_ports
        self.clock, self.response_timeout = clock, response_timeout
        self.session = uuid.uuid4().hex
        self.port = None
        self.buffer = bytearray()
        self.stopped = threading.Event()
        self.job_id = None
        self.manifest = None
        self.commit_sent = False
        self.capability = None

    def check(self):
        if self.stopped.is_set() or not self.current():
            raise RelayError("connection_changed")

    def _write(self, request):
        self.check()
        frame = json.dumps(dict(request, session=self.session, imei=self.context.imei),
                           ensure_ascii=True, separators=(",", ":")).encode("ascii") + b"\n"
        if len(frame) > MAX_FRAME:
            raise RelayError("frame_size")
        # Serial writes may be short even with write_timeout configured.
        offset = 0
        while offset < len(frame):
            self.check()
            count = self.port.write(frame[offset:])
            if not isinstance(count, int) or count <= 0 or count > len(frame) - offset:
                raise RelayError("usb_write")
            offset += count

    def _read(self, request, timeout, alternate=None):
        deadline = self.clock() + timeout
        while self.clock() < deadline:
            self.check()
            while b"\n" in self.buffer:
                raw, _, rest = self.buffer.partition(b"\n")
                self.buffer = bytearray(rest)
                if len(raw) > MAX_FRAME:
                    raise RelayError("frame_size")
                try:
                    reply = json.loads(raw)
                except (ValueError, UnicodeError, RecursionError):
                    continue
                if (not isinstance(reply, dict) or reply.get("type") != "firmware_ota_reply"
                        or reply.get("session") != self.session):
                    continue
                expected = request if reply.get("request_id") == request.get("request_id") else alternate
                if not expected or reply.get("request_id") != expected.get("request_id"):
                    continue
                if reply.get("imei") != self.context.imei:
                    raise RelayError("wrong_device")
                if expected["action"] != "info" and reply.get("job_id") != expected.get("job_id"):
                    continue
                self.check()
                return reply
            if len(self.buffer) > MAX_FRAME:
                raise RelayError("frame_size")
            self.buffer.extend(self.port.read(1024))
        raise RelayError("usb_timeout")

    def _probe(self):
        request = dict(type="firmware_ota", action="info", request_id=uuid.uuid4().hex)
        self._write(request)
        reply = self._read(request, 1)
        if reply.get("ok") is not True or not isinstance(reply.get("ota"), dict):
            raise RelayError("unsupported")
        self.capability = {key: reply["ota"].get(key) for key in CAP_FIELDS}
        return self.capability

    def connect(self):
        if self.port is not None:
            return
        self.check()
        if self.serial_factory is None:
            import serial
            self.serial_factory = serial.Serial
        if self.list_ports is None:
            from serial.tools.list_ports import comports
            self.list_ports = comports
        candidates = usb_at_candidates(self.list_ports(), getattr(self.context.serial, "port", ""))
        for name in candidates:
            try:
                self.check()
                self.port = self.serial_factory(name, baudrate=115200, timeout=0.1, write_timeout=2)
                self.buffer.clear()
                self.port.reset_input_buffer()
                self._probe()
                return
            except Exception:
                self._close_port()
                self.check()
        raise RelayError("usb_unavailable")

    def exchange(self, request):
        self.check()
        action = request["action"]
        if action != "info" and self.job_id and self.job_id != request["job_id"]:
            raise RelayError("busy")
        if action not in {"info", "begin"} and not self.job_id:
            raise RelayError("session")
        self.connect()
        if action == "info":
            capability = self._probe()
            return dict(type="firmware_ota_reply", action="info", request_id=request["request_id"],
                        ok=True, ota=capability, imei=self.context.imei)
        if not reserve_serial_for_ota(self.context.serial, self):
            raise RelayError("busy")
        if action == "begin":
            self.job_id = request["job_id"]
            self.manifest = request.get("manifest") or {}
        if action == "commit":
            # A lost ACK must not unlock commands against a possibly installing
            # device. This lease expires after the firmware's 120 s watchdog.
            self.commit_sent = True
        self._write(request)
        reply = self._read(request, self.response_timeout)
        if reply.get("ok") is not True or action == "abort":
            self.commit_sent = False
        return reply

    def wait_install_result(self, request):
        # Native failures use the original commit request ID. Reboot normally
        # disconnects USB; either outcome ends the reader without a retry.
        deadline = self.clock() + 120
        while self.clock() < deadline:
            try:
                reply = self._read(request, min(2, deadline - self.clock()))
            except RelayError as exc:
                if str(exc) != "usb_timeout":
                    raise
                # Some USB drivers keep their handle open across a soft reboot.
                # A fresh read-only nonce probe can prove that boot as well.
                probe = dict(type="firmware_ota", action="info", request_id=uuid.uuid4().hex)
                self._write(probe)
                try:
                    reply = self._read(probe, 1, alternate=request)
                except RelayError as exc:
                    if str(exc) == "usb_timeout":
                        continue
                    raise
            if reply.get("phase") == "failed":
                self.commit_sent = False
                return reply
            cap, target = reply.get("ota"), self.manifest or {}
            if (isinstance(cap, dict) and cap.get("supported") is True and cap.get("last_job") == self.job_id
                    and cap.get("version") == target.get("version") and cap.get("main_md5") == target.get("main_md5")):
                self.commit_sent = False
                return dict(ok=True, phase="rebooted")
        raise RelayError("usb_timeout")

    def _close_port(self):
        if self.port is not None:
            with suppress(Exception):
                self.port.close()
        self.port = None
        self.buffer.clear()

    def close(self):
        self._close_port()
        if not self.commit_sent:
            release_serial_from_ota(self.context.serial, self)

    def abort(self):
        if self.job_id and not self.commit_sent:
            with suppress(Exception):
                self._write(dict(type="firmware_ota", action="abort", job_id=self.job_id,
                                 request_id=uuid.uuid4().hex))


class CloudFirmwareOtaRelay:
    def __init__(self, namespace, websocket, *, transport_factory=UsbOtaTransport):
        self.namespace, self.websocket = namespace, websocket
        self.transport_factory = transport_factory
        self.queue = asyncio.Queue(maxsize=2)
        self.task = None
        self.transport = None
        self.closed = False
        self.pending = set()

    def _current(self, context):
        with self.namespace["serial_lock"]:
            stop_event = self.namespace.get("cloud_stop_event")
            return (not self.closed and not context.rejection_reason(self.namespace)
                    and not (stop_event is not None and stop_event.is_set())
                    and bool(getattr(context.serial, "is_open", False)))

    async def _reply(self, data, request, context):
        if not self._current(context):
            return
        # Whitelist replies: serial peers cannot inject arbitrary cloud events.
        payload = {key: data[key] for key in ("ok", "reason", "phase", "offset", "ota") if key in data}
        payload.update(type="firmware_ota_reply", action=request["action"],
                       request_id=request["request_id"], job_id=request.get("job_id"),
                       imei=context.imei, serial_connection_generation=context.generation)
        try:
            await self.websocket.send(json.dumps(payload, ensure_ascii=True, separators=(",", ":")))
        except Exception:
            # No exception repr: an adapter may include the outbound frame.
            self.closed = True

    async def submit(self, data):
        context = capture_cloud_serial_command_context(self.namespace, websocket=self.websocket)
        if (not self._current(context) or data.get("transport") != "desktop"
                or data.get("imei") != context.imei
                or type(data.get("serial_connection_generation")) is not int
                or data["serial_connection_generation"] != context.generation
                or not HEX_ID.fullmatch(str(data.get("request_id", "")))):
            return
        action = data.get("action")
        if action not in {"info", "begin", "chunk", "verify", "commit", "abort"}:
            return
        if action != "info" and not HEX_ID.fullmatch(str(data.get("job_id", ""))):
            return
        if action in {"begin", "commit"} and (not context.secret or data.get("secret") != context.secret):
            await self._reply(dict(ok=False, reason="auth"), data, context)
            return
        if len(json.dumps(data, ensure_ascii=True)) > MAX_FRAME:
            await self._reply(dict(ok=False, reason="frame_size"), data, context)
            return
        if data["request_id"] in self.pending:
            return  # An identical in-flight request gets its original reply.
        if self.queue.full() or (self.transport and self.transport.commit_sent):
            await self._reply(dict(ok=False, reason="busy"), data, context)
            return
        request = {key: data[key] for key in ("type", "action", "request_id", "job_id", "secret",
                                            "manifest", "data", "sha256", "offset") if key in data}
        self.pending.add(data["request_id"])
        self.queue.put_nowait((request, context))
        if self.task is None:
            self.task = asyncio.create_task(self._run(), name="usb-firmware-ota")

    async def _thread(self, function, *args):
        # Cancellation must join the bounded serial worker before closing its
        # handle or letting a replacement task open the same USB interface.
        work = asyncio.create_task(asyncio.to_thread(function, *args))
        try:
            return await asyncio.shield(work)
        except asyncio.CancelledError:
            if self.transport:
                self.transport.stopped.set()
            with suppress(Exception):
                await work
            raise

    async def _run(self):
        idle_since = time.monotonic()
        try:
            while not self.closed:
                if self.transport and not self._current(self.transport.context):
                    break
                try:
                    item = await asyncio.wait_for(self.queue.get(), 1)
                except asyncio.TimeoutError:
                    limit = 90 if self.transport and self.transport.job_id else 10
                    if time.monotonic() - idle_since >= limit:
                        break
                    continue
                if item is None:
                    break
                request, context = item
                try:
                    if self.transport is None:
                        self.transport = self.transport_factory(context, lambda ctx=context: self._current(ctx))
                    elif self.transport.context != context:
                        raise RelayError("connection_changed")
                    reply = await self._thread(self.transport.exchange, request)
                    await self._reply(reply, request, context)
                    if request["action"] == "commit" and reply.get("ok") is True:
                        try:
                            result = await self._thread(self.transport.wait_install_result, request)
                            await self._reply(result, request, context)
                        except Exception:
                            pass  # Installation is uncertain until a fresh probe.
                        break
                    if reply.get("ok") is not True or request["action"] == "abort":
                        break
                except Exception as exc:
                    reason = str(exc) if isinstance(exc, RelayError) else "usb_io"
                    if not (self.transport and self.transport.commit_sent):
                        await self._reply(dict(ok=False, reason=reason), request, context)
                    break
                finally:
                    self.pending.discard(request["request_id"])
                    idle_since = time.monotonic()
        finally:
            if self.transport:
                await self._thread(self.transport.abort)
                self.transport.close()
            self.closed = True
            self.pending.clear()
            while not self.queue.empty():
                self.queue.get_nowait()

    async def stop(self):
        self.closed = True
        if self.transport:
            self.transport.stopped.set()
        if not self.queue.full():
            self.queue.put_nowait(None)
        if self.task:
            with suppress(asyncio.CancelledError):
                await self.task


async def handle_firmware_ota(namespace, websocket, data):
    relay = namespace.get("_firmware_ota_relay")
    if relay and (relay.websocket is not websocket or relay.closed):
        await relay.stop()
        relay = None
    if relay is None:
        relay = CloudFirmwareOtaRelay(namespace, websocket)
        namespace["_firmware_ota_relay"] = relay
    await relay.submit(data)


async def stop_firmware_ota(namespace):
    relay = namespace.pop("_firmware_ota_relay", None)
    if relay:
        await relay.stop()
