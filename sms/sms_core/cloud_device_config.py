"""Short configuration transactions on the verified independent USB AT port."""
import asyncio
import json
import math
import re
import uuid

from sms_core.cloud_command_context import capture_cloud_serial_command_context
from sms_core.cloud_firmware_ota import (
    CloudFirmwareOtaRelay, HEX_ID, MAX_FRAME, RelayError, UsbOtaTransport,
)
from sms_core.serial_sender import release_serial_from_ota, reserve_serial_for_ota


CONFIG_KEY = re.compile(r"[A-Z][A-Z0-9_]{0,79}\Z")
CONFIG_VERSION = re.compile(r"[0-9]+\.[0-9]+\.[0-9]+\Z")
MAX_REVISION = 2147483647
MAX_JSON_NUMBER = 9007199254740991
BOOL_FIELDS = (
    "ok", "supported", "saved", "applied", "accepted", "restart_required",
    "allow_nvm_server_config", "cellular_data_enabled", "console_connection_enabled", "console_connection_requested",
)
REASONS = frozenset((
    "config_conflict", "config_persist", "config_invalid", "invalid_request", "unsupported",
    "busy", "serial_busy", "session", "wrong_device", "connection_changed",
    "usb_unavailable", "usb_timeout", "usb_write", "frame_size", "config_unconfirmed",
    "restart_unconfirmed", "auth", "not_allowed", "storage", "config_apply", "restart_schedule",
))


def sensitive_config_key(key):
    return (key in {"PIN_CODE", "CUSTOM_POST_BODY_TABLE"}
            or any(part in key for part in ("SECRET", "TOKEN", "PASSWORD", "PASSWD", "WEBHOOK", "_KEY", "_API", "_URL")))


def valid_revision(value):
    return type(value) is int and 0 <= value <= MAX_REVISION


def valid_json_value(value, depth=0):
    if depth > 6:
        return False
    if value is None or type(value) in (bool, str):
        return not isinstance(value, str) or len(value) <= MAX_FRAME
    if type(value) is int:
        return abs(value) <= MAX_JSON_NUMBER
    if type(value) is float:
        return math.isfinite(value) and abs(value) <= MAX_JSON_NUMBER
    if isinstance(value, list):
        return len(value) <= 128 and all(valid_json_value(item, depth + 1) for item in value)
    if isinstance(value, dict):
        return len(value) <= 128 and all(isinstance(key, str) and len(key) <= 128
                                        and valid_json_value(item, depth + 1) for key, item in value.items())
    return False


def validate_config_request(data):
    if (not isinstance(data, dict) or data.get("type") != "device_config"
            or type(data.get("protocol")) is not int or data["protocol"] != 1
            or not isinstance(data.get("action"), str) or data["action"] not in {"info", "set", "restart"}
            or not isinstance(data.get("request_id"), str) or not HEX_ID.fullmatch(data["request_id"])):
        raise RelayError("invalid_request")
    request = {key: data[key] for key in ("type", "protocol", "action", "request_id")}
    if data["action"] != "info":
        if not valid_revision(data.get("expected_revision")):
            raise RelayError("invalid_request")
        request["expected_revision"] = data["expected_revision"]
    if data["action"] == "set":
        configs = data.get("configs")
        if (not isinstance(configs, dict) or not configs or len(configs) > 128
                or any(not isinstance(key, str) or not CONFIG_KEY.fullmatch(key)
                       or not valid_json_value(value) for key, value in configs.items())):
            raise RelayError("invalid_request")
        # Only explicit fields are sent. An empty string here is an explicit
        # clear, while untouched password placeholders are omitted upstream.
        request["configs"] = json.loads(json.dumps(configs, ensure_ascii=True, allow_nan=False))
    if len(json.dumps(request, ensure_ascii=True, allow_nan=False)) > MAX_FRAME - 160:
        raise RelayError("frame_size")
    return request


def safe_config_reply(data):
    if not isinstance(data, dict) or type(data.get("ok")) is not bool:
        raise RelayError("config_invalid")
    result = {key: data[key] for key in BOOL_FIELDS if type(data.get(key)) is bool}
    if "server_config_batch" in data:
        if type(data["server_config_batch"]) is not bool:
            raise RelayError("config_invalid")
        result["server_config_batch"] = data["server_config_batch"]
    if type(data.get("protocol")) is int and data["protocol"] == 1:
        result["protocol"] = 1
    if valid_revision(data.get("revision")):
        result["revision"] = data["revision"]
    if isinstance(data.get("version"), str) and CONFIG_VERSION.fullmatch(data["version"]):
        result["version"] = data["version"]
    if isinstance(data.get("phase"), str) and data["phase"] in {"accepted", "failed", "restarted"}:
        result["phase"] = data["phase"]
    if "reason" in data:
        result["reason"] = data["reason"] if isinstance(data["reason"], str) and data["reason"] in REASONS else "config_invalid"
    for values_key in ("values", "effective_values"):
        if isinstance(data.get(values_key), dict):
            result[values_key] = {key: value for key, value in data[values_key].items()
                                  if isinstance(key, str) and CONFIG_KEY.fullmatch(key)
                                  and not sensitive_config_key(key) and valid_json_value(value)}
    if isinstance(data.get("configured"), dict):
        result["configured"] = {key: value for key, value in data["configured"].items()
                                if isinstance(key, str) and CONFIG_KEY.fullmatch(key) and type(value) is bool}
    for key in ("keys", "immediate_keys", "restart_required_keys"):
        if isinstance(data.get(key), list):
            result[key] = [item for item in data[key] if isinstance(item, str) and CONFIG_KEY.fullmatch(item)][:128]
    for key in ("supported_fields", "file_only"):
        if key not in data:
            continue
        entries = data[key]
        if type(entries) is dict and not entries:
            entries = []
        if (not isinstance(entries, list) or len(entries) > 128
                or any(not isinstance(item, str) or not CONFIG_KEY.fullmatch(item) for item in entries)):
            raise RelayError("config_invalid")
        result[key] = list(entries)
    if isinstance(data.get("apply_errors"), dict):
        result["apply_errors"] = {key: "config_apply" for key in data["apply_errors"]
                                  if isinstance(key, str) and CONFIG_KEY.fullmatch(key)}
    return result


class UsbDeviceConfigTransport(UsbOtaTransport):
    reply_type = "device_config_reply"

    def _probe_request(self):
        return dict(type="device_config", protocol=1, action="info", request_id=uuid.uuid4().hex)

    def _reply_matches(self, request, reply):
        return (reply.get("action") == request["action"]
                and type(reply.get("protocol")) is int and reply["protocol"] == 1)

    def _probe_capability(self, reply):
        capability = safe_config_reply(reply)
        if capability.get("ok") is not True or capability.get("supported") is not True:
            raise RelayError("unsupported")
        if not valid_revision(capability.get("revision")):
            raise RelayError("config_invalid")
        return capability

    def exchange(self, data):
        request = validate_config_request(data)
        self.check()
        if not reserve_serial_for_ota(self.context.serial, self, seconds=15, wait_timeout=2,
                                      before_reserve=self.check):
            self.check()
            raise RelayError("serial_busy")
        attempted = False
        try:
            self.check()
            opened = self.connect()
            if request["action"] == "info":
                return dict(self.capability if opened else self._probe())
            attempted = True
            self._write(request)
            reply = self._read(request, self.response_timeout)
            result = safe_config_reply(reply)
            if result.get("ok") is True and request["action"] == "set":
                if result.get("saved") is not True or not valid_revision(result.get("revision")):
                    raise RelayError("config_unconfirmed")
            if result.get("ok") is True and request["action"] == "restart" and result.get("phase") != "accepted":
                raise RelayError("restart_unconfirmed")
            return result
        except Exception as exc:
            if attempted and (not isinstance(exc, RelayError) or str(exc) not in REASONS
                              or str(exc) in {"usb_timeout", "usb_write", "frame_size", "config_invalid"}):
                raise RelayError("config_unconfirmed" if request["action"] == "set" else "restart_unconfirmed") from None
            raise
        finally:
            self.close()

    def abort(self):
        return None

    def close(self):
        self._close_port()
        release_serial_from_ota(self.context.serial, self)


class CloudDeviceConfigRelay(CloudFirmwareOtaRelay):
    def __init__(self, namespace, websocket, *, transport_factory=UsbDeviceConfigTransport):
        super().__init__(namespace, websocket, transport_factory=transport_factory)

    async def _reply(self, data, request, context):
        if not self._current(context):
            return
        try:
            payload = safe_config_reply(data)
            payload.update(type="device_config_reply", protocol=1, action=request["action"],
                           request_id=request["request_id"], imei=context.imei,
                           serial_connection_generation=context.generation)
            await self.websocket.send(json.dumps(payload, ensure_ascii=True, allow_nan=False, separators=(",", ":")))
        except Exception:
            self.closed = True

    async def submit(self, data):
        context = capture_cloud_serial_command_context(self.namespace, websocket=self.websocket)
        if (not isinstance(data, dict) or not self._current(context) or data.get("transport") != "desktop"
                or data.get("imei") != context.imei or type(data.get("serial_connection_generation")) is not int
                or data["serial_connection_generation"] != context.generation
                or not isinstance(data.get("request_id"), str) or not HEX_ID.fullmatch(data["request_id"])):
            return
        try:
            request = validate_config_request(data)
        except (RelayError, ValueError, TypeError, OverflowError):
            request = dict(action=data.get("action") if isinstance(data.get("action"), str) else "",
                           request_id=data["request_id"])
            await self._reply(dict(ok=False, reason="invalid_request"), request, context)
            return
        if request["request_id"] in self.pending:
            return
        if self.closing or self.queue.full():
            await self._reply(dict(ok=False, reason="busy"), request, context)
            return
        self.pending.add(request["request_id"])
        self.queue.put_nowait((request, context))
        if self.task is None:
            self.task = asyncio.create_task(self._run(), name="usb-device-config")

    async def _run(self):
        try:
            while not self.closed and not self.queue.empty():
                item = self.queue.get_nowait()
                if item is None:
                    break
                request, context = item
                try:
                    if not self._current(context):
                        continue
                    self.transport = self.transport_factory(context, lambda ctx=context: self._current(ctx))
                    reply = await self._thread(self.transport.exchange, request)
                except Exception as exc:
                    fallback = "restart_unconfirmed" if request["action"] == "restart" else "config_unconfirmed"
                    reason = str(exc) if isinstance(exc, RelayError) and str(exc) in REASONS else fallback
                    reply = dict(ok=False, reason=reason)
                finally:
                    if self.transport:
                        self.transport.close()
                        self.transport = None
                try:
                    await self._reply(reply, request, context)
                finally:
                    self.pending.discard(request["request_id"])
        finally:
            self.closing = True
            if self.transport:
                self.transport.close()
                self.transport = None
            while not self.queue.empty():
                item = self.queue.get_nowait()
                if item is None:
                    continue
                request, context = item
                await self._reply(dict(ok=False, reason="connection_changed"), request, context)
            self.pending.clear()
            self.closed = True


async def handle_device_config(namespace, websocket, data):
    relay = namespace.get("_device_config_relay")
    if relay and relay.closing and not relay.closed:
        await asyncio.shield(relay.task)
    if relay and (relay.websocket is not websocket or relay.closed):
        await relay.stop()
        relay = None
    if relay is None:
        relay = CloudDeviceConfigRelay(namespace, websocket)
        namespace["_device_config_relay"] = relay
    await relay.submit(data)


async def stop_device_config(namespace):
    relay = namespace.pop("_device_config_relay", None)
    if relay:
        await relay.stop()
