import asyncio
import base64
import binascii
from dataclasses import dataclass
from datetime import datetime
import hashlib
import json
import os
import re
import threading
import time


RECORDING_ID_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9._-]{0,63}\Z")
IMEI_RE = re.compile(r"\d{14,17}\Z")
RECORDING_ID_IMEI_RE = re.compile(r"^rec-(\d{14,17})-")
SERIAL_RECORDING_PREFIX = "@@CALL_RECORD_"
MAX_CALL_RECORDING_BYTES = 512 * 1024
MAX_SERIAL_RECORDING_CHUNK_BYTES = 4096
# Serial transport frames are kept below the modem's ~128-byte line wrap.
# 512 KiB / 24-byte chunks needs at most 21846 chunks; leave a small margin.
MAX_SERIAL_RECORDING_CHUNKS = 24000
SERIAL_RECORDING_TIMEOUT_SECONDS = 30.0
CLOUD_RECORDING_CHUNK_BYTES = 3072
CLOUD_RECORDING_OFFER_TIMEOUT_SECONDS = 8.0
CLOUD_RECORDING_RESULT_TIMEOUT_SECONDS = 15.0
CLOUD_RECORDING_RETRY_SECONDS = 30.0


def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def _safe_int(value, default=0, minimum=None, maximum=None):
    try:
        number = int(value)
    except (TypeError, ValueError, OverflowError):
        number = int(default)
    if minimum is not None:
        number = max(int(minimum), number)
    if maximum is not None:
        number = min(int(maximum), number)
    return number


def _valid_recording_id(value):
    value = str(value or "").strip()
    return value if RECORDING_ID_RE.fullmatch(value) else ""


def _valid_imei(value):
    value = str(value or "").strip()
    return value if IMEI_RE.fullmatch(value) else ""


def _recording_imei_from_id(recording_id):
    match = RECORDING_ID_IMEI_RE.match(str(recording_id or "").strip())
    return _valid_imei(match.group(1)) if match else ""


def _decode_base64(value, *, max_bytes):
    value = str(value or "").strip()
    if not value or len(value) > max_bytes * 2:
        raise ValueError("base64 payload size is invalid")
    try:
        decoded = base64.b64decode(value.encode("ascii"), validate=True)
    except (UnicodeEncodeError, binascii.Error, ValueError) as exc:
        raise ValueError("base64 payload is invalid") from exc
    if len(decoded) > max_bytes:
        raise ValueError("decoded payload exceeds the limit")
    return decoded


def _safe_phone_filename(value):
    text = re.sub(r"[^0-9+]", "", str(value or ""))[:24]
    return text or "unknown"


def _recording_date(started_at):
    timestamp = _safe_int(started_at, 0)
    if timestamp < 946684800 or timestamp > int(time.time()) + 86400:
        timestamp = int(time.time())
    return datetime.fromtimestamp(timestamp), timestamp


@dataclass(frozen=True)
class SavedCallRecording:
    recording_id: str
    imei: str
    path: str
    metadata_path: str
    phone: str
    started_at: int
    duration_ms: int
    size: int
    sha256: str
    format: str = "amr"
    mime_type: str = "audio/amr"


class CallRecordingRepository:
    def __init__(self, recordings_dir, *, log_error=None):
        self.recordings_dir = os.path.abspath(recordings_dir)
        self.incoming_dir = os.path.join(self.recordings_dir, ".incoming")
        self.log_error = log_error
        self.lock = threading.RLock()
        self.records = {}
        os.makedirs(self.incoming_dir, exist_ok=True)
        self.cleanup_incoming()
        self.rebuild_index()

    def cleanup_incoming(self):
        removed = 0
        try:
            names = os.listdir(self.incoming_dir)
        except OSError:
            return 0
        for name in names:
            path = os.path.join(self.incoming_dir, name)
            if not os.path.isfile(path):
                continue
            try:
                os.remove(path)
                removed += 1
            except OSError:
                pass
        return removed

    def rebuild_index(self):
        records = {}
        with self.lock:
            for root, dirs, files in os.walk(self.recordings_dir):
                dirs[:] = [item for item in dirs if item != ".incoming"]
                for name in files:
                    if not name.endswith(".amr.json"):
                        continue
                    metadata_path = os.path.join(root, name)
                    metadata = self._read_metadata(metadata_path)
                    if not metadata:
                        continue
                    recording_id = _valid_recording_id(metadata.get("recording_id"))
                    imei = _valid_imei(metadata.get("imei")) or _recording_imei_from_id(
                        recording_id
                    )
                    if not recording_id:
                        continue
                    path = str(metadata.get("path") or metadata_path[:-5])
                    if not os.path.isabs(path):
                        path = os.path.join(self.recordings_dir, path)
                    path = os.path.abspath(path)
                    try:
                        inside_recordings = (
                            os.path.commonpath([self.recordings_dir, path])
                            == self.recordings_dir
                        )
                    except ValueError:
                        inside_recordings = False
                    if not inside_recordings or not os.path.isfile(path):
                        continue
                    normalized = dict(metadata)
                    normalized["recording_id"] = recording_id
                    normalized["imei"] = imei
                    records[(imei, recording_id)] = self._saved_from_metadata(
                        path,
                        metadata_path,
                        normalized,
                    )
            self.records = records
        return len(records)

    def incoming_path(self, recording_id):
        recording_id = _valid_recording_id(recording_id)
        if not recording_id:
            raise ValueError("invalid recording id")
        return os.path.join(self.incoming_dir, recording_id + ".part")

    def _target_paths(self, metadata):
        date_value, started_at = _recording_date(metadata.get("started_at"))
        imei = _valid_imei(metadata.get("imei")) or "unassigned"
        directory = os.path.join(
            self.recordings_dir,
            imei,
            date_value.strftime("%Y-%m-%d"),
        )
        filename = "{}_{}_{}.amr".format(
            date_value.strftime("%H%M%S"),
            _safe_phone_filename(metadata.get("phone")),
            metadata["recording_id"],
        )
        path = os.path.join(directory, filename)
        return path, path + ".json", started_at

    def _write_metadata(self, path, payload):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        temp_path = path + ".tmp"
        with open(temp_path, "w", encoding="utf-8") as file:
            json.dump(payload, file, ensure_ascii=False, indent=2)
            file.flush()
            os.fsync(file.fileno())
        os.replace(temp_path, path)

    def _read_metadata(self, path):
        try:
            with open(path, "r", encoding="utf-8-sig") as file:
                payload = json.load(file)
            return payload if isinstance(payload, dict) else None
        except (OSError, ValueError, TypeError):
            return None

    def find(self, recording_id, imei=""):
        recording_id = _valid_recording_id(recording_id)
        imei = _valid_imei(imei)
        if not recording_id:
            return None
        with self.lock:
            if imei:
                recording = self.records.get((imei, recording_id))
                return recording if recording and os.path.isfile(recording.path) else None
            matches = [
                recording
                for (source_imei, source_id), recording in self.records.items()
                if source_id == recording_id and os.path.isfile(recording.path)
            ]
            return matches[0] if len(matches) == 1 else None

    def _saved_from_metadata(self, path, metadata_path, metadata):
        return SavedCallRecording(
            recording_id=str(metadata.get("recording_id") or ""),
            imei=_valid_imei(metadata.get("imei"))
            or _recording_imei_from_id(metadata.get("recording_id")),
            path=path,
            metadata_path=metadata_path,
            phone=str(metadata.get("phone") or "unknown"),
            started_at=_safe_int(metadata.get("started_at"), 0),
            duration_ms=_safe_int(metadata.get("duration_ms"), 0, 0),
            size=_safe_int(metadata.get("size"), 0, 0),
            sha256=str(metadata.get("sha256") or ""),
            format=str(metadata.get("format") or "amr"),
            mime_type=str(metadata.get("mime_type") or "audio/amr"),
        )

    def commit(self, temp_path, metadata, sha256_hex):
        recording_id = _valid_recording_id(metadata.get("recording_id"))
        imei = _valid_imei(metadata.get("imei")) or _recording_imei_from_id(recording_id)
        if not recording_id:
            raise ValueError("invalid recording id")
        try:
            with open(temp_path, "rb") as file:
                header = file.read(6)
        except OSError as exc:
            raise ValueError("recording file is unavailable") from exc
        if header != b"#!AMR\n":
            raise ValueError("recording file format is invalid")
        with self.lock:
            existing = self.find(recording_id, imei)
            if existing is not None:
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
                return existing
            path, metadata_path, started_at = self._target_paths(
                {**dict(metadata), "imei": imei}
            )
            os.makedirs(os.path.dirname(path), exist_ok=True)
            os.replace(temp_path, path)
            payload = {
                "recording_id": recording_id,
                "imei": imei,
                "path": os.path.relpath(path, self.recordings_dir),
                "phone": str(metadata.get("phone") or "unknown")[:64],
                "started_at": started_at,
                "duration_ms": _safe_int(metadata.get("duration_ms"), 0, 0, 3600000),
                "size": os.path.getsize(path),
                "sha256": str(sha256_hex or ""),
                "format": "amr",
                "mime_type": "audio/amr",
                "upload_status": "pending",
                "upload_attempted_at": 0,
                "uploaded_at": 0,
                "upload_error": "",
            }
            try:
                self._write_metadata(metadata_path, payload)
            except Exception:
                try:
                    os.remove(path)
                except OSError:
                    pass
                raise
            saved = self._saved_from_metadata(path, metadata_path, payload)
            self.records[(saved.imei, saved.recording_id)] = saved
            return saved

    def _update_upload_state(self, recording, **updates):
        with self.lock:
            metadata = self._read_metadata(recording.metadata_path)
            if not metadata:
                return False
            metadata.update(updates)
            self._write_metadata(recording.metadata_path, metadata)
            updated = self._saved_from_metadata(
                recording.path,
                recording.metadata_path,
                metadata,
            )
            self.records[(updated.imei, updated.recording_id)] = updated
            return True

    def mark_uploading(self, recording):
        return self._update_upload_state(
            recording,
            upload_status="uploading",
            upload_attempted_at=int(time.time()),
            upload_error="",
        )

    def mark_uploaded(self, recording):
        return self._update_upload_state(
            recording,
            upload_status="uploaded",
            uploaded_at=int(time.time()),
            upload_error="",
        )

    def mark_pending(self, recording, error=""):
        return self._update_upload_state(
            recording,
            upload_status="pending",
            upload_error=str(error or "")[:240],
        )

    def pending(self, imei=""):
        records = []
        now = int(time.time())
        imei = _valid_imei(imei)
        with self.lock:
            for record in self.records.values():
                if imei and record.imei != imei:
                    continue
                metadata = self._read_metadata(record.metadata_path)
                if not metadata:
                    continue
                status = str(metadata.get("upload_status") or "pending")
                attempted_at = _safe_int(metadata.get("upload_attempted_at"), 0)
                if status == "uploaded":
                    continue
                if status == "uploading" and now - attempted_at < CLOUD_RECORDING_RETRY_SECONDS:
                    continue
                if os.path.isfile(record.path) and 0 < record.size <= MAX_CALL_RECORDING_BYTES:
                    records.append(record)
        records.sort(key=lambda item: (item.started_at, item.recording_id))
        return records


@dataclass
class _IncomingRecording:
    recording_id: str
    temp_path: str
    file: object
    metadata: dict
    expected_size: int
    expected_seq: int
    received_size: int
    chunk_count: int
    digest: object
    last_activity: float
    duplicate: bool = False


class SerialCallRecordingReceiver:
    def __init__(
        self,
        repository,
        *,
        on_saved=None,
        on_started=None,
        on_aborted=None,
        log_error=None,
        monotonic=None,
        source_imei=None,
    ):
        self.repository = repository
        self.on_saved = on_saved
        self.on_started = on_started
        self.on_aborted = on_aborted
        self.log_error = log_error
        self.monotonic = monotonic or time.monotonic
        self.source_imei = source_imei or (lambda: "")
        self.lock = threading.RLock()
        self.current = None

    def _abort(self, reason=""):
        transfer = self.current
        self.current = None
        if transfer is None:
            return
        try:
            if transfer.file is not None:
                transfer.file.close()
        except Exception:
            pass
        if not transfer.duplicate:
            try:
                os.remove(transfer.temp_path)
            except OSError:
                pass
            if self.on_aborted is not None:
                try:
                    self.on_aborted(dict(transfer.metadata), str(reason or "aborted"))
                except Exception as exc:
                    _safe_log(self.log_error, "Call recording abort callback failed: {!r}".format(exc))
        if reason:
            _safe_log(self.log_error, "Call recording serial transfer aborted: " + str(reason))

    def _expire_stale(self):
        if self.current is None:
            return
        if self.monotonic() - self.current.last_activity > SERIAL_RECORDING_TIMEOUT_SECONDS:
            self._abort("timeout")

    def expire_stale(self):
        with self.lock:
            before = self.current
            self._expire_stale()
            return before is not None and self.current is None

    def abort(self, reason=""):
        with self.lock:
            had_transfer = self.current is not None
            self._abort(reason or "aborted")
            return had_transfer

    def _begin(self, parts):
        if len(parts) not in (7, 8):
            raise ValueError("invalid begin frame")
        recording_id = _valid_recording_id(parts[1])
        if not recording_id:
            raise ValueError("invalid recording id")
        phone = _decode_base64(parts[2], max_bytes=64).decode("utf-8", errors="replace")[:64]
        started_at = _safe_int(parts[3], 0)
        duration_ms = _safe_int(parts[4], 0, 0, 3600000)
        expected_size = _safe_int(parts[5], 0)
        if expected_size <= 0 or expected_size > MAX_CALL_RECORDING_BYTES:
            raise ValueError("recording size exceeds the limit")
        if str(parts[6] or "").lower() != "amr":
            raise ValueError("unsupported recording format")
        source_imei = _valid_imei(parts[7] if len(parts) == 8 else "")
        if not source_imei:
            try:
                source_imei = _valid_imei(self.source_imei())
            except Exception:
                source_imei = ""
        source_imei = source_imei or _recording_imei_from_id(recording_id)
        self._abort("replaced")
        existing = self.repository.find(recording_id, source_imei)
        temp_path = self.repository.incoming_path(recording_id)
        file = None
        if existing is None:
            os.makedirs(os.path.dirname(temp_path), exist_ok=True)
            file = open(temp_path, "wb")
        self.current = _IncomingRecording(
            recording_id=recording_id,
            temp_path=temp_path,
            file=file,
            metadata={
                "recording_id": recording_id,
                "imei": source_imei,
                "phone": phone or "unknown",
                "started_at": started_at,
                "duration_ms": duration_ms,
                "size": expected_size,
                "format": "amr",
                "mime_type": "audio/amr",
            },
            expected_size=expected_size,
            expected_seq=1,
            received_size=0,
            chunk_count=0,
            digest=hashlib.sha256(),
            last_activity=self.monotonic(),
            duplicate=existing is not None,
        )
        if not self.current.duplicate and self.on_started is not None:
            try:
                self.on_started(dict(self.current.metadata))
            except Exception as exc:
                _safe_log(self.log_error, "Call recording start callback failed: {!r}".format(exc))

    def _chunk(self, parts):
        transfer = self.current
        if transfer is None or len(parts) != 4 or parts[1] != transfer.recording_id:
            raise ValueError("chunk without matching begin frame")
        sequence = _safe_int(parts[2], -1)
        if sequence != transfer.expected_seq or sequence > MAX_SERIAL_RECORDING_CHUNKS:
            raise ValueError("recording chunk sequence mismatch")
        decoded = _decode_base64(parts[3], max_bytes=MAX_SERIAL_RECORDING_CHUNK_BYTES)
        if not decoded:
            raise ValueError("empty recording chunk")
        next_size = transfer.received_size + len(decoded)
        if next_size > transfer.expected_size or next_size > MAX_CALL_RECORDING_BYTES:
            raise ValueError("recording chunk exceeds declared size")
        if not transfer.duplicate:
            transfer.file.write(decoded)
            transfer.digest.update(decoded)
        transfer.received_size = next_size
        transfer.chunk_count += 1
        transfer.expected_seq += 1
        transfer.last_activity = self.monotonic()

    def _end(self, parts):
        transfer = self.current
        if transfer is None or len(parts) != 4 or parts[1] != transfer.recording_id:
            raise ValueError("end without matching begin frame")
        chunk_count = _safe_int(parts[2], -1)
        declared_size = _safe_int(parts[3], -1)
        if chunk_count != transfer.chunk_count:
            raise ValueError("recording chunk count mismatch")
        if declared_size != transfer.expected_size or transfer.received_size != transfer.expected_size:
            raise ValueError("recording size mismatch")
        self.current = None
        if transfer.duplicate:
            return self.repository.find(
                transfer.recording_id,
                transfer.metadata.get("imei"),
            )
        try:
            transfer.file.flush()
            os.fsync(transfer.file.fileno())
            transfer.file.close()
            transfer.file = None
            saved = self.repository.commit(
                transfer.temp_path,
                transfer.metadata,
                transfer.digest.hexdigest(),
            )
        except Exception:
            try:
                if transfer.file is not None:
                    transfer.file.close()
            except Exception:
                pass
            try:
                os.remove(transfer.temp_path)
            except OSError:
                pass
            if self.on_aborted is not None:
                try:
                    self.on_aborted(dict(transfer.metadata), "save_failed")
                except Exception as exc:
                    _safe_log(self.log_error, "Call recording save failure callback failed: {!r}".format(exc))
            raise
        if self.on_saved is not None:
            try:
                self.on_saved(saved)
            except Exception as exc:
                _safe_log(self.log_error, "Call recording saved callback failed: {!r}".format(exc))
        return saved

    def consume_line(self, line):
        with self.lock:
            text = str(line or "").strip()
            if not text.startswith(SERIAL_RECORDING_PREFIX):
                self._expire_stale()
                return False
            try:
                parts = text.split("|")
                frame_type = parts[0]
                if frame_type == "@@CALL_RECORD_BEGIN":
                    self._begin(parts)
                elif frame_type == "@@CALL_RECORD_CHUNK":
                    self._chunk(parts)
                elif frame_type == "@@CALL_RECORD_END":
                    self._end(parts)
                elif frame_type == "@@CALL_RECORD_ABORT":
                    self._abort("firmware_abort")
                else:
                    raise ValueError("unknown recording frame")
            except Exception as exc:
                self._abort(type(exc).__name__)
            return True


class CloudCallRecordingUploader:
    def __init__(self, repository, *, log_error=None):
        self.repository = repository
        self.log_error = log_error
        self._task = None
        self._next_schedule = None
        self._offer_waiters = {}
        self._result_waiters = {}

    def handle_server_message(self, data):
        data = data if isinstance(data, dict) else {}
        msg_type = str(data.get("type") or "")
        if msg_type not in ("call_recording_offer_ack", "call_recording_result"):
            return False
        recording_id = _valid_recording_id(data.get("recording_id"))
        if not recording_id:
            return True
        waiters = self._offer_waiters if msg_type == "call_recording_offer_ack" else self._result_waiters
        future = waiters.pop(recording_id, None)
        if future is not None and not future.done():
            future.set_result(dict(data))
        return True

    @staticmethod
    def _send_succeeded(result):
        return result is True or result == "sent"

    async def _send_recording(self, recording, ws, send_payload, identity_payload):
        loop = asyncio.get_running_loop()
        identity = dict(identity_payload() or {})
        current_imei = _valid_imei(identity.get("imei") or identity.get("device_imei"))
        if not current_imei or recording.imei != current_imei:
            return False, "source_imei_mismatch"
        offer_future = loop.create_future()
        self._offer_waiters[recording.recording_id] = offer_future
        offer = {
            "type": "call_recording_offer",
            "recording_id": recording.recording_id,
            "phone": recording.phone,
            "started_at": recording.started_at,
            "duration_ms": recording.duration_ms,
            "size": recording.size,
            "sha256": recording.sha256,
            "format": recording.format,
            "mime_type": recording.mime_type,
            "source_imei": recording.imei,
            **identity,
        }
        try:
            if not self._send_succeeded(await send_payload(ws, offer)):
                return False, "offer_send_failed"
            ack = await asyncio.wait_for(offer_future, CLOUD_RECORDING_OFFER_TIMEOUT_SECONDS)
        except asyncio.TimeoutError:
            return False, "offer_timeout"
        finally:
            self._offer_waiters.pop(recording.recording_id, None)
        if ack.get("already_uploaded") is True:
            return True, "already_uploaded"
        if ack.get("ok") is not True or ack.get("accepted") is not True:
            return False, str(ack.get("reason") or ack.get("message") or "offer_rejected")

        sequence = 0
        try:
            with open(recording.path, "rb") as file:
                while True:
                    chunk = file.read(CLOUD_RECORDING_CHUNK_BYTES)
                    if not chunk:
                        break
                    sequence += 1
                    payload = {
                        "type": "call_recording_chunk",
                        "recording_id": recording.recording_id,
                        "seq": sequence,
                        "data": base64.b64encode(chunk).decode("ascii"),
                        "source_imei": recording.imei,
                        **identity,
                    }
                    if not self._send_succeeded(await send_payload(ws, payload)):
                        return False, "chunk_send_failed"
        except OSError:
            return False, "recording_file_unavailable"

        result_future = loop.create_future()
        self._result_waiters[recording.recording_id] = result_future
        try:
            end_payload = {
                "type": "call_recording_end",
                "recording_id": recording.recording_id,
                "chunk_count": sequence,
                "size": recording.size,
                "sha256": recording.sha256,
                "source_imei": recording.imei,
                **identity,
            }
            if not self._send_succeeded(await send_payload(ws, end_payload)):
                return False, "end_send_failed"
            result = await asyncio.wait_for(result_future, CLOUD_RECORDING_RESULT_TIMEOUT_SECONDS)
        except asyncio.TimeoutError:
            return False, "result_timeout"
        finally:
            self._result_waiters.pop(recording.recording_id, None)
        return (
            result.get("ok") is True,
            str(result.get("reason") or result.get("message") or "upload_failed"),
        )

    async def _drain(self, ws, send_payload, identity_payload, is_current, is_authorized):
        try:
            identity = dict(identity_payload() or {})
            current_imei = _valid_imei(identity.get("imei") or identity.get("device_imei"))
            if not current_imei:
                return
            for recording in self.repository.pending(current_imei):
                if not is_current(ws) or not is_authorized():
                    break
                if not self.repository.mark_uploading(recording):
                    _safe_log(
                        self.log_error,
                        "Unable to persist call recording upload state",
                    )
                    break
                ok, reason = await self._send_recording(
                    recording,
                    ws,
                    send_payload,
                    identity_payload,
                )
                if ok:
                    if not self.repository.mark_uploaded(recording):
                        _safe_log(
                            self.log_error,
                            "Unable to persist uploaded call recording state",
                        )
                        break
                else:
                    if not self.repository.mark_pending(recording, reason):
                        _safe_log(
                            self.log_error,
                            "Unable to persist pending call recording state",
                        )
                    break
        except Exception as exc:
            _safe_log(self.log_error, "Upload call recording failed: {!r}".format(exc))
        finally:
            next_schedule = self._next_schedule
            self._next_schedule = None
            self._task = None
            if next_schedule is not None:
                next_ws = next_schedule[0]
                next_is_current = next_schedule[3]
                next_is_authorized = next_schedule[4]
                if next_is_current(next_ws) and next_is_authorized():
                    self._task = asyncio.get_running_loop().create_task(
                        self._drain(*next_schedule)
                    )

    def schedule(self, loop, ws, *, send_payload, identity_payload, is_current, is_authorized):
        if loop is None or not loop.is_running() or ws is None or not is_authorized():
            return False

        def start():
            schedule_args = (
                ws,
                send_payload,
                identity_payload,
                is_current,
                is_authorized,
            )
            if self._task is not None and not self._task.done():
                self._next_schedule = schedule_args
                return
            self._task = loop.create_task(
                self._drain(*schedule_args)
            )

        try:
            loop.call_soon_threadsafe(start)
            return True
        except Exception as exc:
            _safe_log(self.log_error, "Schedule call recording upload failed: {!r}".format(exc))
            return False
