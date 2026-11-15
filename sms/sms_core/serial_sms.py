from dataclasses import dataclass

import base64
import binascii

from sms_core.serial_parsers import is_sms_collection_boundary
from sms_core.sms_processing import process_pending_sms


SMS_CALLBACK_PREFIX = "[I]-[handler_sms.smsCallback]"
SMS_FRAME_PREFIX = "@@SMS_V1_"
SMS_FRAME_CHUNK_BYTES = 24
SMS_FRAME_MAX_BYTES = 131072
SMS_FRAME_MAX_ID = 999999999


@dataclass
class PendingSms:
    callback_head: str
    full_msg: str
    display_lines: list
    concat_info: object = None
    concat_body: object = None
    concat_sender: object = None
    concat_timestamp: object = None
    concat_reference: object = None
    concat_reference_bits: object = None
    concat_total: object = None
    message_trace_id: object = None
    concat_sender_is_alphanumeric: bool = False


@dataclass
class CollectedSmsCallback:
    callback_text: str
    raw_lines: list
    framed_body: str | None = None


@dataclass
class SmsLineDecision:
    started: bool = False
    action: str = "pass"
    flushed: bool = False
    continue_read: bool = False


class SmsPendingCollector:
    """Collect legacy lines and complete encoded frames without SMS business logic."""

    def __init__(
        self,
        parse_callback_head=None,
        correction_cache=None,
        initial_timeout: float = 1.0,
        continuation_timeout: float = 0.4,
    ):
        self.parse_callback_head = parse_callback_head
        self.initial_timeout = initial_timeout
        self.continuation_timeout = continuation_timeout
        self.reset()

    def reset(self):
        self.callback_head = ""
        self.raw_lines = []
        self.deadline = 0.0
        self.active = False
        self.framed_body = None
        self.frame_active = False
        self._frame_id = None
        self._frame_length = 0
        self._frame_count = 0
        self._frame_next_seq = 1
        self._frame_payload = bytearray()
        self._frame_diagnostic = None

    def expired(self, now: float) -> bool:
        return self.active and now > self.deadline

    def start(self, callback_text: str, now: float) -> bool:
        text = str(callback_text or "").strip()
        self.reset()
        if not text:
            return False

        self.callback_head = text
        self.raw_lines = []
        self.deadline = now + self.initial_timeout
        self.active = True
        return True

    def flush(self):
        if self.frame_active:
            self._drop_frame("INCOMPLETE")
            return None
        if not self.active:
            return None

        collected = CollectedSmsCallback(
            callback_text=self.callback_head,
            raw_lines=list(self.raw_lines),
            framed_body=self.framed_body,
        )
        self.reset()
        return collected

    def consume_line(self, line: str, now: float) -> str:
        if not self.active or self.frame_active:
            return "pass"
        if is_sms_collection_boundary(line):
            return "boundary"

        self.raw_lines.append(line)
        self.deadline = now + self.continuation_timeout
        return "consumed"

    def _drop_frame(self, reason):
        diagnostic = self._frame_diagnostic
        self.reset()
        if diagnostic is not None:
            diagnostic(f"[SMS] FRAME DROP Reason={reason}")

    def consume_frame_line(self, line, diagnostic_callback=None):
        if diagnostic_callback is not None:
            self._frame_diagnostic = diagnostic_callback
        fields = line.split("|")
        if fields[0] == SMS_FRAME_PREFIX + "BEGIN":
            if self.frame_active:
                self._drop_frame("REPLACED")
            self._frame_diagnostic = diagnostic_callback
            if len(fields) != 4:
                self._drop_frame("INVALID_BEGIN")
                return "consumed"
            frame_id = _frame_number(fields[1], SMS_FRAME_MAX_ID)
            length = _frame_number(fields[2], SMS_FRAME_MAX_BYTES)
            maximum_chunks = (SMS_FRAME_MAX_BYTES + SMS_FRAME_CHUNK_BYTES - 1) // SMS_FRAME_CHUNK_BYTES
            count = _frame_number(fields[3], maximum_chunks)
            if frame_id is None or length is None or count != (length + SMS_FRAME_CHUNK_BYTES - 1) // SMS_FRAME_CHUNK_BYTES:
                self._drop_frame("INVALID_BEGIN")
                return "consumed"
            self.active = True
            self.frame_active = True
            self.deadline = float("inf")
            self._frame_id = frame_id
            self._frame_length = length
            self._frame_count = count
            return "start"

        if not self.frame_active:
            self._drop_frame("ORPHAN_FRAME")
            return "consumed"
        if fields[0] == SMS_FRAME_PREFIX + "CHUNK":
            if len(fields) != 4:
                self._drop_frame("INVALID_CHUNK")
                return "consumed"
            frame_id = _frame_number(fields[1], SMS_FRAME_MAX_ID)
            sequence = _frame_number(fields[2], self._frame_count)
            encoded = fields[3]
            if frame_id != self._frame_id or sequence != self._frame_next_seq or not 1 <= len(encoded) <= 32:
                self._drop_frame("INVALID_CHUNK")
                return "consumed"
            try:
                chunk = base64.b64decode(encoded, validate=True)
            except (ValueError, binascii.Error):
                self._drop_frame("INVALID_BASE64")
                return "consumed"
            expected_length = min(SMS_FRAME_CHUNK_BYTES, self._frame_length - len(self._frame_payload))
            if len(chunk) != expected_length or base64.b64encode(chunk).decode("ascii") != encoded:
                self._drop_frame("INVALID_CHUNK_LENGTH")
                return "consumed"
            self._frame_payload.extend(chunk)
            self._frame_next_seq += 1
            return "consumed"

        if fields[0] == SMS_FRAME_PREFIX + "END":
            if (
                len(fields) != 3
                or _frame_number(fields[1], SMS_FRAME_MAX_ID) != self._frame_id
                or _frame_number(fields[2], self._frame_count) != self._frame_count
            ):
                self._drop_frame("INVALID_END")
                return "consumed"
            if self._frame_next_seq != self._frame_count + 1 or len(self._frame_payload) != self._frame_length:
                self._drop_frame("INCOMPLETE")
                return "consumed"
            payload_fields = bytes(self._frame_payload).split(b"\0", 2)
            if len(payload_fields) != 3 or len(payload_fields[0]) > 128 or len(payload_fields[1]) > 64:
                self._drop_frame("INVALID_PAYLOAD")
                return "consumed"
            try:
                sender, timestamp, body = (field.decode("utf-8", errors="strict") for field in payload_fields)
            except UnicodeDecodeError:
                self._drop_frame("INVALID_UTF8")
                return "consumed"
            self.callback_head = f"{sender} {timestamp} {body}"
            self.framed_body = body
            self.frame_active = False
            self._frame_payload.clear()
            return "complete"

        self._drop_frame("UNKNOWN_FRAME")
        return "consumed"


def _frame_number(value, maximum):
    text = str(value or "")
    if not text or len(text) > len(str(maximum)) or not text.isascii() or not text.isdigit():
        return None
    number = int(text)
    return number if 1 <= number <= maximum else None


def callback_body_from_line(line: str, prefix: str = SMS_CALLBACK_PREFIX) -> str:
    text = str(line or "").lstrip()
    if not text.startswith(prefix):
        return ""
    return text[len(prefix):].strip()


def handle_sms_collector_line(collector, line: str, now: float, flush_callback, prefix: str = SMS_CALLBACK_PREFIX, diagnostic_callback=None):
    text = str(line or "").lstrip()
    legacy_active = collector.active and not collector.frame_active and collector.framed_body is None
    if not legacy_active and text.startswith(SMS_FRAME_PREFIX):
        action = collector.consume_frame_line(text, diagnostic_callback)
        complete = action == "complete"
        if complete:
            flush_callback()
        return SmsLineDecision(
            started=action == "start",
            action="start" if action == "start" else "consumed",
            flushed=complete,
            continue_read=True,
        )
    if str(line or "").lstrip().startswith(prefix):
        flushed = False
        if collector.active:
            flush_callback()
            flushed = True
        collector.start(callback_body_from_line(line, prefix), now)
        return SmsLineDecision(
            started=True,
            action="start",
            flushed=flushed,
            continue_read=True,
        )

    if collector.active:
        action = collector.consume_line(line, now)
        flushed = action in ("boundary", "flush")
        if flushed:
            flush_callback()
        return SmsLineDecision(
            started=False,
            action=action,
            flushed=flushed,
            continue_read=action in ("consumed", "flush"),
        )

    return SmsLineDecision()


def flush_pending_sms(
    collector,
    keywords,
    log_unmatched_sms,
    log_dir,
    log_prefix,
    ignore_repeat_state,
    error_repeat_limit,
    enqueue_push,
    send_cloud_sms_event,
    port_ui,
    play_alert,
    show_sms_popup,
    file_log,
    system_ui,
    assembler=None,
    now=None,
    concat_log=None,
):
    collected = collector.flush()
    if collected is None:
        pending = None
    elif hasattr(assembler, "add_collected"):
        pending = assembler.add_collected(collected, now=now, log=concat_log)
        if pending is None:
            return "pending"
    else:
        callback_head = collected.callback_text
        _sender, body = collector.parse_callback_head(callback_head) if getattr(collector, "parse_callback_head", None) else ("", callback_head)
        raw_lines = list(getattr(collected, "raw_lines", []) or [])
        framed_body = getattr(collected, "framed_body", None)
        full_msg = framed_body if framed_body is not None else "\n".join([body or callback_head] + raw_lines).strip()
        pending = PendingSms(
            callback_head=callback_head,
            full_msg=full_msg,
            display_lines=["📩 收到短信：", callback_head] + raw_lines,
        )
    if assembler is not None and not hasattr(assembler, "add_collected"):
        pending = assembler.add_message(pending, now=now, log=concat_log)
        if pending is None:
            return "pending"
    return process_pending_sms_items(
        pending,
        keywords,
        log_unmatched_sms,
        log_dir,
        log_prefix,
        ignore_repeat_state,
        error_repeat_limit,
        enqueue_push,
        send_cloud_sms_event,
        port_ui,
        play_alert,
        show_sms_popup,
        file_log,
        system_ui,
    )


def process_pending_sms_items(
    pending,
    keywords,
    log_unmatched_sms,
    log_dir,
    log_prefix,
    ignore_repeat_state,
    error_repeat_limit,
    enqueue_push,
    send_cloud_sms_event,
    port_ui,
    play_alert,
    show_sms_popup,
    file_log,
    system_ui,
):
    if pending is None:
        return "pending"
    pending_items = pending if isinstance(pending, list) else [pending]
    result = "pending"
    for pending_item in pending_items:
        result = process_pending_sms(
            pending_item,
            keywords,
            log_unmatched_sms,
            log_dir,
            log_prefix,
            ignore_repeat_state,
            error_repeat_limit,
            enqueue_push,
            send_cloud_sms_event,
            port_ui,
            play_alert,
            show_sms_popup,
            file_log,
            system_ui,
        )
    return result
