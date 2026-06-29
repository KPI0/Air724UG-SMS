from dataclasses import dataclass

from sms_core.serial_parsers import is_sms_collection_boundary
from sms_core.sms_processing import process_pending_sms


SMS_CALLBACK_PREFIX = "[I]-[handler_sms.smsCallback]"


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


@dataclass
class CollectedSmsCallback:
    callback_text: str
    follow_lines: list


@dataclass
class SmsLineDecision:
    started: bool = False
    action: str = "pass"
    flushed: bool = False
    continue_read: bool = False


class SmsPendingCollector:
    """Collect raw smsCallback frame lines; long-SMS semantics live in LongSmsAssembler."""

    def __init__(
        self,
        parse_callback_head=None,
        correction_cache=None,
        initial_timeout: float = 1.0,
        fragment_timeout: float = 0.4,
        max_follow_lines=None,
    ):
        self.parse_callback_head = parse_callback_head
        self.initial_timeout = initial_timeout
        self.fragment_timeout = fragment_timeout
        self.max_follow_lines = max_follow_lines
        self.reset()

    def reset(self):
        self.callback_head = ""
        self.follow_lines = []
        self.deadline = 0.0
        self.active = False

    def expired(self, now: float) -> bool:
        return self.active and now > self.deadline

    def start(self, callback_text: str, now: float) -> bool:
        text = str(callback_text or "").strip()
        if not text:
            self.reset()
            return False

        self.callback_head = text
        self.follow_lines = []
        self.deadline = now + self.initial_timeout
        self.active = True
        return True

    def flush(self):
        if not self.active:
            return None

        collected = CollectedSmsCallback(
            callback_text=self.callback_head,
            follow_lines=list(self.follow_lines),
        )
        self.reset()
        return collected

    def consume_line(self, line: str, now: float) -> str:
        if not self.active:
            return "pass"
        if is_sms_collection_boundary(line):
            return "boundary"
        if self._follow_line_limit_reached():
            return "flush"

        self.follow_lines.append(line)
        self.deadline = now + self.fragment_timeout
        return "consumed"

    def _follow_line_limit_reached(self) -> bool:
        if self.max_follow_lines is None:
            return False
        try:
            limit = int(self.max_follow_lines)
        except Exception:
            return False
        return limit >= 0 and len(self.follow_lines) >= limit


def callback_body_from_line(line: str, prefix: str = SMS_CALLBACK_PREFIX) -> str:
    text = str(line or "").lstrip()
    if not text.startswith(prefix):
        return ""
    return text[len(prefix):].strip()


def handle_sms_collector_line(collector, line: str, now: float, flush_callback, prefix: str = SMS_CALLBACK_PREFIX):
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
        full_msg = "".join([body or callback_head] + list(collected.follow_lines or [])).strip()
        pending = PendingSms(
            callback_head=callback_head,
            full_msg=full_msg,
            display_lines=["📩 收到短信：", callback_head] + list(collected.follow_lines or []),
        )
    if assembler is not None and not hasattr(assembler, "add_collected"):
        pending = assembler.add_message(pending, now=now, log=concat_log)
        if pending is None:
            return "pending"
    return process_pending_sms(
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
