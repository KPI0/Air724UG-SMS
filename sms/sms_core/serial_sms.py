from dataclasses import dataclass

from sms_core.serial_parsers import is_sms_collection_boundary
from sms_core.sms_processing import process_pending_sms


SMS_CALLBACK_PREFIX = "[I]-[handler_sms.smsCallback]"


@dataclass
class PendingSms:
    callback_head: str
    full_msg: str
    display_lines: list


@dataclass
class SmsLineDecision:
    started: bool = False
    action: str = "pass"
    flushed: bool = False
    continue_read: bool = False


class SmsPendingCollector:
    def __init__(
        self,
        parse_callback_head,
        correction_cache=None,
        initial_timeout: float = 1.0,
        fragment_timeout: float = 0.4,
        max_follow_lines: int = 40,
    ):
        self.parse_callback_head = parse_callback_head
        self.correction_cache = correction_cache
        self.initial_timeout = initial_timeout
        self.fragment_timeout = fragment_timeout
        self.max_follow_lines = max_follow_lines
        self.reset()

    def reset(self):
        self.parts = []
        self.display_lines = []
        self.callback_head = ""
        self.ignore_follow_lines = False
        self.deadline = 0.0
        self.follow_lines_left = 0
        self.active = False

    def expired(self, now: float) -> bool:
        return self.active and now > self.deadline

    def start(self, callback_text: str, now: float) -> bool:
        text = str(callback_text or "").strip()
        if not text:
            self.reset()
            return False

        original_text = text
        if self.correction_cache is not None:
            text = self.correction_cache.correct_callback_text(
                text,
                self.parse_callback_head,
                now,
            )
        _sender, body = self.parse_callback_head(text)
        self.callback_head = text
        self.parts = [body or text]
        self.display_lines = ["📩 收到短信：", text]
        self.ignore_follow_lines = text != original_text
        self.deadline = now + self.initial_timeout
        self.follow_lines_left = self.max_follow_lines
        self.active = True
        return True

    def flush(self):
        if not self.active:
            return None

        pending = PendingSms(
            callback_head=self.callback_head,
            full_msg="".join([part for part in self.parts if part]).strip(),
            display_lines=list(self.display_lines),
        )
        self.reset()
        return pending

    def consume_line(self, line: str, now: float) -> str:
        if not self.active:
            return "pass"
        if is_sms_collection_boundary(line):
            return "boundary"
        if self.follow_lines_left <= 0:
            return "flush"

        if not self.ignore_follow_lines:
            self.parts.append(line)
            self.display_lines.append(line)
        self.deadline = now + self.fragment_timeout
        self.follow_lines_left -= 1
        return "flush" if self.follow_lines_left <= 0 else "consumed"


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
):
    return process_pending_sms(
        collector.flush(),
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
