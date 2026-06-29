from sms_core.serial_sms import PendingSms


class SmsReceivePipeline:
    def __init__(self, parse_callback_head, correction_cache, long_sms_assembler):
        self.parse_callback_head = parse_callback_head
        self.correction_cache = correction_cache
        self.long_sms_assembler = long_sms_assembler

    def observe_line(self, line: str, now: float, log=None):
        try:
            self.correction_cache.observe_line(line, now, log=log)
        except TypeError:
            self.correction_cache.observe_line(line, now)

    def add_collected(self, collected, now=None, log=None):
        pending = self._pending_from_collected(collected, now)
        pending = self.long_sms_assembler.add_message(pending, now=now, log=log)
        return pending

    def reset(self):
        self.long_sms_assembler.reset()

    def _pending_from_collected(self, collected, now):
        if collected is None:
            return None

        original_text = str(getattr(collected, "callback_text", "") or "").strip()
        if not original_text:
            return None

        corrected_text = self.correction_cache.correct_callback_text(
            original_text,
            self.parse_callback_head,
            now,
        )
        corrected = corrected_text != original_text
        corrected_message = None
        full_msg_override = None
        if corrected and hasattr(self.correction_cache, "last_corrected_message"):
            corrected_message = self.correction_cache.last_corrected_message()
            full_msg_override = getattr(corrected_message, "body", None)

        concat_part = None
        if not corrected and hasattr(self.correction_cache, "concat_part_for_callback"):
            concat_part = self.correction_cache.concat_part_for_callback(
                corrected_text,
                self.parse_callback_head,
                now,
            )

        sender, body = self.parse_callback_head(corrected_text)
        original_follow_lines = list(getattr(collected, "follow_lines", []) or [])
        if concat_part is not None and original_follow_lines:
            merged_full_msg = self._full_message(corrected_text, body, original_follow_lines, None)
            if self._looks_like_merged_concat_callback(merged_full_msg, concat_part):
                merged_metadata = self._complete_metadata_for_concat_part(
                    concat_part,
                    now,
                    merged_full_msg,
                )
                if merged_metadata is not None:
                    corrected_message = merged_metadata
                    full_msg_override = getattr(merged_metadata, "body", None)
                    concat_part = None
                else:
                    # The Lua callback already contains more than the matched
                    # first PDU segment. Prefer showing that callback over
                    # holding the SMS indefinitely when PDU metadata is stale
                    # or carrier timestamps do not line up.
                    concat_part = None

        follow_lines = list(original_follow_lines)
        if corrected or concat_part is not None:
            follow_lines = []

        full_msg = (
            str(full_msg_override)
            if full_msg_override is not None
            else self._full_message(corrected_text, body, follow_lines, concat_part)
        )
        display_lines = ["📩 收到短信：", corrected_text] + follow_lines
        if corrected_message is not None:
            concat_reference = getattr(corrected_message, "reference", None)
            concat_reference_bits = getattr(corrected_message, "reference_bits", None)
            concat_total = getattr(corrected_message, "total", None)
            trace_id = getattr(corrected_message, "message_trace_id", None)
        elif concat_part and concat_part.concat_info:
            concat_reference = getattr(concat_part.concat_info, "reference", None)
            concat_reference_bits = getattr(concat_part.concat_info, "reference_bits", None)
            concat_total = getattr(concat_part.concat_info, "total", None)
            trace_id = None
        else:
            concat_reference = None
            concat_reference_bits = None
            concat_total = None
            trace_id = None

        return PendingSms(
            callback_head=corrected_text,
            full_msg=full_msg,
            display_lines=display_lines,
            concat_info=concat_part.concat_info if concat_part else None,
            concat_body=concat_part.body if concat_part else None,
            concat_sender=concat_part.sender if concat_part else sender,
            concat_timestamp=getattr(concat_part, "timestamp", "") if concat_part else "",
            concat_reference=concat_reference,
            concat_reference_bits=concat_reference_bits,
            concat_total=concat_total,
            message_trace_id=trace_id,
        )

    def _full_message(self, callback_text, body, follow_lines, concat_part):
        if concat_part is not None:
            return str(concat_part.body or "")
        return "\n".join([body or callback_text] + follow_lines).strip()

    def _looks_like_merged_concat_callback(self, full_msg, concat_part):
        full_text = _normalize_callback_text(full_msg)
        part_body = _normalize_callback_text(getattr(concat_part, "body", ""))
        if not full_text or not part_body:
            return False
        if len(full_text) <= len(part_body):
            return False
        return full_text.startswith(part_body)

    def _complete_metadata_for_concat_part(self, concat_part, now, full_msg):
        if not hasattr(self.correction_cache, "complete_metadata_for_concat_part"):
            return None
        try:
            return self.correction_cache.complete_metadata_for_concat_part(
                concat_part,
                now,
                full_msg=full_msg,
            )
        except TypeError:
            return self.correction_cache.complete_metadata_for_concat_part(concat_part, now)


def _normalize_callback_text(value):
    return str(value or "").replace("\r", "").replace("\n", "")
