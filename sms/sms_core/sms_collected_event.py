"""
Collected SMS event adapter contract:
- convert a raw collector frame into PendingSms objects;
- attach decoded PDU segment metadata when available;
- never decide multipart completion or join concat segment bodies.
"""

from sms_core.serial_sms import PendingSms


def pending_from_collected(collected, parse_callback_head, correction_cache=None, now=None):
    if collected is None:
        return None

    original_text = str(getattr(collected, "callback_text", "") or "").strip()
    if not original_text:
        return None

    corrected_text = _correct_single_pdu_callback_text(
        correction_cache,
        original_text,
        parse_callback_head,
        now,
    )
    corrected = corrected_text != original_text
    corrected_message = None
    full_msg_override = None
    if corrected and correction_cache is not None and hasattr(correction_cache, "last_corrected_message"):
        corrected_message = correction_cache.last_corrected_message()
        full_msg_override = getattr(corrected_message, "body", None)
        segment_pendings = _segment_pendings_from_message(
            correction_cache,
            corrected_message,
            corrected_text,
            parse_callback_head,
            now,
        )
        if segment_pendings:
            return segment_pendings

    concat_part = None
    if not corrected and correction_cache is not None and hasattr(correction_cache, "concat_part_for_callback"):
        concat_part = correction_cache.concat_part_for_callback(
            corrected_text,
            parse_callback_head,
            now,
        )

    sender, body = parse_callback_head(corrected_text)
    raw_lines = list(getattr(collected, "raw_lines", []) or getattr(collected, "follow_lines", []) or [])
    if concat_part is not None:
        segment_pendings = _segment_pendings_from_part(
            correction_cache,
            concat_part,
            corrected_text,
            parse_callback_head,
            now,
        )
        if segment_pendings:
            return segment_pendings

    message_lines = [] if corrected or concat_part is not None else list(raw_lines)
    full_msg = (
        str(full_msg_override)
        if full_msg_override is not None
        else _full_message_from_callback(corrected_text, body, message_lines, concat_part)
    )
    display_lines = ["📩 收到短信：", corrected_text] + message_lines
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


def _segment_pendings_from_part(correction_cache, concat_part, callback_text, parse_callback_head, now):
    if correction_cache is None or not hasattr(correction_cache, "segments_for_concat_part"):
        return []
    segments = correction_cache.segments_for_concat_part(concat_part, now)
    return _segment_pendings(segments, callback_text, parse_callback_head)


def _segment_pendings_from_message(correction_cache, message, callback_text, parse_callback_head, now):
    if correction_cache is None or not hasattr(correction_cache, "segments_for_message"):
        return []
    segments = correction_cache.segments_for_message(message, now)
    return _segment_pendings(segments, callback_text, parse_callback_head)


def _correct_single_pdu_callback_text(correction_cache, callback_text, parse_callback_head, now):
    if correction_cache is None:
        return callback_text
    if hasattr(correction_cache, "correct_single_pdu_callback_text"):
        return correction_cache.correct_single_pdu_callback_text(
            callback_text,
            parse_callback_head,
            now,
        )
    if hasattr(correction_cache, "correct_callback_text"):
        return correction_cache.correct_callback_text(callback_text, parse_callback_head, now)
    return callback_text


def _segment_pendings(segments, callback_text, parse_callback_head):
    items = list(segments or [])
    if not items:
        return []

    sender, _body = parse_callback_head(callback_text)
    display_lines = ["📩 收到短信：", callback_text]
    pendings = []
    for segment in items:
        concat_info = getattr(segment, "concat_info", None)
        segment_body = str(getattr(segment, "body", "") or "")
        pendings.append(PendingSms(
            callback_head=_callback_with_body(callback_text, segment_body, parse_callback_head),
            full_msg=segment_body,
            display_lines=display_lines,
            concat_info=concat_info,
            concat_body=segment_body,
            concat_sender=getattr(segment, "sender", "") or sender,
            concat_timestamp=getattr(segment, "timestamp", "") or "",
            concat_reference=getattr(concat_info, "reference", None),
            concat_reference_bits=getattr(concat_info, "reference_bits", None),
            concat_total=getattr(concat_info, "total", None),
        ))
    return pendings


def _full_message_from_callback(callback_text, body, message_lines, concat_part):
    if concat_part is not None:
        return str(concat_part.body or "")
    return "\n".join([body or callback_text] + list(message_lines or [])).strip()


def _callback_with_body(callback_text, body, parse_callback_head):
    sender, old_body = parse_callback_head(callback_text)
    text = str(callback_text or "")
    old_body = str(old_body or "")
    if old_body and text.endswith(old_body):
        return text[:-len(old_body)] + str(body or "")
    if sender:
        return (str(sender or "") + " " + str(body or "")).strip()
    return str(body or "")
