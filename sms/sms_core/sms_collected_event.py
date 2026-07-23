"""
Collected SMS event adapter contract:
- convert a raw collector frame into PendingSms objects;
- attach decoded PDU segment metadata when available;
- never decide multipart completion or join concat segment bodies.
"""

from dataclasses import dataclass

from sms_core.serial_sms import PendingSms


@dataclass(frozen=True)
class CollectedPendingCandidates:
    segments: tuple
    fallback: PendingSms
    require_fallback_match: bool = False


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
            callback_extends_segment = _callback_extends_segment(
                body,
                getattr(concat_part, "body", ""),
            )
            require_fallback_match = _candidate_requires_fallback_match(
                concat_part,
                segment_pendings,
                callback_extends_segment=callback_extends_segment,
                callback_body=body,
                raw_lines=raw_lines,
            )
            if not raw_lines and not callback_extends_segment:
                if require_fallback_match:
                    matched_segment = _matched_segment_pending(segment_pendings, concat_part)
                    if matched_segment is not None:
                        return [matched_segment]
                return segment_pendings
            fallback_body = _callback_body_with_segment_anchors(
                body,
                raw_lines,
                segment_pendings,
            )
            return CollectedPendingCandidates(
                segments=tuple(segment_pendings),
                fallback=PendingSms(
                    callback_head=corrected_text,
                    full_msg=fallback_body,
                    display_lines=["📩 收到短信：", corrected_text] + raw_lines,
                    concat_sender=sender,
                ),
                require_fallback_match=require_fallback_match,
            )

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


def _callback_extends_segment(body, segment_body):
    body_text = str(body or "")
    segment_text = str(segment_body or "")
    return (
        bool(segment_text)
        and len(body_text) > len(segment_text)
        and body_text.startswith(segment_text)
    )


def _candidate_requires_fallback_match(
    concat_part,
    segment_pendings,
    *,
    callback_extends_segment=False,
    callback_body="",
    raw_lines=(),
):
    match_kind = str(getattr(concat_part, "match_kind", "") or "")
    if match_kind == "sender_fallback":
        return True

    callback_text = str(callback_body or "")
    segment_text = str(getattr(concat_part, "body", "") or "")
    if raw_lines:
        return True
    if callback_extends_segment and "\ufffd" not in callback_text:
        return True
    if "\ufffd" in callback_text and _corruption_fragment_count(callback_text) >= 2:
        return True

    timestamps = {
        str(getattr(pending, "concat_timestamp", "") or "")
        for pending in list(segment_pendings or [])
        if str(getattr(pending, "concat_timestamp", "") or "")
    }
    if (
        len(timestamps) > 1
        and callback_text == segment_text
    ):
        return True

    bodies_by_index = {}
    for pending in list(segment_pendings or []):
        concat_info = getattr(pending, "concat_info", None)
        index = int(getattr(concat_info, "index", 0) or 0)
        if index <= 0:
            continue
        bodies_by_index.setdefault(index, set()).add(str(getattr(pending, "concat_body", "") or ""))
    return any(len(bodies) > 1 for bodies in bodies_by_index.values())


def _corruption_fragment_count(value):
    return sum(
        1
        for fragment in str(value or "").split("\ufffd")
        if len(_body_without_line_breaks(fragment)) >= 4
    )


def _body_without_line_breaks(value):
    return str(value or "").replace("\r\n", "\n").replace("\r", "\n").replace("\n", "").strip()


def _matched_segment_pending(segment_pendings, concat_part):
    target_info = getattr(concat_part, "concat_info", None)
    target_key = (
        int(getattr(target_info, "reference_bits", None) or 8),
        int(getattr(target_info, "reference", 0) or 0),
        int(getattr(target_info, "total", 0) or 0),
        int(getattr(target_info, "index", 0) or 0),
        str(getattr(concat_part, "body", "") or ""),
        str(getattr(concat_part, "timestamp", "") or ""),
    )
    for pending in list(segment_pendings or []):
        concat_info = getattr(pending, "concat_info", None)
        candidate_key = (
            int(getattr(concat_info, "reference_bits", None) or 8),
            int(getattr(concat_info, "reference", 0) or 0),
            int(getattr(concat_info, "total", 0) or 0),
            int(getattr(concat_info, "index", 0) or 0),
            str(getattr(pending, "concat_body", "") or ""),
            str(getattr(pending, "concat_timestamp", "") or ""),
        )
        if candidate_key == target_key:
            return pending
    return None


def _callback_body_with_segment_anchors(body, raw_lines, segment_pendings):
    result = str(body or "")
    anchors = _ordered_segment_anchors(segment_pendings)
    for raw_line in list(raw_lines or []):
        line = str(raw_line or "")
        separator = _separator_from_segment_anchors(result, line, anchors)
        result += separator + line
    return result.strip()


def _ordered_segment_anchors(segment_pendings):
    anchors = []
    seen = set()
    for pending in list(segment_pendings or []):
        concat_info = getattr(pending, "concat_info", None)
        index = int(getattr(concat_info, "index", 0) or 0)
        body = str(getattr(pending, "concat_body", None) or getattr(pending, "full_msg", "") or "")
        key = (index, body)
        if not body or key in seen:
            continue
        seen.add(key)
        anchors.append(key)
    anchors.sort(key=lambda item: (item[0] <= 0, item[0], item[1]))
    return anchors


def _separator_from_segment_anchors(result, line, anchors):
    within_matches = []
    for _index, anchor in anchors:
        match = _separator_within_segment_anchor(result, line, anchor)
        if match is not None:
            within_matches.append(match)
    if within_matches:
        best_overlap = max(overlap for _separator, overlap in within_matches)
        best_separators = {
            separator
            for separator, overlap in within_matches
            if overlap == best_overlap
        }
        return "\n" if "\n" in best_separators else ""

    boundary_separator = _separator_at_segment_boundary(result, line, anchors)
    return "\n" if boundary_separator is None else boundary_separator


def _separator_within_segment_anchor(result, line, anchor):
    max_overlap = min(len(str(result or "")), len(str(anchor or "")))
    for overlap in range(max_overlap, 3, -1):
        if not str(result or "").endswith(anchor[:overlap]):
            continue
        remaining = anchor[overlap:]
        if not remaining:
            continue
        if remaining.startswith("\n"):
            if _anchor_matches_line(remaining[1:], line):
                return "\n", overlap
        elif _anchor_matches_line(remaining, line):
            return "", overlap
    return None


def _separator_at_segment_boundary(result, line, anchors):
    anchors_by_index = {}
    for index, anchor in anchors:
        anchors_by_index.setdefault(index, []).append(anchor)

    decisions = set()
    for index, anchor in anchors:
        if index <= 0:
            continue
        anchor_has_newline = anchor.endswith("\n")
        visible_anchor = anchor[:-1] if anchor_has_newline else anchor
        if visible_anchor and not str(result or "").endswith(visible_anchor):
            continue
        for next_anchor in anchors_by_index.get(index + 1, []):
            next_has_newline = next_anchor.startswith("\n")
            visible_next = next_anchor[1:] if next_has_newline else next_anchor
            if _anchor_matches_line(visible_next, line):
                decisions.add("\n" if anchor_has_newline or next_has_newline else "")

    if not decisions:
        return None
    return "\n" if "\n" in decisions else ""


def _anchor_matches_line(anchor, line):
    anchor_text = str(anchor or "")
    line_text = str(line or "")
    if not anchor_text:
        return True
    return _prefix_compatible(anchor_text, line_text) or _shared_prefix_length(anchor_text, line_text) >= 4


def _prefix_compatible(left, right):
    left_text = str(left or "")
    right_text = str(right or "")
    return left_text.startswith(right_text) or right_text.startswith(left_text)


def _shared_prefix_length(left, right):
    count = 0
    for left_char, right_char in zip(str(left or ""), str(right or "")):
        if left_char != right_char:
            break
        count += 1
    return count


def _callback_with_body(callback_text, body, parse_callback_head):
    sender, old_body = parse_callback_head(callback_text)
    text = str(callback_text or "")
    old_body = str(old_body or "")
    if old_body and text.endswith(old_body):
        return text[:-len(old_body)] + str(body or "")
    if sender:
        return (str(sender or "") + " " + str(body or "")).strip()
    return str(body or "")
