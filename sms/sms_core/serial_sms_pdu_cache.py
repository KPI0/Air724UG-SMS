"""
PDU cache contract:
- store decoded single-PDU bodies for corruption correction;
- store decoded concat PDU segments with UDH metadata;
- correct only single-PDU callback corruption/truncation;
- never join concat bodies, infer completion, or emit complete multipart SMS.
LongSmsAssembler is the only multipart completion state machine.
"""

import re
from dataclasses import dataclass
from datetime import datetime

from sms_core.sms_pdu import ConcatSmsInfo, decode_received_pdu


CMGR_LOG_RE = re.compile(r"\[I\]-\[lib_sms rsp\]\s+\+CMGR\b")
HEX_LINE_RE = re.compile(r"^[0-9A-Fa-f]+$")
SMS_CALLBACK_HEAD_RE = re.compile(
    r"^(?P<prefix>\s*\+?\d+\s+\d{2}/\d{2}/\d{2},\d{2}:\d{2}:\d{2}\+\d+\s*)(?P<body>.*)$",
    re.DOTALL,
)
SMS_CALLBACK_TIMESTAMP_RE = re.compile(
    r"^\s*\+?\d+\s+(?P<timestamp>\d{2}/\d{2}/\d{2},\d{2}:\d{2}:\d{2}\+\d+)"
)
DEFAULT_PDU_MAX_SEGMENT_ENTRIES = 4096
DEFAULT_PDU_MAX_COMPLETE_ENTRIES = 1024


@dataclass(frozen=True)
class CachedSmsPduSegment:
    sender: str
    body: str
    concat_info: ConcatSmsInfo
    timestamp: str = ""


def decode_incoming_sms_pdu(pdu_hex: str):
    return decode_received_pdu(pdu_hex)


class SmsPduCorrectionCache:
    def __init__(
        self,
        max_age: float = 30.0,
        multipart_timestamp_tolerance: float = 5.0,
        multipart_part_timestamp_step: float = 1.0,
        max_segment_entries: int = DEFAULT_PDU_MAX_SEGMENT_ENTRIES,
        max_complete_entries: int = DEFAULT_PDU_MAX_COMPLETE_ENTRIES,
    ):
        self.max_age = max_age
        self.multipart_timestamp_tolerance = multipart_timestamp_tolerance
        self.multipart_part_timestamp_step = multipart_part_timestamp_step
        self.max_segment_entries = max(1, int(max_segment_entries or DEFAULT_PDU_MAX_SEGMENT_ENTRIES))
        self.max_complete_entries = max(1, int(max_complete_entries or DEFAULT_PDU_MAX_COMPLETE_ENTRIES))
        self._collecting = False
        self._pdu_lines = []
        self._complete_by_key = {}
        self._segments_by_key = {}
        self._last_corrected_message = None

    def observe_line(self, line: str, now: float, log=None):
        text = str(line or "").strip()

        if self._collecting:
            if text and HEX_LINE_RE.fullmatch(text):
                self._pdu_lines.append(text)
                return
            self._finalize_pdu(now, log=log)

        if CMGR_LOG_RE.search(text):
            self._collecting = True
            self._pdu_lines = []

        self._expire(now)
        self._enforce_cache_limits(log=log)

    def correct_single_pdu_callback_text(self, callback_text: str, parse_callback_head, now: float) -> str:
        """Correct a callback body from a cached single-PDU body only."""
        self._last_corrected_message = None
        text = str(callback_text or "")
        sender, body = parse_callback_head(text)
        timestamp = _callback_timestamp(text)
        lookup_key = self._cache_key(sender, timestamp)
        candidates = self._lookup_candidates(sender, timestamp, now)

        corrected = ""
        consumed_index = None
        for index, item in enumerate(candidates):
            candidate = _cached_message_body(item)
            if _should_correct_body(body, candidate):
                corrected = candidate
                self._last_corrected_message = None
                consumed_index = index
                break

        if consumed_index is None:
            return text
        else:
            remaining = list(candidates)
            remaining.pop(consumed_index)
            if remaining:
                self._complete_by_key[lookup_key] = remaining
            else:
                self._complete_by_key.pop(lookup_key, None)

        match = SMS_CALLBACK_HEAD_RE.match(text)
        if match:
            return match.group("prefix") + corrected
        if body and text.endswith(body):
            return text[:-len(body)] + corrected
        return f"{sender} {corrected}".strip()

    def correct_callback_text(self, callback_text: str, parse_callback_head, now: float) -> str:
        """Compatibility alias; do not use for multipart reconstruction."""
        return self.correct_single_pdu_callback_text(callback_text, parse_callback_head, now)

    def concat_part_for_callback(self, callback_text: str, parse_callback_head, now: float):
        text = str(callback_text or "")
        sender, body = parse_callback_head(text)
        timestamp = _callback_timestamp(text)
        exact_match = self._match_concat_part(
            self._lookup_segments(sender, timestamp, now),
            body,
        )
        if exact_match is not None:
            return exact_match

        near_match = self._match_concat_part(
            self._lookup_near_segments(sender, timestamp, now),
            body,
            require_unique=True,
        )
        if near_match is not None:
            return near_match

        return self._match_concat_part(
            self._lookup_all_segments(sender, timestamp, now),
            body,
            require_unique=True,
        )

    def concat_info_for_callback(self, callback_text: str, parse_callback_head, now: float):
        part = self.concat_part_for_callback(callback_text, parse_callback_head, now)
        return part.concat_info if part else None

    def last_corrected_message(self):
        return self._last_corrected_message

    def segments_for_concat_part(self, part, now: float, full_msg=None):
        concat_info = getattr(part, "concat_info", None)
        return self._segments_for_concat(
            getattr(part, "sender", ""),
            concat_info,
            now,
        )

    def segments_for_message(self, message, now: float, full_msg=None):
        concat_info = ConcatSmsInfo(
            getattr(message, "reference", None),
            getattr(message, "total", None),
            1,
            getattr(message, "reference_bits", None) or 8,
        )
        return self._segments_for_concat(
            getattr(message, "sender", ""),
            concat_info,
            now,
        )

    def _segments_for_concat(self, sender, concat_info, now: float):
        if concat_info is None:
            return []
        total = int(getattr(concat_info, "total", 0) or 0)
        reference = getattr(concat_info, "reference", None)
        if total <= 1 or reference is None:
            return []

        self._expire(now)
        normalized_sender = _normalize_sender_for_match(sender)
        reference_bits = int(getattr(concat_info, "reference_bits", None) or 8)
        segments = []
        for (candidate_sender, _candidate_timestamp), items in self._segments_by_key.items():
            if candidate_sender != normalized_sender:
                continue
            for _sender, _timestamp, body, candidate_info, _seen in items:
                candidate_total = int(getattr(candidate_info, "total", 0) or 0)
                candidate_index = int(getattr(candidate_info, "index", 0) or 0)
                candidate_reference = getattr(candidate_info, "reference", None)
                candidate_bits = int(getattr(candidate_info, "reference_bits", None) or 8)
                if (
                    candidate_total != total
                    or candidate_bits != reference_bits
                    or candidate_reference is None
                    or int(candidate_reference) != int(reference)
                    or candidate_index < 1
                    or candidate_index > total
                ):
                    continue
                segments.append(
                    CachedSmsPduSegment(
                        str(_sender or sender or ""),
                        str(body or ""),
                        ConcatSmsInfo(reference, total, candidate_index, reference_bits),
                        str(_timestamp or ""),
                    )
                )
        return segments

    def _finalize_pdu(self, now: float, log=None):
        pdu_hex = "".join(self._pdu_lines)
        self._collecting = False
        self._pdu_lines = []
        decoded = decode_incoming_sms_pdu(pdu_hex)
        if decoded is None:
            return

        if decoded.total and decoded.total > 1 and decoded.index:
            self._store_segment(decoded, now)
            self._enforce_segment_limit(log=log)
            return

        if decoded.body:
            self._store_complete(decoded.sender, decoded.timestamp, decoded.body, now)
            self._enforce_complete_limit(log=log)

    def _lookup_candidates(self, sender: str, timestamp: str, now: float):
        self._expire(now)
        return list(self._complete_by_key.get(self._cache_key(sender, timestamp)) or [])

    def _lookup_segments(self, sender: str, timestamp: str, now: float):
        self._expire(now)
        return list(self._segments_by_key.get(self._cache_key(sender, timestamp)) or [])

    def _lookup_near_segments(self, sender: str, timestamp: str, now: float):
        normalized_sender = _normalize_sender_for_match(sender)
        if not normalized_sender or not timestamp:
            return []

        self._expire(now)
        matches = []
        for (candidate_sender, candidate_timestamp), items in self._segments_by_key.items():
            if candidate_sender != normalized_sender:
                continue
            if candidate_timestamp == timestamp:
                continue
            if _timestamp_matches(timestamp, candidate_timestamp, self.multipart_timestamp_tolerance):
                matches.extend(items)
        return matches

    def _lookup_all_segments(self, sender: str, timestamp: str, now: float):
        if sender and timestamp:
            return []
        self._expire(now)
        matches = []
        for items in self._segments_by_key.values():
            matches.extend(items)
        return matches

    def _store_complete(
        self,
        sender: str,
        timestamp: str,
        body: str,
        now: float,
        reference=None,
        total=None,
        reference_bits=None,
    ):
        key = self._cache_key(sender, timestamp)
        self._complete_by_key.setdefault(key, []).append((body, now))

    def _store_segment(self, decoded, now: float):
        key = self._cache_key(decoded.sender, decoded.timestamp)
        self._segments_by_key.setdefault(key, []).append((
            decoded.sender,
            decoded.timestamp,
            decoded.body,
            ConcatSmsInfo(
                decoded.reference,
                decoded.total,
                decoded.index,
                getattr(decoded, "reference_bits", None) or 8,
            ),
            now,
        ))

    def _cache_key(self, sender: str, timestamp: str):
        return _normalize_sender_for_match(sender), str(timestamp or "")

    def _expire(self, now: float):
        self._complete_by_key = {
            key: items
            for key, items in (
                (key, [item for item in items if now - _cached_message_seen(item) <= self.max_age])
                for key, items in self._complete_by_key.items()
            )
            if items
        }
        self._segments_by_key = {
            key: items
            for key, items in (
                (key, [item for item in items if now - item[4] <= self.max_age])
                for key, items in self._segments_by_key.items()
            )
            if items
        }

    def _enforce_cache_limits(self, log=None):
        self._enforce_segment_limit(log=log)
        self._enforce_complete_limit(log=log)

    def _enforce_segment_limit(self, log=None):
        while _nested_item_count(self._segments_by_key) > self.max_segment_entries:
            oldest_key, oldest_index, oldest_item = _oldest_nested_item(
                self._segments_by_key,
                seen_func=lambda item: item[4],
            )
            if oldest_key is None:
                return
            self._segments_by_key[oldest_key].pop(oldest_index)
            if not self._segments_by_key[oldest_key]:
                self._segments_by_key.pop(oldest_key, None)
            sender, timestamp, _body, concat_info, _seen = oldest_item
            _emit_cache_log(
                log,
                (
                    "[SMS] PDU CACHE EVICT "
                    f"Cache=SEGMENT Sender={sender or 'unknown'} Timestamp={timestamp or ''} "
                    f"Ref=0x{int(getattr(concat_info, 'reference', 0) or 0):X} "
                    f"Part={int(getattr(concat_info, 'index', 0) or 0)}/"
                    f"{int(getattr(concat_info, 'total', 0) or 0)}"
                ),
            )

    def _enforce_complete_limit(self, log=None):
        while _nested_item_count(self._complete_by_key) > self.max_complete_entries:
            oldest_key, oldest_index, oldest_item = _oldest_nested_item(
                self._complete_by_key,
                seen_func=_cached_message_seen,
            )
            if oldest_key is None:
                return
            self._complete_by_key[oldest_key].pop(oldest_index)
            if not self._complete_by_key[oldest_key]:
                self._complete_by_key.pop(oldest_key, None)
            _emit_cache_log(
                log,
                (
                    "[SMS] PDU CACHE EVICT "
                    f"Cache=COMPLETE Sender={oldest_key[0] or 'unknown'} Timestamp={oldest_key[1] or ''} "
                ),
            )

    def _match_concat_part(self, candidates, body, require_unique=False):
        matches = []
        for sender, timestamp, candidate, concat_info, _seen in candidates:
            if candidate == body or _matches_corrupted_or_truncated_body(body, candidate):
                matches.append(CachedSmsPduSegment(sender, candidate, concat_info, timestamp))
            elif body and len(body) >= 16 and candidate.startswith(body):
                matches.append(CachedSmsPduSegment(sender, candidate, concat_info, timestamp))
            elif "\ufffd" in str(body or "") and candidate and str(body or "").startswith(candidate):
                matches.append(CachedSmsPduSegment(sender, candidate, concat_info, timestamp))
        matches = _dedupe_concat_part_matches(matches)
        if not matches:
            return None
        if require_unique and len(matches) != 1:
            return None
        return matches[0]


def _dedupe_concat_part_matches(matches):
    unique = {}
    for match in matches:
        concat_info = getattr(match, "concat_info", None)
        key = (
            str(getattr(match, "sender", "") or ""),
            str(getattr(match, "timestamp", "") or ""),
            str(getattr(match, "body", "") or ""),
            int(getattr(concat_info, "reference_bits", None) or 8),
            int(getattr(concat_info, "reference", 0) or 0),
            int(getattr(concat_info, "total", 0) or 0),
            int(getattr(concat_info, "index", 0) or 0),
        )
        unique.setdefault(key, match)
    return list(unique.values())


def _nested_item_count(mapping):
    return sum(len(items or []) for items in dict(mapping or {}).values())


def _oldest_nested_item(mapping, *, seen_func):
    oldest_key = None
    oldest_index = None
    oldest_item = None
    oldest_seen = None
    for key, items in dict(mapping or {}).items():
        for index, item in enumerate(list(items or [])):
            try:
                seen = float(seen_func(item))
            except Exception:
                seen = 0.0
            if oldest_seen is None or seen < oldest_seen:
                oldest_key = key
                oldest_index = index
                oldest_item = item
                oldest_seen = seen
    return oldest_key, oldest_index, oldest_item


def _emit_cache_log(log, message):
    if log is None:
        return
    try:
        log(message, "normal")
    except TypeError:
        try:
            log(message)
        except Exception:
            pass
    except Exception:
        pass


def _matches_corrupted_or_truncated_body(body: str, corrected: str) -> bool:
    body_text = str(body or "")
    corrected_text = str(corrected or "")
    if "\ufffd" not in body_text:
        return False

    clean_prefix = body_text.split("\ufffd", 1)[0]
    if len(clean_prefix) >= 8:
        return corrected_text.startswith(clean_prefix)

    clean_body = body_text.replace("\ufffd", "")
    return len(clean_body) >= 16 and corrected_text.startswith(clean_body)


def _should_correct_body(body: str, candidate: str) -> bool:
    body_text = str(body or "")
    candidate_text = str(candidate or "")
    return _matches_corrupted_or_truncated_body(body_text, candidate_text) or (
        bool(body_text)
        and len(body_text) >= 16
        and body_text != candidate_text
        and candidate_text.startswith(body_text)
    )


def _cached_message_body(item) -> str:
    if isinstance(item, (list, tuple)) and item:
        return str(item[0] or "")
    return ""


def _cached_message_seen(item) -> float:
    if isinstance(item, (list, tuple)) and len(item) >= 2:
        try:
            return float(item[1])
        except Exception:
            return 0.0
    return 0.0


def _callback_timestamp(text: str) -> str:
    match = SMS_CALLBACK_TIMESTAMP_RE.match(str(text or ""))
    return match.group("timestamp") if match else ""


def _normalize_sender_for_match(sender: str) -> str:
    text = str(sender or "").strip()
    if text.startswith("+86") and len(text) > 3:
        return text[3:]
    if text.startswith("86") and len(text) > 2:
        return text[2:]
    return text


def _timestamp_matches(expected: str, candidate: str, tolerance: float) -> bool:
    expected_text = str(expected or "").strip()
    candidate_text = str(candidate or "").strip()
    if expected_text == candidate_text:
        return True
    expected_dt = _parse_sms_timestamp(expected_text)
    candidate_dt = _parse_sms_timestamp(candidate_text)
    if expected_dt is None or candidate_dt is None:
        return False
    return abs((expected_dt - candidate_dt).total_seconds()) <= float(tolerance)

def _parse_sms_timestamp(value: str):
    text = str(value or "").strip()
    try:
        return datetime.strptime(text[:17], "%y/%m/%d,%H:%M:%S")
    except Exception:
        return None
