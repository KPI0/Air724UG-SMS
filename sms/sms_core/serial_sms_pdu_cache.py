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
    r"^(?P<prefix>\s*[^\r\n]+?\s+"
    r"\d{2}/\d{2}/\d{2},\d{2}:\d{2}:\d{2}\+\d+\s*)(?P<body>.*)$",
    re.DOTALL,
)
SMS_CALLBACK_TIMESTAMP_RE = re.compile(
    r"^\s*[^\r\n]+?\s+(?P<timestamp>\d{2}/\d{2}/\d{2},\d{2}:\d{2}:\d{2}\+\d+)"
)
DEFAULT_PDU_MAX_SEGMENT_ENTRIES = 4096
DEFAULT_PDU_MAX_COMPLETE_ENTRIES = 1024


@dataclass(frozen=True)
class CachedSmsPduSegment:
    sender: str
    body: str
    concat_info: ConcatSmsInfo
    timestamp: str = ""
    match_kind: str = ""
    sender_is_alphanumeric: bool = False
    sender_legacy_alias: str = ""


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

    def reset(self):
        self._collecting = False
        self._pdu_lines.clear()
        self._complete_by_key.clear()
        self._segments_by_key.clear()
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
        """Correct a callback body and legacy sender from one cached PDU."""
        self._last_corrected_message = None
        text = str(callback_text or "")
        sender, body = parse_callback_head(text)
        timestamp = _callback_timestamp(text)
        lookup_key, candidates = self._lookup_candidates(sender, timestamp, now)

        exact_index = next(
            (
                index
                for index, item in enumerate(candidates)
                if _cached_message_body(item) == str(body or "")
            ),
            None,
        )
        if exact_index is not None:
            self._consume_complete_candidate(lookup_key, candidates, exact_index)
            return text

        correction_matches = []
        for index, item in enumerate(candidates):
            candidate = _cached_message_body(item)
            if _should_correct_body(body, candidate):
                correction_matches.append((index, candidate))

        if not correction_matches:
            legacy_matches = []
            legacy_candidates = self._lookup_legacy_complete_candidates(sender, timestamp, now)
            for legacy_key, legacy_index, legacy_item in legacy_candidates:
                candidate = _cached_message_body(legacy_item)
                if candidate == str(body or "") or _should_correct_body(body, candidate):
                    legacy_matches.append((legacy_key, legacy_index, legacy_item, candidate))
            unique_matches = {
                (_cached_message_sender(item), candidate)
                for _key, _index, item, candidate in legacy_matches
            }
            if len(unique_matches) == 1 and legacy_matches:
                legacy_key, legacy_index, legacy_item, corrected = legacy_matches[0]
                legacy_candidates = list(self._complete_by_key.get(legacy_key) or [])
                self._consume_complete_candidate(legacy_key, legacy_candidates, legacy_index)
                sender = _cached_message_sender(legacy_item) or sender
                return _callback_with_sender_and_body(text, sender, corrected, parse_callback_head)

        corrected_bodies = {candidate for _index, candidate in correction_matches}
        if len(corrected_bodies) != 1:
            return text

        consumed_index, corrected = correction_matches[0]
        self._consume_complete_candidate(lookup_key, candidates, consumed_index)

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
            match_kind="exact",
        )
        if exact_match is not None:
            return exact_match

        near_match = self._match_concat_part(
            self._lookup_near_segments(sender, timestamp, now),
            body,
            require_unique=True,
            match_kind="near",
        )
        if near_match is not None:
            return near_match

        legacy_sender_candidates = self._lookup_legacy_alpha_segments(sender, timestamp, now)
        return self._match_concat_part(
            self._lookup_all_segments(sender, timestamp, now),
            body,
            require_unique=True,
            match_kind="sender_fallback",
        ) or self._match_concat_part(
            legacy_sender_candidates,
            body,
            require_unique=True,
            match_kind="legacy_sender",
        )

    def correct_concat_callback_sender(self, callback_text: str, parse_callback_head, part) -> str:
        if str(getattr(part, "match_kind", "") or "") != "legacy_sender":
            return callback_text
        sender = str(getattr(part, "sender", "") or "")
        if not sender:
            return callback_text
        _old_sender, body = parse_callback_head(callback_text)
        return _callback_with_sender_and_body(callback_text, sender, body, parse_callback_head)

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
        reference_bits = int(getattr(concat_info, "reference_bits", None) or 8)
        for sender_key in _sender_key_candidates(sender):
            segments = []
            for (candidate_sender, _candidate_timestamp), items in self._segments_by_key.items():
                if candidate_sender != sender_key:
                    continue
                for _sender, _timestamp, body, candidate_info, _seen, _alpha, _legacy_alias in items:
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
                            sender_is_alphanumeric=bool(_alpha),
                            sender_legacy_alias=str(_legacy_alias or ""),
                        )
                    )
            if segments:
                return segments
        return []

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
            self._store_complete(
                decoded.sender,
                decoded.timestamp,
                decoded.body,
                now,
                sender_is_alphanumeric=bool(getattr(decoded, "sender_is_alphanumeric", False)),
                sender_legacy_alias=str(getattr(decoded, "sender_legacy_alias", "") or ""),
            )
            self._enforce_complete_limit(log=log)

    def _lookup_candidates(self, sender: str, timestamp: str, now: float):
        self._expire(now)
        for sender_key in _sender_key_candidates(sender):
            key = sender_key, str(timestamp or "")
            items = list(self._complete_by_key.get(key) or [])
            if items:
                return key, items
        return self._cache_key(sender, timestamp), []

    def _lookup_legacy_complete_candidates(self, sender: str, timestamp: str, now: float):
        self._expire(now)
        if not timestamp:
            return []
        matches = []
        for key, items in self._complete_by_key.items():
            for index, item in enumerate(items):
                if not _cached_message_is_alphanumeric(item):
                    continue
                if not _legacy_sender_alias_matches(sender, _cached_message_legacy_alias(item)):
                    continue
                if _timestamp_matches(timestamp, key[1], self.multipart_timestamp_tolerance):
                    matches.append((key, index, item))
        return matches

    def _consume_complete_candidate(self, lookup_key, candidates, consumed_index):
        remaining = list(candidates or [])
        remaining.pop(consumed_index)
        if remaining:
            self._complete_by_key[lookup_key] = remaining
        else:
            self._complete_by_key.pop(lookup_key, None)

    def _lookup_segments(self, sender: str, timestamp: str, now: float):
        self._expire(now)
        for sender_key in _sender_key_candidates(sender):
            items = list(self._segments_by_key.get((sender_key, str(timestamp or ""))) or [])
            if items:
                return items
        return []

    def _lookup_near_segments(self, sender: str, timestamp: str, now: float):
        sender_keys = _sender_key_candidates(sender)
        if not sender_keys or not timestamp:
            return []

        self._expire(now)
        for sender_key in sender_keys:
            matches = []
            for (candidate_sender, candidate_timestamp), items in self._segments_by_key.items():
                if candidate_sender != sender_key:
                    continue
                if candidate_timestamp == timestamp:
                    continue
                if _timestamp_matches(timestamp, candidate_timestamp, self.multipart_timestamp_tolerance):
                    matches.extend(items)
            if matches:
                return matches
        return []

    def _lookup_all_segments(self, sender: str, timestamp: str, now: float):
        self._expire(now)
        sender_keys = _sender_key_candidates(sender)
        if not sender_keys:
            return [
                item
                for items in self._segments_by_key.values()
                for item in items
            ]
        for sender_key in sender_keys:
            matches = []
            for (candidate_sender, _candidate_timestamp), items in self._segments_by_key.items():
                if candidate_sender == sender_key:
                    matches.extend(items)
            if matches:
                return matches
        return []

    def _lookup_legacy_alpha_segments(self, sender: str, timestamp: str, now: float):
        self._expire(now)
        if not timestamp:
            return []
        matches = []
        for _key, items in self._segments_by_key.items():
            for item in items:
                if not _segment_is_alphanumeric(item):
                    continue
                if not _legacy_sender_alias_matches(sender, _segment_legacy_alias(item)):
                    continue
                if isinstance(item, (list, tuple)) and len(item) >= 2 and _timestamp_matches(
                    timestamp,
                    item[1],
                    self.multipart_timestamp_tolerance,
                ):
                    matches.append(item)
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
        sender_is_alphanumeric=False,
        sender_legacy_alias="",
    ):
        key = self._cache_key(
            sender,
            timestamp,
            sender_is_alphanumeric=bool(sender_is_alphanumeric),
        )
        self._complete_by_key.setdefault(key, []).append((
            body,
            now,
            sender,
            bool(sender_is_alphanumeric),
            str(sender_legacy_alias or ""),
        ))

    def _store_segment(self, decoded, now: float):
        sender_is_alphanumeric = bool(getattr(decoded, "sender_is_alphanumeric", False))
        key = self._cache_key(
            decoded.sender,
            decoded.timestamp,
            sender_is_alphanumeric=sender_is_alphanumeric,
        )
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
            sender_is_alphanumeric,
            str(getattr(decoded, "sender_legacy_alias", "") or ""),
        ))

    def _cache_key(self, sender: str, timestamp: str, *, sender_is_alphanumeric=False):
        sender_key = (
            str(sender or "").strip()
            if sender_is_alphanumeric
            else _normalize_sender_for_match(sender)
        )
        return sender_key, str(timestamp or "")

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
                (key, [item for item in items if now - _segment_seen(item) <= self.max_age])
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
                seen_func=_segment_seen,
            )
            if oldest_key is None:
                return
            self._segments_by_key[oldest_key].pop(oldest_index)
            if not self._segments_by_key[oldest_key]:
                self._segments_by_key.pop(oldest_key, None)
            sender, timestamp, _body, concat_info, _seen, _alpha, _legacy_alias = oldest_item
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

    def _match_concat_part(self, candidates, body, require_unique=False, match_kind=""):
        matches = []
        for sender, timestamp, candidate, concat_info, _seen, sender_is_alphanumeric, sender_legacy_alias in candidates:
            if candidate == body or _matches_corrupted_or_truncated_body(body, candidate):
                matches.append(CachedSmsPduSegment(
                    sender,
                    candidate,
                    concat_info,
                    timestamp,
                    match_kind,
                    bool(sender_is_alphanumeric),
                    str(sender_legacy_alias or ""),
                ))
            elif body and len(body) >= 16 and candidate.startswith(body):
                matches.append(CachedSmsPduSegment(sender, candidate, concat_info, timestamp, match_kind, bool(sender_is_alphanumeric), str(sender_legacy_alias or "")))
            elif _matches_merged_callback_prefix(body, candidate, concat_info):
                matches.append(CachedSmsPduSegment(sender, candidate, concat_info, timestamp, match_kind, bool(sender_is_alphanumeric), str(sender_legacy_alias or "")))
            elif "\ufffd" in str(body or "") and candidate and str(body or "").startswith(candidate):
                matches.append(CachedSmsPduSegment(sender, candidate, concat_info, timestamp, match_kind, bool(sender_is_alphanumeric), str(sender_legacy_alias or "")))
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


def _matches_merged_callback_prefix(body: str, candidate: str, concat_info) -> bool:
    body_text = str(body or "")
    candidate_text = str(candidate or "")
    candidate_index = int(getattr(concat_info, "index", 0) or 0)
    return (
        candidate_index == 1
        and len(candidate_text) >= 16
        and len(body_text) > len(candidate_text)
        and body_text.startswith(candidate_text)
    )


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


def _cached_message_sender(item) -> str:
    if isinstance(item, (list, tuple)) and len(item) >= 3:
        return str(item[2] or "")
    return ""


def _cached_message_is_alphanumeric(item) -> bool:
    if isinstance(item, (list, tuple)) and len(item) >= 4:
        return bool(item[3])
    return False


def _cached_message_legacy_alias(item) -> str:
    if isinstance(item, (list, tuple)) and len(item) >= 5:
        return str(item[4] or "")
    return ""


def _segment_seen(item) -> float:
    if isinstance(item, (list, tuple)) and len(item) >= 5:
        try:
            return float(item[4])
        except Exception:
            return 0.0
    return 0.0


def _segment_is_alphanumeric(item) -> bool:
    return isinstance(item, (list, tuple)) and len(item) >= 6 and bool(item[5])


def _segment_legacy_alias(item) -> str:
    if isinstance(item, (list, tuple)) and len(item) >= 7:
        return str(item[6] or "")
    return ""


def _legacy_sender_alias_matches(sender: str, alias: str) -> bool:
    return bool(alias) and str(sender or "").strip().upper() == str(alias).strip().upper()


def _callback_with_sender_and_body(callback_text, sender, body, parse_callback_head):
    text = str(callback_text or "")
    _old_sender, old_body = parse_callback_head(text)
    match = SMS_CALLBACK_HEAD_RE.match(text)
    if match:
        prefix = match.group("prefix")
        timestamp_match = re.search(r"\d{2}/\d{2}/\d{2},\d{2}:\d{2}:\d{2}\+\d+", prefix)
        if timestamp_match:
            timestamp = timestamp_match.group(0)
            spacing = prefix[timestamp_match.end():]
            return f"{sender} {timestamp}{spacing}{str(body or '')}"
    if old_body and text.endswith(old_body):
        return f"{sender} {str(body or '')}".strip()
    return f"{sender} {str(body or '')}".strip()


def _callback_timestamp(text: str) -> str:
    match = SMS_CALLBACK_TIMESTAMP_RE.match(str(text or ""))
    return match.group("timestamp") if match else ""


def _normalize_sender_for_match(sender: str) -> str:
    text = str(sender or "").strip()
    if text.startswith("+86") and len(text) > 3 and text[3:].isdigit():
        return text[3:]
    if text.startswith("86") and len(text) > 2 and text[2:].isdigit():
        return text[2:]
    return text


def _sender_key_candidates(sender: str):
    exact = str(sender or "").strip()
    normalized = _normalize_sender_for_match(exact)
    return list(dict.fromkeys(key for key in (exact, normalized) if key))


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
