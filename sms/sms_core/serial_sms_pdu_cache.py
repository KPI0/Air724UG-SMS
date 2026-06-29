import re
from dataclasses import dataclass
from datetime import datetime
from typing import Optional

from sms_core.long_sms_assembler import message_trace_id
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
DEFAULT_PDU_MAX_MULTIPART_ENTRIES = 512
DEFAULT_PDU_MAX_SEGMENT_ENTRIES = 4096
DEFAULT_PDU_MAX_COMPLETE_ENTRIES = 1024


@dataclass(frozen=True)
class DecodedIncomingSmsPdu:
    sender: str
    body: str
    timestamp: str = ""
    reference: Optional[int] = None
    total: Optional[int] = None
    index: Optional[int] = None
    reference_bits: Optional[int] = None


@dataclass(frozen=True)
class CachedSmsPduSegment:
    sender: str
    body: str
    concat_info: ConcatSmsInfo
    timestamp: str = ""


@dataclass(frozen=True)
class CachedSmsPduMessage:
    sender: str
    body: str
    timestamp: str = ""
    reference: Optional[int] = None
    total: Optional[int] = None
    reference_bits: Optional[int] = None
    message_trace_id: str = ""


def _decode_semi_octet_number(value: str, number_type: str, digit_count: int) -> str:
    digits = []
    for i in range(0, len(value), 2):
        pair = value[i:i + 2]
        if len(pair) == 2:
            digits.append(pair[1])
            digits.append(pair[0])
    number = "".join(digits)[:digit_count].rstrip("Ff")
    return ("+" if number_type == "91" else "") + number


def _read_octet(hex_text: str, offset: int):
    if offset + 2 > len(hex_text):
        raise ValueError("truncated PDU")
    return int(hex_text[offset:offset + 2], 16), offset + 2


def _decode_timestamp(scts_hex: str) -> str:
    if len(scts_hex) < 14:
        return ""
    fields = []
    for i in range(0, 14, 2):
        pair = scts_hex[i:i + 2]
        fields.append(pair[1] + pair[0])
    return f"{fields[0]}/{fields[1]}/{fields[2]},{fields[3]}:{fields[4]}:{fields[5]}+{fields[6]}"


def _parse_concat_udh(udh: bytes):
    pos = 0
    while pos + 2 <= len(udh):
        iei = udh[pos]
        iedl = udh[pos + 1]
        pos += 2
        content = udh[pos:pos + iedl]
        pos += iedl
        if len(content) != iedl:
            break
        if iei == 0x00 and iedl == 3:
            return content[0], content[1], content[2]
        if iei == 0x08 and iedl == 4:
            return (content[0] << 8) | content[1], content[2], content[3]
    return None, None, None


def decode_incoming_sms_pdu(pdu_hex: str):
    return decode_received_pdu(pdu_hex)


class SmsPduCorrectionCache:
    def __init__(
        self,
        max_age: float = 30.0,
        multipart_timestamp_tolerance: float = 5.0,
        multipart_part_timestamp_step: float = 1.0,
        max_multipart_entries: int = DEFAULT_PDU_MAX_MULTIPART_ENTRIES,
        max_segment_entries: int = DEFAULT_PDU_MAX_SEGMENT_ENTRIES,
        max_complete_entries: int = DEFAULT_PDU_MAX_COMPLETE_ENTRIES,
    ):
        self.max_age = max_age
        self.multipart_timestamp_tolerance = multipart_timestamp_tolerance
        self.multipart_part_timestamp_step = multipart_part_timestamp_step
        self.max_multipart_entries = max(1, int(max_multipart_entries or DEFAULT_PDU_MAX_MULTIPART_ENTRIES))
        self.max_segment_entries = max(1, int(max_segment_entries or DEFAULT_PDU_MAX_SEGMENT_ENTRIES))
        self.max_complete_entries = max(1, int(max_complete_entries or DEFAULT_PDU_MAX_COMPLETE_ENTRIES))
        self._collecting = False
        self._pdu_lines = []
        self._multipart = {}
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

    def correct_callback_text(self, callback_text: str, parse_callback_head, now: float) -> str:
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
                self._last_corrected_message = _cached_message_metadata(item)
                consumed_index = index
                break

        if consumed_index is None:
            for item in self._lookup_assembled_candidates(sender, timestamp, now):
                candidate = _cached_message_body(item)
                if _should_correct_body(body, candidate):
                    corrected = candidate
                    self._last_corrected_message = _cached_message_metadata(item)
                    break
            if not corrected:
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

    def complete_metadata_for_concat_part(self, part, now: float, full_msg=None):
        concat_info = getattr(part, "concat_info", None)
        if concat_info is None:
            return None
        total = int(getattr(concat_info, "total", 0) or 0)
        reference = getattr(concat_info, "reference", None)
        if total <= 1 or reference is None:
            return None

        self._expire(now)
        normalized_sender = _normalize_sender_for_match(getattr(part, "sender", ""))
        reference_bits = int(getattr(concat_info, "reference_bits", None) or 8)
        parts = {}
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
                parts.setdefault(candidate_index, set()).add(str(body or ""))

        if any(index not in parts or len(parts[index]) != 1 for index in range(1, total + 1)):
            return None
        body = "".join(next(iter(parts[index])) for index in range(1, total + 1))
        if full_msg is not None and _normalize_message_text(full_msg) != _normalize_message_text(body):
            return None

        return _build_cached_message(
            getattr(part, "sender", ""),
            getattr(part, "timestamp", ""),
            body,
            reference,
            total,
            reference_bits,
        )

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
            reference_bits = int(getattr(decoded, "reference_bits", None) or 8)
            key = (
                decoded.sender,
                decoded.timestamp,
                reference_bits,
                decoded.reference,
                decoded.total,
            )
            entry = self._multipart.setdefault(
                key,
                {
                    "sender": decoded.sender,
                    "timestamp": decoded.timestamp,
                    "reference_bits": reference_bits,
                    "reference": decoded.reference,
                    "total": decoded.total,
                    "parts": {},
                    "seen": now,
                },
            )
            entry["seen"] = now
            entry["parts"][decoded.index] = decoded.body
            self._enforce_multipart_limit(log=log)
            if len(entry["parts"]) == decoded.total:
                body = "".join(entry["parts"].get(i, "") for i in range(1, decoded.total + 1))
                self._store_complete(
                    decoded.sender,
                    decoded.timestamp,
                    body,
                    now,
                    reference=decoded.reference,
                    total=decoded.total,
                    reference_bits=reference_bits,
                )
                self._multipart.pop(key, None)
                self._enforce_complete_limit(log=log)
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

    def _lookup_assembled_candidates(self, sender: str, timestamp: str, now: float):
        normalized_sender = _normalize_sender_for_match(sender)
        if not normalized_sender or not timestamp:
            return []

        self._expire(now)
        grouped = {}
        for (candidate_sender, candidate_timestamp), items in self._segments_by_key.items():
            if candidate_sender != normalized_sender:
                continue
            if not _timestamp_matches(timestamp, candidate_timestamp, self.multipart_timestamp_tolerance):
                continue
            for _sender, segment_timestamp, body, concat_info, seen in items:
                total = int(getattr(concat_info, "total", 0) or 0)
                index = int(getattr(concat_info, "index", 0) or 0)
                reference = getattr(concat_info, "reference", None)
                if total <= 1 or index < 1 or index > total or reference is None:
                    continue
                group_key = (
                    int(getattr(concat_info, "reference_bits", None) or 8),
                    int(reference),
                    total,
                )
                entry = grouped.setdefault(group_key, {
                    "sender": _sender or candidate_sender,
                    "timestamp": timestamp,
                    "reference_bits": group_key[0],
                    "reference": group_key[1],
                    "total": total,
                    "parts": {},
                })
                entry["parts"].setdefault(index, []).append((str(body or ""), seen, segment_timestamp))

        candidates = []
        for entry in grouped.values():
            total = int(entry.get("total") or 0)
            parts = entry.get("parts") or {}
            if total <= 1 or any(i not in parts for i in range(1, total + 1)):
                continue
            body_parts = []
            seen_values = []
            segment_timestamps = []
            ambiguous = False
            for i in range(1, total + 1):
                unique_bodies = {part_body for part_body, _seen, _timestamp in parts[i]}
                unique_timestamps = {_timestamp for _part_body, _seen, _timestamp in parts[i]}
                if len(unique_bodies) != 1 or len(unique_timestamps) != 1:
                    ambiguous = True
                    break
                body_parts.append(next(iter(unique_bodies)))
                segment_timestamps.append(next(iter(unique_timestamps)))
                seen_values.extend(_seen for _part_body, _seen, _timestamp in parts[i])
            if not ambiguous and _multipart_timestamps_have_single_anchor(segment_timestamps):
                body = "".join(body_parts)
                seen = max(seen_values or [now])
                message = _build_cached_message(
                    entry.get("sender") or normalized_sender,
                    timestamp,
                    body,
                    entry.get("reference"),
                    total,
                    entry.get("reference_bits"),
                )
                candidates.append((body, seen, message))
        return candidates

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
        message = _build_cached_message(
            sender,
            timestamp,
            body,
            reference,
            total,
            reference_bits,
        )
        item = (body, now, message) if message else (body, now)
        self._complete_by_key.setdefault(key, []).append(item)

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
        self._multipart = {
            key: entry
            for key, entry in self._multipart.items()
            if now - entry.get("seen", now) <= self.max_age
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
        self._enforce_multipart_limit(log=log)
        self._enforce_segment_limit(log=log)
        self._enforce_complete_limit(log=log)

    def _enforce_multipart_limit(self, log=None):
        while len(self._multipart) > self.max_multipart_entries:
            oldest_key = min(
                self._multipart,
                key=lambda key: self._multipart[key].get("seen", 0.0),
            )
            entry = self._multipart.pop(oldest_key, None)
            if entry is not None:
                _emit_cache_log(
                    log,
                    (
                        "[SMS] PDU CACHE EVICT "
                        f"Cache=MULTIPART Sender={entry.get('sender') or 'unknown'} "
                        f"Ref=0x{int(entry.get('reference') or 0):X} "
                        f"Parts={len(entry.get('parts') or {})}/{int(entry.get('total') or 0)}"
                    ),
                )

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
            metadata = _cached_message_metadata(oldest_item)
            reference = getattr(metadata, "reference", None) if metadata else None
            total = getattr(metadata, "total", None) if metadata else None
            _emit_cache_log(
                log,
                (
                    "[SMS] PDU CACHE EVICT "
                    f"Cache=COMPLETE Sender={oldest_key[0] or 'unknown'} Timestamp={oldest_key[1] or ''} "
                    f"Ref=0x{int(reference or 0):X} Total={int(total or 0)}"
                ),
            )

    def _match_concat_part(self, candidates, body, require_unique=False):
        matches = []
        for sender, timestamp, candidate, concat_info, _seen in candidates:
            if candidate == body or _matches_corrupted_or_truncated_body(body, candidate):
                matches.append(CachedSmsPduSegment(sender, candidate, concat_info, timestamp))
            elif body and len(body) >= 16 and candidate.startswith(body):
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


def _build_cached_message(
    sender: str,
    timestamp: str,
    body: str,
    reference=None,
    total=None,
    reference_bits=None,
):
    if reference is None or total is None:
        return None
    try:
        reference_value = int(reference)
        total_value = int(total)
        reference_bits_value = int(reference_bits or 8)
    except Exception:
        return None
    if total_value <= 1:
        return None

    normalized_sender = _normalize_sender_for_match(sender)
    trace_id = message_trace_id(
        normalized_sender,
        timestamp,
        reference_bits_value,
        reference_value,
        total_value,
    )
    return CachedSmsPduMessage(
        sender=str(sender or ""),
        body=str(body or ""),
        timestamp=str(timestamp or ""),
        reference=reference_value,
        total=total_value,
        reference_bits=reference_bits_value,
        message_trace_id=trace_id,
    )


def _cached_message_body(item) -> str:
    if isinstance(item, CachedSmsPduMessage):
        return str(item.body or "")
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


def _cached_message_metadata(item):
    if isinstance(item, CachedSmsPduMessage):
        return item
    if isinstance(item, (list, tuple)) and len(item) >= 3:
        meta = item[2]
        if isinstance(meta, CachedSmsPduMessage):
            return meta
    return None


def _normalize_message_text(value) -> str:
    return re.sub(r"[\r\n]+", "", str(value or ""))


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


def _multipart_timestamps_have_single_anchor(timestamps) -> bool:
    normalized = {str(timestamp or "").strip() for timestamp in list(timestamps or [])}
    normalized.discard("")
    return len(normalized) == 1


def _parse_sms_timestamp(value: str):
    text = str(value or "").strip()
    try:
        return datetime.strptime(text[:17], "%y/%m/%d,%H:%M:%S")
    except Exception:
        return None
