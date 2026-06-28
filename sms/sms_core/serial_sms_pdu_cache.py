import re
from dataclasses import dataclass
from typing import Optional


CMGR_LOG_RE = re.compile(r"\[I\]-\[lib_sms rsp\]\s+\+CMGR\b")
HEX_LINE_RE = re.compile(r"^[0-9A-Fa-f]+$")
SMS_CALLBACK_HEAD_RE = re.compile(
    r"^(?P<prefix>\s*\+?\d+\s+\d{2}/\d{2}/\d{2},\d{2}:\d{2}:\d{2}\+\d+\s*)(?P<body>.*)$",
    re.DOTALL,
)
SMS_CALLBACK_TIMESTAMP_RE = re.compile(
    r"^\s*\+?\d+\s+(?P<timestamp>\d{2}/\d{2}/\d{2},\d{2}:\d{2}:\d{2}\+\d+)"
)


@dataclass(frozen=True)
class DecodedIncomingSmsPdu:
    sender: str
    body: str
    timestamp: str = ""
    reference: Optional[int] = None
    total: Optional[int] = None
    index: Optional[int] = None


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
    pdu = re.sub(r"\s+", "", str(pdu_hex or ""))
    if not pdu or len(pdu) % 2:
        return None

    try:
        smsc_len, offset = _read_octet(pdu, 0)
        offset += smsc_len * 2
        first_octet, offset = _read_octet(pdu, offset)
        sender_digits, offset = _read_octet(pdu, offset)
        number_type = pdu[offset:offset + 2]
        offset += 2
        sender_hex_len = ((sender_digits + 1) // 2) * 2
        sender = _decode_semi_octet_number(
            pdu[offset:offset + sender_hex_len],
            number_type,
            sender_digits,
        )
        offset += sender_hex_len

        _pid, offset = _read_octet(pdu, offset)
        dcs, offset = _read_octet(pdu, offset)
        timestamp = _decode_timestamp(pdu[offset:offset + 14])
        offset += 14
        user_data_len, offset = _read_octet(pdu, offset)
        user_data = pdu[offset:offset + user_data_len * 2]
        if len(user_data) < user_data_len * 2:
            return None

        reference = total = index = None
        payload = user_data
        if first_octet & 0x40:
            udhl, _udh_offset = _read_octet(user_data, 0)
            udh_hex_len = udhl * 2
            udh = bytes.fromhex(user_data[2:2 + udh_hex_len])
            payload = user_data[2 + udh_hex_len:]
            reference, total, index = _parse_concat_udh(udh)

        if dcs != 0x08:
            return None
        body = bytes.fromhex(payload).decode("utf-16-be")
        return DecodedIncomingSmsPdu(sender, body, timestamp, reference, total, index)
    except Exception:
        return None


class SmsPduCorrectionCache:
    def __init__(self, max_age: float = 30.0):
        self.max_age = max_age
        self._collecting = False
        self._pdu_lines = []
        self._multipart = {}
        self._complete_by_key = {}

    def observe_line(self, line: str, now: float):
        text = str(line or "").strip()

        if self._collecting:
            if text and HEX_LINE_RE.fullmatch(text):
                self._pdu_lines.append(text)
                return
            self._finalize_pdu(now)

        if CMGR_LOG_RE.search(text):
            self._collecting = True
            self._pdu_lines = []

        self._expire(now)

    def correct_callback_text(self, callback_text: str, parse_callback_head, now: float) -> str:
        text = str(callback_text or "")
        sender, body = parse_callback_head(text)
        timestamp = _callback_timestamp(text)
        lookup_key = self._cache_key(sender, timestamp)
        candidates = self._lookup_candidates(sender, timestamp, now)
        if not candidates:
            return text

        corrected = ""
        consumed_index = None
        for index, (candidate, _seen) in enumerate(candidates):
            should_correct = _matches_corrupted_or_truncated_body(body, candidate) or (
                bool(body)
                and len(body) >= 16
                and body != candidate
                and candidate.startswith(body)
            )
            if should_correct:
                corrected = candidate
                consumed_index = index
                break

        if consumed_index is None:
            return text

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

    def _finalize_pdu(self, now: float):
        pdu_hex = "".join(self._pdu_lines)
        self._collecting = False
        self._pdu_lines = []
        decoded = decode_incoming_sms_pdu(pdu_hex)
        if decoded is None:
            return

        if decoded.total and decoded.total > 1 and decoded.index:
            key = (decoded.sender, decoded.timestamp, decoded.reference, decoded.total)
            entry = self._multipart.setdefault(
                key,
                {
                    "sender": decoded.sender,
                    "timestamp": decoded.timestamp,
                    "total": decoded.total,
                    "parts": {},
                    "seen": now,
                },
            )
            entry["seen"] = now
            entry["parts"][decoded.index] = decoded.body
            if len(entry["parts"]) == decoded.total:
                body = "".join(entry["parts"].get(i, "") for i in range(1, decoded.total + 1))
                self._store_complete(decoded.sender, decoded.timestamp, body, now)
                self._multipart.pop(key, None)
            return

        if decoded.body:
            self._store_complete(decoded.sender, decoded.timestamp, decoded.body, now)

    def _lookup_candidates(self, sender: str, timestamp: str, now: float):
        self._expire(now)
        return list(self._complete_by_key.get(self._cache_key(sender, timestamp)) or [])

    def _store_complete(self, sender: str, timestamp: str, body: str, now: float):
        key = self._cache_key(sender, timestamp)
        self._complete_by_key.setdefault(key, []).append((body, now))

    def _cache_key(self, sender: str, timestamp: str):
        return _normalize_sender_for_match(sender), str(timestamp or "")

    def _expire(self, now: float):
        self._complete_by_key = {
            key: items
            for key, items in (
                (key, [item for item in items if now - item[1] <= self.max_age])
                for key, items in self._complete_by_key.items()
            )
            if items
        }
        self._multipart = {
            key: entry
            for key, entry in self._multipart.items()
            if now - entry.get("seen", now) <= self.max_age
        }


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
