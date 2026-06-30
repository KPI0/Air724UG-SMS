from dataclasses import dataclass
import re
import secrets
from typing import Optional


SINGLE_SMS_UCS2_BYTES = 140
CONCAT_SMS_UCS2_BYTES = 134
CONCAT_SMS_SEGMENT_LIMIT = 255


@dataclass(frozen=True)
class TextSmsPduInfo:
    char_count: int
    ucs2_bytes: int
    segment_count: int
    segment_limit: int = CONCAT_SMS_SEGMENT_LIMIT

    @property
    def too_long(self):
        return self.segment_count > self.segment_limit


@dataclass(frozen=True)
class ConcatSmsInfo:
    reference: int
    total: int
    index: int
    reference_bits: int = 8


@dataclass(frozen=True)
class ReceivedSmsPdu:
    sender: str
    body: str
    timestamp: str = ""
    reference: Optional[int] = None
    total: Optional[int] = None
    index: Optional[int] = None
    reference_bits: Optional[int] = None

    @property
    def concat_info(self):
        if self.reference is None or self.total is None or self.index is None:
            return None
        return ConcatSmsInfo(
            self.reference,
            self.total,
            self.index,
            self.reference_bits or 8,
        )


def _encode_phone_number(phone: str):
    phone_text = str(phone or "").strip()
    if phone_text.startswith("+"):
        number_type = "91"
        number = phone_text[1:]
    else:
        number_type = "81"
        number = phone_text
    if phone_text and (not number or not number.isdigit()):
        raise ValueError("手机号只能包含数字和可选开头 +")
    number_len = f"{len(number):02X}"

    if len(number) % 2 != 0:
        number += "F"
    swapped_number = "".join(number[i + 1] + number[i] for i in range(0, len(number), 2))
    return number_len, number_type, swapped_number


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


def _hex_to_bytes(value):
    if isinstance(value, bytes):
        return value
    if isinstance(value, bytearray):
        return bytes(value)
    text = re.sub(r"\s+", "", str(value or ""))
    if not text or len(text) % 2:
        return b""
    try:
        return bytes.fromhex(text)
    except ValueError:
        return b""


def decode_concat_udh(udh):
    data = _hex_to_bytes(udh)
    if not data:
        return None

    payload = data
    if len(data) >= 2 and data[0] == len(data) - 1:
        payload = data[1:]

    pos = 0
    while pos + 2 <= len(payload):
        iei = payload[pos]
        iedl = payload[pos + 1]
        pos += 2
        content = payload[pos:pos + iedl]
        pos += iedl
        if len(content) != iedl:
            break
        if iei == 0x00 and iedl == 3:
            return _valid_concat_info(content[0], content[1], content[2], 8)
        if iei == 0x08 and iedl == 4:
            reference = (content[0] << 8) | content[1]
            return _valid_concat_info(reference, content[2], content[3], 16)
    return None


def _valid_concat_info(reference, total, index, reference_bits):
    total = int(total or 0)
    index = int(index or 0)
    if total <= 1 or index < 1 or index > total:
        return None
    return ConcatSmsInfo(int(reference), total, index, int(reference_bits))


def _fallback_concat_payload(user_data_hex: str):
    data = _hex_to_bytes(user_data_hex)
    if len(data) >= 6 and data[0:3] == b"\x05\x00\x03":
        info = _valid_concat_info(data[3], data[4], data[5], 8)
        if info is not None:
            return info, data[6:].hex().upper()
    if len(data) >= 7 and data[0:3] == b"\x06\x08\x04":
        reference = (data[3] << 8) | data[4]
        info = _valid_concat_info(reference, data[5], data[6], 16)
        if info is not None:
            return info, data[7:].hex().upper()
    return None, user_data_hex


def _has_concat_udh_marker(udh_hex: str):
    data = _hex_to_bytes(udh_hex)
    if not data:
        return False
    payload = data[1:] if len(data) >= 2 and data[0] == len(data) - 1 else data
    pos = 0
    while pos + 2 <= len(payload):
        iei = payload[pos]
        iedl = payload[pos + 1]
        pos += 2
        content = payload[pos:pos + iedl]
        pos += iedl
        if len(content) != iedl:
            return iei in (0x00, 0x08)
        if (iei == 0x00 and iedl == 3) or (iei == 0x08 and iedl == 4):
            return True
    return False


def _has_fallback_concat_marker(user_data_hex: str):
    data = _hex_to_bytes(user_data_hex)
    return (
        len(data) >= 3
        and (data[0:3] == b"\x05\x00\x03" or data[0:3] == b"\x06\x08\x04")
    )


def decode_received_pdu(pdu_hex: str):
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

        concat_info = None
        payload = user_data
        if first_octet & 0x40:
            udhl, _udh_offset = _read_octet(user_data, 0)
            udh_hex_len = udhl * 2
            udh_with_len = user_data[:2 + udh_hex_len]
            payload = user_data[2 + udh_hex_len:]
            concat_info = decode_concat_udh(udh_with_len)
            if concat_info is None and _has_concat_udh_marker(udh_with_len):
                return None
        else:
            if _has_fallback_concat_marker(user_data):
                concat_info, payload = _fallback_concat_payload(user_data)
                if concat_info is None:
                    return None
            else:
                concat_info, payload = None, user_data

        if dcs != 0x08:
            return None
        body = bytes.fromhex(payload).decode("utf-16-be")
        return ReceivedSmsPdu(
            sender=sender,
            body=body,
            timestamp=timestamp,
            reference=concat_info.reference if concat_info else None,
            total=concat_info.total if concat_info else None,
            index=concat_info.index if concat_info else None,
            reference_bits=concat_info.reference_bits if concat_info else None,
        )
    except Exception:
        return None


def _split_ucs2_segments(message: str, max_bytes: int):
    segments = []
    current = []
    current_len = 0
    for char in str(message or ""):
        encoded = char.encode("utf-16-be")
        if current and current_len + len(encoded) > max_bytes:
            segments.append("".join(current).encode("utf-16-be"))
            current = []
            current_len = 0
        current.append(char)
        current_len += len(encoded)
    if current or not segments:
        segments.append("".join(current).encode("utf-16-be"))
    return segments


def _count_ucs2_segments(message: str, max_bytes: int):
    segment_count = 0
    current_len = 0
    for char in str(message or ""):
        char_len = len(char.encode("utf-16-be"))
        if current_len and current_len + char_len > max_bytes:
            segment_count += 1
            current_len = 0
        current_len += char_len
    if current_len or not segment_count:
        segment_count += 1
    return segment_count


def measure_text_sms_pdus(message: str):
    message_text = str(message or "")
    message_bytes = message_text.encode("utf-16-be")
    if len(message_bytes) <= SINGLE_SMS_UCS2_BYTES:
        segment_count = 1
    else:
        segment_count = _count_ucs2_segments(message_text, CONCAT_SMS_UCS2_BYTES)
    return TextSmsPduInfo(
        char_count=len(message_text),
        ucs2_bytes=len(message_bytes),
        segment_count=segment_count,
    )


def _build_text_sms_pdu(phone, message_bytes, *, first_octet="11", udh=b""):
    number_len, number_type, swapped_number = _encode_phone_number(phone)
    user_data = udh.hex().upper() + message_bytes.hex().upper()
    user_data_len = f"{len(udh) + len(message_bytes):02X}"

    smsc = "00"
    tpdu = f"{first_octet}00{number_len}{number_type}{swapped_number}0008C0{user_data_len}{user_data}"
    return smsc + tpdu, len(tpdu) // 2


def encode_text_sms_pdu(phone: str, message: str):
    """
    Encode a single UCS2 SMS in PDU mode.

    Returns (pdu_hex, cmgs_length). cmgs_length is the TPDU byte length,
    excluding the SMSC field, as required by AT+CMGS.
    """
    message_bytes = str(message or "").encode("utf-16-be")
    if len(message_bytes) > SINGLE_SMS_UCS2_BYTES:
        raise ValueError("短信内容超过单条 UCS2 PDU 容量，请使用 encode_text_sms_pdus 分段编码")
    return _build_text_sms_pdu(phone, message_bytes)


def encode_text_sms_pdus(phone: str, message: str, *, reference=None):
    """
    Encode UCS2 SMS text into one or more PDU segments.

    Concatenated SMS uses a 6-octet UDH, leaving 134 bytes for UCS2 data per
    segment. The reference is one byte, matching common GSM 03.40 UDH format.
    """
    message_bytes = str(message or "").encode("utf-16-be")
    if len(message_bytes) <= SINGLE_SMS_UCS2_BYTES:
        return [encode_text_sms_pdu(phone, message)]

    ref = secrets.randbelow(256) if reference is None else int(reference) & 0xFF
    segments = _split_ucs2_segments(message, CONCAT_SMS_UCS2_BYTES)
    total = len(segments)
    if total > CONCAT_SMS_SEGMENT_LIMIT:
        raise ValueError("短信内容过长，超过 255 个分段")

    pdus = []
    for index, segment in enumerate(segments, start=1):
        udh = bytes((0x05, 0x00, 0x03, ref, total, index))
        pdus.append(_build_text_sms_pdu(phone, segment, first_octet="51", udh=udh))
    return pdus
