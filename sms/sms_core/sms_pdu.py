import secrets


def _encode_phone_number(phone: str):
    phone_text = str(phone or "").strip()
    number_type = "91" if phone_text.startswith("+") else "81"
    number = phone_text.lstrip("+")
    number_len = f"{len(number):02X}"

    if len(number) % 2 != 0:
        number += "F"
    swapped_number = "".join(number[i + 1] + number[i] for i in range(0, len(number), 2))
    return number_len, number_type, swapped_number


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
    if len(message_bytes) > 140:
        raise ValueError("短信内容超过单条 UCS2 PDU 容量，请使用 encode_text_sms_pdus 分段编码")
    return _build_text_sms_pdu(phone, message_bytes)


def encode_text_sms_pdus(phone: str, message: str, *, reference=None):
    """
    Encode UCS2 SMS text into one or more PDU segments.

    Concatenated SMS uses a 6-octet UDH, leaving 134 bytes for UCS2 data per
    segment. The reference is one byte, matching common GSM 03.40 UDH format.
    """
    message_bytes = str(message or "").encode("utf-16-be")
    if len(message_bytes) <= 140:
        return [encode_text_sms_pdu(phone, message)]

    ref = secrets.randbelow(256) if reference is None else int(reference) & 0xFF
    segments = _split_ucs2_segments(message, 134)
    total = len(segments)
    if total > 255:
        raise ValueError("短信内容过长，超过 255 个分段")

    pdus = []
    for index, segment in enumerate(segments, start=1):
        udh = bytes((0x05, 0x00, 0x03, ref, total, index))
        pdus.append(_build_text_sms_pdu(phone, segment, first_octet="51", udh=udh))
    return pdus
