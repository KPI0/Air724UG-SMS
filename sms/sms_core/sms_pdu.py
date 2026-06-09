def encode_text_sms_pdu(phone: str, message: str):
    """
    Encode a single UCS2 SMS in PDU mode.

    Returns (pdu_hex, cmgs_length). cmgs_length is the TPDU byte length,
    excluding the SMSC field, as required by AT+CMGS.
    """
    phone_text = str(phone or "").strip()
    number_type = "91" if phone_text.startswith("+") else "81"
    number = phone_text.lstrip("+")
    number_len = f"{len(number):02X}"

    if len(number) % 2 != 0:
        number += "F"
    swapped_number = "".join(number[i + 1] + number[i] for i in range(0, len(number), 2))

    message_bytes = str(message or "").encode("utf-16-be")
    user_data_len = f"{len(message_bytes):02X}"
    user_data = message_bytes.hex().upper()

    smsc = "00"
    tpdu = f"1100{number_len}{number_type}{swapped_number}0008C0{user_data_len}{user_data}"
    return smsc + tpdu, len(tpdu) // 2

