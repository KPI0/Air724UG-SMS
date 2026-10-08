import re


CALL_FILTER_NUMBER_RE = re.compile(r"^\+?[0-9]+$")


def normalize_sms_destination(number: str, *, allow_empty=False) -> str:
    text = str(number or "").strip()
    if allow_empty and not text:
        return ""
    if not re.fullmatch(r"\+?[0-9]{1,20}", text):
        raise ValueError("号码需为 1-20 位半角数字，可在开头加 +")
    return text


def normalize_call_number(number: str) -> str:
    """Normalize caller ID for local whitelist/blacklist matching."""
    text = str(number or "").strip()
    if text.startswith("+86"):
        return text[3:]
    if text.startswith("86") and len(text) > 11:
        return text[2:]
    return text


def is_valid_call_filter_number(number: str) -> bool:
    """Return whether a call filter entry is a numeric caller ID."""
    return bool(CALL_FILTER_NUMBER_RE.fullmatch(str(number or "").strip()))
