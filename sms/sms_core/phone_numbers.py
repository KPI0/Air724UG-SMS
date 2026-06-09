def normalize_call_number(number: str) -> str:
    """Normalize caller ID for local whitelist/blacklist matching."""
    text = str(number or "").strip()
    if text.startswith("+86"):
        return text[3:]
    if text.startswith("86") and len(text) > 11:
        return text[2:]
    return text

