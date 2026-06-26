import re

from sms_core.phone_numbers import normalize_call_number


LOCAL_NUMBER_RE = re.compile(r"\+?\d{5,}")
SERIAL_AT_PREFIX = r"(?:\[[IWE]\]-\[[^\]]+\]\s*)?"
CLIP_NUMBER_RE = re.compile(rf'^\s*{SERIAL_AT_PREFIX}\+CLIP:\s*"?(\+?\d+)"?')
TEMPERATURE_RE = re.compile(
    rf"^\s*{SERIAL_AT_PREFIX}\+RFTEMPERATURE:\s*(-?\d+(?:\.\d+)?)\s*$"
)
CESQ_RE = re.compile(
    rf"^\s*{SERIAL_AT_PREFIX}\+CESQ:\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*(\d+)\s*$"
)
AT_RESPONSE_RE_TEMPLATE = r"^\s*" + SERIAL_AT_PREFIX + r"{command}:\s*(?P<body>.*?)\s*$"
SERIAL_PREFIX_RE = re.compile(rf"^\s*{SERIAL_AT_PREFIX}(?P<body>.*?)\s*$")
CIEV_CALL_RE = re.compile(r'^\+CIEV:\s*"CALL"\s*,\s*(?P<state>[01])$', re.IGNORECASE)


def _at_response_body(line: str, command: str):
    pattern = AT_RESPONSE_RE_TEMPLATE.format(command=re.escape(command))
    match = re.match(pattern, str(line or ""))
    return match.group("body").strip() if match else None


def _serial_body(line: str):
    match = SERIAL_PREFIX_RE.match(str(line or ""))
    return match.group("body").strip() if match else str(line or "").strip()


def parse_temperature(line: str):
    match = TEMPERATURE_RE.match(str(line or ""))
    return match.group(1) if match else None


def parse_cesq_rsrp(line: str):
    match = CESQ_RE.match(str(line or ""))
    return match.group(1) if match else None


def parse_cnum_number(line: str):
    body = _at_response_body(line, "+CNUM")
    if body is None:
        return None
    matches = LOCAL_NUMBER_RE.findall(body)
    if not matches:
        return None
    return max(matches, key=lambda item: len(item.lstrip("+"))).strip()


def parse_local_number(line: str):
    return parse_cnum_number(line)


def parse_operator_message(line: str):
    body = _at_response_body(line, "+COPS")
    if body is None or '"' not in body:
        return None
    try:
        plmn = body.split('"')[1]
    except Exception:
        return None

    carrier_map = {
        "46000": "中国移动",
        "46002": "中国移动",
        "46007": "中国移动",
        "46008": "中国移动",
        "46001": "中国联通",
        "46006": "中国联通",
        "46009": "中国联通",
        "46003": "中国电信",
        "46005": "中国电信",
        "46011": "中国电信",
        "46013": "中国电信(物联卡)",
        "46004": "中国移动(物联卡)",
        "46015": "中国广电",
    }
    carrier = carrier_map.get(plmn, "未知运营商")
    return f">>> 识别到网络运营商：{carrier} ({plmn})"


def parse_sim_status_message(line: str):
    status = _at_response_body(line, "+CPIN")
    if status is None:
        return None

    status_map = {
        "READY": "PIN码锁未开启 (SIM卡正常)",
        "SIM PIN": "等待输入PIN码 (卡已被锁)",
        "SIM PUK": "等待输入PUK码 (需PUK解锁)",
        "NOT INSERTED": "未检测到SIM卡",
        "NOT READY": "SIM卡未准备好",
    }
    status_text = status_map.get(status, f"未知状态 ({status})")
    return f">>> 识别到SIM卡状态：{status_text}"


def parse_attach_status_message(line: str):
    status = _at_response_body(line, "+CGATT")
    if status is None:
        return None

    if status == "1":
        status_text = "已附着数据网络 (通道已打通，可以上网)"
    elif status == "0":
        status_text = "未附着数据网络 (无法连接互联网)"
    else:
        status_text = f"未知状态 ({status})"
    return f">>> 识别到网络附着状态：{status_text}"


def parse_cfun_status_message(line: str):
    status = _at_response_body(line, "+CFUN")
    if status is None:
        return None

    status_map = {
        "0": "飞行模式 (射频关闭)",
        "1": "正常模式 (射频开启)",
    }
    status_text = status_map.get(status, f"未知状态 ({status})")
    return f">>> 识别到射频/飞行模式状态：{status_text}"


def parse_lte_cell_messages(line: str):
    body = _at_response_body(line, "+EEMLTESVC")
    if body is None:
        return []
    try:
        parts = body.split(",")
        if len(parts) < 10:
            return []
        raw_mcc = int(parts[0].strip())
        raw_mnc = int(parts[1].strip())
        tac = parts[3].strip()
        ci = parts[9].strip()
    except Exception:
        return []

    return [
        ">>> 解析到基站定位数据：",
        f"    MCC (国家代码) : {raw_mcc:x}",
        f"    MNC (网络代码) : {raw_mnc:02x}",
        f"    LAC/TAC (区域) : {tac}",
        f"    CI  (小区ID)   : {ci}",
    ]


def parse_serial_debug_insights(line: str):
    messages = []
    for parser in (
        parse_operator_message,
        parse_sim_status_message,
        parse_attach_status_message,
        parse_cfun_status_message,
    ):
        msg = parser(line)
        if msg:
            messages.append(msg)
    messages.extend(parse_lte_cell_messages(line))
    return messages


def parse_clip_number(line: str, default: str = "未知号码") -> str:
    match = CLIP_NUMBER_RE.search(str(line or ""))
    return match.group(1) if match else default


def evaluate_call_filter(caller_num: str, mode: str, whitelist, blacklist):
    normalized = normalize_call_number(caller_num)
    mode = str(mode or "").strip()

    if mode == "Whitelist":
        for number in whitelist or []:
            if normalize_call_number(number) == normalized:
                return False, ""
        return True, "不在白名单"

    if mode == "Blacklist":
        for number in blacklist or []:
            if normalize_call_number(number) == normalized:
                return True, "命中黑名单"

    return False, ""


def is_new_clip(caller_num: str, last_clip_num: str, now: float, last_clip_time: float, window_sec: float = 4.0) -> bool:
    return caller_num != last_clip_num or now - last_clip_time > window_sec


def is_ring_line(line: str) -> bool:
    return _serial_body(line) == "RING"


def is_hangup_event(line: str) -> bool:
    text = _serial_body(line)
    compact = re.sub(r"\s+", "", text).upper()
    ciev = CIEV_CALL_RE.match(text)
    return (
        text in ("NO CARRIER", "BUSY", "NO ANSWER")
        or bool(ciev and ciev.group("state") == "0")
    )


def is_call_connected_event(line: str) -> bool:
    text = _serial_body(line)
    ciev = CIEV_CALL_RE.match(text)
    return text == "CONNECT" or bool(ciev and ciev.group("state") == "1")


def is_sms_collection_boundary(line: str) -> bool:
    text = str(line or "")
    return (
        text.startswith("[I]-")
        or text.startswith("[W]-")
        or text.startswith("[E]-")
        or text.startswith(">>>")
        or text.startswith("AT+")
        or text.startswith("+")
    )
