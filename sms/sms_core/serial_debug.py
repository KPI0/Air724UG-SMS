import re

from sms_core.phone_numbers import normalize_sms_destination


SERIAL_DEBUG_MAX_STORE_LINES = 20000
SERIAL_DEBUG_MAX_VISIBLE_LINES = 5000

COMMON_SERIAL_COMMANDS = [
    ("AT", "测试通信"),
    ("ATI", "查看模块信息"),
    ("AT+CGMR", "查看固件版本"),
    ("AT+CSQ", "查询信号 RSSI/通用"),
    ("AT+CESQ", "查询精确信号 4G RSRP"),
    ("AT+CGSN", "查询模块 IMEI"),
    ("AT+WISN?", "查询模块 SN"),
    ("AT+MIFIMAC=R", "查询 WiFi 热点 MAC 地址"),
    ("AT+CGPADDR", "查询 PDP 上下文 IP 地址"),
    ("AT+CGDCONT?", "查看 APN 配置"),
    ("AT+RFTEMPERATURE?", "查询模块温度"),
    ("AT+CNUM", "查询本机号码"),
    ("AT+CSCA?", "查询信息中心号码"),
    ("AT+CPBS?", "查看电话簿存储区"),
    ("AT+CPMS?", "查询短信存储状态"),
    ("AT+CEER", "查询最后一次呼叫错误"),
    ("AT+CCLK?", "查询模块时间"),
    ("AT+COPS?", "查询当前运营商"),
    ("AT+COPS=?", "查询附近可用运营商"),
    ("AT+COPS=0", "自动选择运营商"),
    ("AT+CPIN?", "查看 PIN 码锁状态"),
    ("AT+ICCID", "查询 SIM 卡 ICCID"),
    ("AT+CIMI", "查询 SIM 卡 IMSI"),
    ("AT+CGATT?", "查看网络附着状态"),
    ("AT+CFUN=1,1", "重启基带"),
    ("AT+RESET", "重启模块"),
    ("AT+CFUN?", "查看飞行模式状态"),
    ("AT+CFUN=0", "打开飞行模式"),
    ("AT+CFUN=1", "关闭飞行模式"),
    ("AT+EEMGINFO?", "查询基站定位数据"),
]

CALL_FORWARD_ALL_QUERY_COMMAND = "AT+CCFC=4,2"
CALL_FORWARD_QUERY_COMMAND = "AT+CCFC=0,2"
CALL_FORWARD_DISABLE_COMMAND = "AT+CCFC=4,4,0"


def quick_command_label(command: str, description: str) -> str:
    return f"{command}  ({description})"


def build_serial_command_payload(command: str, append_crlf: bool = True):
    text = str(command or "")
    suffix = "\r\n" if append_crlf else ""
    display_suffix = "\\r\\n" if append_crlf else ""
    return (text + suffix).encode("utf-8", "ignore"), display_suffix


def _normalize_pin(value, *, puk=False):
    text = str(value or "").strip()
    pattern = r"[0-9]{8}" if puk else r"[0-9]{4,8}"
    if not re.fullmatch(pattern, text):
        raise ValueError("PUK 码必须为 8 位半角数字" if puk else "PIN 码必须为 4-8 位半角数字")
    return text


def build_pin_unlock_command(pin: str) -> str:
    return f'AT+CPIN="{_normalize_pin(pin)}"'


def build_puk_unlock_command(puk: str, new_pin: str) -> str:
    return f'AT+CPIN="{_normalize_pin(puk, puk=True)}","{_normalize_pin(new_pin)}"'


def build_pin_lock_command(pin: str, enable: bool) -> str:
    mode = "1" if enable else "0"
    return f'AT+CLCK="SC",{mode},"{_normalize_pin(pin)}"'


def build_pin_change_command(old_pin: str, new_pin: str) -> str:
    return f'AT+CPWD="SC","{_normalize_pin(old_pin)}","{_normalize_pin(new_pin)}"'


def normalize_own_number(phone: str) -> str:
    text = str(phone or "").strip()
    if text and not text.startswith("+"):
        text = "+86" + text
    if not re.fullmatch(r"\+[1-9][0-9]{6,14}", text):
        raise ValueError("本机号码需为 7-15 位半角数字，并带有效的 + 国际前缀")
    return text


def build_own_number_commands(phone: str):
    normalized = normalize_own_number(phone)
    return 'AT+CPBS="ON"', f'AT+CPBW=1,"{normalized}",145', "AT+CNUM"


def normalize_information_center_number(phone: str) -> str:
    raw = str(phone or "").strip()
    if not raw:
        raise ValueError("信息中心号码不能为空")
    normalized = re.sub(r"[\s\-().（）]", "", raw)
    if not re.fullmatch(r"\+[1-9][0-9]{6,14}", normalized):
        raise ValueError("信息中心号码需以 + 开头，并包含 7-15 位数字")
    return normalized


def build_information_center_command(phone: str) -> str:
    normalized = normalize_information_center_number(phone)
    return f'AT+CSCA="{normalized}",145'


def normalize_call_forward_number(phone: str) -> str:
    raw = str(phone or "").strip()
    normalized = re.sub(r"[\s\-().（）]", "", raw)
    if normalized.startswith("+"):
        valid = bool(re.fullmatch(r"\+[1-9][0-9]{4,14}", normalized))
    else:
        valid = bool(re.fullmatch(r"[0-9]{5,15}", normalized))
    if not valid:
        raise ValueError("呼叫转移号码需为 5-15 位数字，可使用 + 国际前缀")
    return normalized


def build_call_forward_enable_command(phone: str) -> str:
    normalized = normalize_call_forward_number(phone)
    number_type = "145" if normalized.startswith("+") else "129"
    return f"AT+CCFC=0,3,{normalized},{number_type},1"


def normalize_operator_plmn(plmn: str) -> str:
    normalized = str(plmn or "").strip()
    if not re.fullmatch(r"[0-9]{5,6}", normalized):
        raise ValueError("运营商 PLMN 必须为 5-6 位数字")
    return normalized


def build_manual_operator_command(plmn: str) -> str:
    normalized = normalize_operator_plmn(plmn)
    return f'AT+COPS=1,2,"{normalized}"'


def build_sn_command(sn: str) -> str:
    text = str(sn or "").strip()
    if not 1 <= len(text) <= 64 or any(
        not "!" <= char <= "~" or char in "\"'\\;," for char in text
    ):
        raise ValueError("SN 码需为 1-64 位半角字符，不能包含空白、引号、反斜杠、逗号或分号")
    return f'AT+WISN={text}'


def normalize_dial_number(phone: str) -> str:
    text = str(phone or "").strip()
    if text.startswith(("*", "#")):
        if not re.fullmatch(r"[*#][0-9*#+]{1,63}", text) or not re.search(r"[0-9]", text):
            raise ValueError("拨号代码仅支持半角数字、*、# 和 +，最长 64 位")
        return text
    text = normalize_sms_destination(text)
    if text.startswith("+86"):
        return normalize_sms_destination(text[3:])
    if text.startswith("86") and len(text) == 13:
        return text[2:]
    return text


def build_dial_command(phone: str) -> str:
    return f"ATD{normalize_dial_number(phone)};"


HANGUP_COMMAND = "ATH"
