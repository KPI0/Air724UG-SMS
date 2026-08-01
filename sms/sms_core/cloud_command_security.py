import re
from dataclasses import dataclass


CLOUD_SENSITIVE_COMMANDS_OPTION = "allow_sensitive_commands"
CLOUD_SEND_SMS_TRANSACTION_COMMAND = "CLOUD:SEND_SMS"
CLOUD_SET_OWN_NUMBER_TRANSACTION_COMMAND = "CLOUD:SET_OWN_NUMBER"
LEGACY_PERMISSION_OPTIONS = (
    "allow_sensitive_sim_security",
    "allow_sensitive_network",
    "allow_sensitive_device",
    "allow_sensitive_privacy",
    "allow_sensitive_raw_data",
    "allow_sensitive_other",
    "allow_sensitive_pin_lock",
)


@dataclass(frozen=True)
class CloudCommandPermissionSpec:
    category: str
    option: str
    label: str
    description: str


@dataclass(frozen=True)
class SensitiveCloudCommandDecision:
    category: str = ""
    reason: str = ""

    @property
    def sensitive(self):
        return bool(self.category)


CLOUD_COMMAND_PERMISSION_SPECS = (
    CloudCommandPermissionSpec(
        "sms",
        "allow_sensitive_sms",
        "发送短信",
        "允许云端发送短信和下发短信 PDU。",
    ),
    CloudCommandPermissionSpec(
        "call",
        "allow_sensitive_call",
        "拨打电话",
        "允许云端向指定号码发起拨号。",
    ),
    CloudCommandPermissionSpec(
        "pin",
        "allow_sensitive_pin",
        "PIN 码操作",
        "允许提交、验证或修改 PIN 码，以及开启或关闭 SIM 卡 PIN 码锁。"
        "错误操作可能导致 SIM 卡被锁定。",
    ),
    CloudCommandPermissionSpec(
        "puk",
        "allow_sensitive_puk",
        "PUK 码操作",
        "允许提交 PUK 并设置新的 PIN。",
    ),
    CloudCommandPermissionSpec(
        "phone_number",
        "allow_sensitive_phone_number",
        "修改本机号码",
        "允许修改 SIM 中保存的本机号码。",
    ),
    CloudCommandPermissionSpec(
        "sn",
        "allow_sensitive_sn",
        "修改 SN 码",
        "允许通过 AT+WISN 修改设备 SN 码。",
    ),
    CloudCommandPermissionSpec(
        "cell_location",
        "allow_sensitive_cell_location",
        "查询基站定位数据",
        "允许通过 AT+EEMGINFO 查询基站定位相关数据。",
    ),
    CloudCommandPermissionSpec(
        "ussd",
        "allow_sensitive_ussd",
        "使用 USSD 服务",
        "允许通过 AT+CUSD 或拨号代码使用运营商 USSD 服务，部分操作可能产生费用。",
    ),
    CloudCommandPermissionSpec(
        "call_control",
        "allow_sensitive_call_control",
        "设置呼叫转移或呼叫限制",
        "允许设置呼叫转移、呼叫限制和固定拨号等通话限制功能。",
    ),
    CloudCommandPermissionSpec(
        "sms_center",
        "allow_sensitive_sms_center",
        "修改信息中心号码",
        "允许修改信息中心号码，错误设置可能导致短信发送异常或产生额外费用。",
    ),
    CloudCommandPermissionSpec(
        "delete_data",
        "allow_sensitive_delete_data",
        "删除设备数据",
        "允许删除短信、电话簿记录或设备文件等数据。",
    ),
    CloudCommandPermissionSpec(
        "device_power",
        "allow_sensitive_device_power",
        "重置或关闭设备",
        "允许重启、复位、关闭设备或恢复设备配置。",
    ),
)

CLOUD_COMMAND_PERMISSION_DEFAULTS = {
    spec.option: "0"
    for spec in CLOUD_COMMAND_PERMISSION_SPECS
}

_SMS_SEND_PREFIXES = (
    "AT+CMGS",
    "AT+CMSS",
    "AT+CMGC",
    "AT+CMGF=",
)
_PIN_CHANGE_PREFIXES = (
    "AT+CPIN",
    "AT+CPIN2",
)
_PHONE_NUMBER_CHANGE_PREFIXES = (
    'AT+CPBS="ON"',
)
_SN_CHANGE_PREFIXES = (
    "AT+WISN=",
)
_PIN_LOCK_FACILITIES = {
    "SC",
    "P2",
    "PS",
    "PF",
    "PN",
    "PU",
    "PP",
    "PC",
}
_DELETE_DATA_PREFIXES = (
    "AT+CMGD",
    "AT+CMGDA",
    "AT+FDELETE",
    "AT+FSDEL",
    "AT+FSDELETE",
    "AT+FSCLEAR",
    "AT+FSFORMAT",
    "AT+QFDEL",
)
_DEVICE_POWER_PREFIXES = (
    "AT+REBOOT",
    "AT+QRESET",
    "AT+RST",
    "AT+CFUN=",
    "AT+CPOWD",
    "AT+POWD",
    "AT+QPOWD",
    "AT+CRESET",
    "AT+QRST",
    "AT&F",
    "ATZ",
)
_CALL_CONTROL_MMI_CODES = {
    "21",
    "61",
    "62",
    "67",
    "002",
    "004",
    "33",
    "331",
    "332",
    "35",
    "351",
    "330",
    "333",
    "353",
}
_PIN_MMI_CODES = {"04", "042"}
_PUK_MMI_CODES = {"05", "052"}
UNSUPPORTED_CONTROL_CHAR_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x19\x1b-\x1f\x7f]")
UNSUPPORTED_CONTROL_CHAR_MESSAGE = "AT 指令包含不支持的控制字符"


def _command_facility(compact, command_prefix):
    match = re.match(
        rf'^{re.escape(command_prefix)}=(?:"([^"]+)"|([^,]+))',
        compact,
    )
    if match is None:
        return ""
    return str(match.group(1) or match.group(2) or "").strip().upper()


def _is_phonebook_delete_command(compact):
    if not compact.startswith("AT+CPBW="):
        return False
    payload = compact.split("=", 1)[1]
    parts = payload.split(",", 2)
    if len(parts) == 1:
        return True
    return not parts[1].strip().strip('"')


def _is_ussd_dial_command(compact):
    return (
        compact.startswith(("ATD*", "ATD#"))
        and "#" in compact[4:]
    )


def _mmi_service_code(compact):
    match = re.match(r"^ATD(?:\*\*|##|\*#|\*|#)(\d{2,3})(?:\*|#)", compact)
    if match is None:
        return ""
    return match.group(1)


def normalize_cloud_command_permissions(values, default=False):
    if isinstance(values, bool):
        return {
            spec.category: values
            for spec in CLOUD_COMMAND_PERMISSION_SPECS
        }

    source = values if isinstance(values, dict) else {}
    normalized = {}
    for spec in CLOUD_COMMAND_PERMISSION_SPECS:
        value = source.get(spec.category, source.get(spec.option, default))
        if isinstance(value, str):
            value = value.strip().lower() in ("1", "true", "yes", "on")
        normalized[spec.category] = bool(value)
    legacy_pin_lock = source.get(
        "pin_lock",
        source.get("allow_sensitive_pin_lock", False),
    )
    if isinstance(legacy_pin_lock, str):
        legacy_pin_lock = legacy_pin_lock.strip().lower() in ("1", "true", "yes", "on")
    if legacy_pin_lock:
        normalized["pin"] = True
    return normalized


def _read_config_bool(config, option, fallback=False):
    try:
        return config.getboolean("cloud_control", option, fallback=fallback)
    except Exception:
        return bool(fallback)


def read_cloud_command_permissions(config):
    if _read_config_bool(config, CLOUD_SENSITIVE_COMMANDS_OPTION):
        return normalize_cloud_command_permissions(True)

    permissions = {
        spec.category: _read_config_bool(config, spec.option)
        for spec in CLOUD_COMMAND_PERMISSION_SPECS
    }

    # Compatibility with the previous split PIN settings: the old PIN-lock
    # switch now grants the combined PIN operation permission.
    if _read_config_bool(config, "allow_sensitive_sim_security"):
        permissions["pin"] = True
        permissions["puk"] = True
    if _read_config_bool(config, "allow_sensitive_pin_lock"):
        permissions["pin"] = True
    if _read_config_bool(config, "allow_sensitive_other"):
        permissions["phone_number"] = True
        permissions["sn"] = True
    return permissions


def read_cloud_sensitive_commands_enabled(config):
    return any(read_cloud_command_permissions(config).values())


def cloud_command_has_line_break(command):
    """Return whether a cloud-provided command contains a serial line break.

    A CR or LF inside a cloud command can turn one payload into multiple AT
    commands when it reaches the modem.  This check intentionally happens on
    the original text, before whitespace trimming or command classification.
    """
    text = str(command or "")
    return "\r" in text or "\n" in text


def cloud_command_has_chained_separator(command):
    """Detect separators that append another command to the same payload.

    A trailing unquoted semicolon is valid for dial/MMI commands.  Ctrl-Z is
    valid as the final byte of a PDU submission.  Either character becomes a
    command-chaining boundary when non-whitespace content follows it.
    """
    text = str(command or "")
    in_quotes = False
    escaped = False
    for index, char in enumerate(text):
        if escaped:
            escaped = False
            continue
        if char == "\\" and in_quotes:
            escaped = True
            continue
        if char == '"':
            in_quotes = not in_quotes
            continue
        if char == ";" and not in_quotes and text[index + 1:].strip():
            return True
        if char == "\x1a" and text[index + 1:].strip():
            return True
    return False


def cloud_command_batch_error(command):
    if cloud_command_has_line_break(command):
        return "AT 指令不允许包含回车或换行，请每次只发送一条指令"
    if cloud_command_has_chained_separator(command):
        return "AT 指令不允许串联多条指令，请每次只发送一条指令"
    return ""


def cloud_command_control_char_error(command):
    """Return the same control-character rejection used by the server policy.

    Tab and a terminal Ctrl-Z are intentionally allowed.  CR/LF and chained
    payloads remain handled by ``cloud_command_batch_error``.
    """
    if UNSUPPORTED_CONTROL_CHAR_RE.search(str(command or "")):
        return UNSUPPORTED_CONTROL_CHAR_MESSAGE
    return ""


def sensitive_cloud_command_decision(command, command_meta=None):
    raw = str(command or "").strip()
    if not raw:
        return SensitiveCloudCommandDecision()

    # Security decisions must be derived from the command itself.  Metadata is
    # supplied by the remote web client and is only suitable for presentation;
    # allowing it to select a permission category lets a reset or PIN command
    # masquerade as an SMS command.
    if "\x1a" in raw:
        return SensitiveCloudCommandDecision("sms", "发送短信")

    if raw.upper() == CLOUD_SEND_SMS_TRANSACTION_COMMAND:
        return SensitiveCloudCommandDecision("sms", "发送短信")
    if raw.upper() == CLOUD_SET_OWN_NUMBER_TRANSACTION_COMMAND:
        return SensitiveCloudCommandDecision("phone_number", "修改本机号码")
    compact = re.sub(r"\s+", "", raw).upper()
    if compact.startswith("AT+EEMGINFO"):
        return SensitiveCloudCommandDecision("cell_location", "查询基站定位数据")
    if compact.endswith("?") or compact.endswith("=?"):
        return SensitiveCloudCommandDecision()

    if compact.startswith(_SMS_SEND_PREFIXES):
        return SensitiveCloudCommandDecision("sms", "发送短信")
    if compact.startswith("AT+CSCA="):
        return SensitiveCloudCommandDecision("sms_center", "修改信息中心号码")
    if compact.startswith(("AT+CPIN=", "AT+CPIN2=")) and "," in compact:
        return SensitiveCloudCommandDecision("puk", "PUK 码操作")
    if compact.startswith("AT+CLCK"):
        if _command_facility(compact, "AT+CLCK") in _PIN_LOCK_FACILITIES:
            return SensitiveCloudCommandDecision("pin", "PIN 码操作（开启或关闭 PIN 码锁）")
        return SensitiveCloudCommandDecision("call_control", "设置呼叫转移或呼叫限制")
    if compact.startswith("AT+CPWD"):
        if _command_facility(compact, "AT+CPWD") in _PIN_LOCK_FACILITIES:
            return SensitiveCloudCommandDecision("pin", "PIN 码操作")
        return SensitiveCloudCommandDecision("call_control", "设置呼叫转移或呼叫限制")
    if compact.startswith(_PIN_CHANGE_PREFIXES):
        return SensitiveCloudCommandDecision("pin", "PIN 码操作")
    mmi_code = _mmi_service_code(compact)
    if mmi_code in _PIN_MMI_CODES:
        return SensitiveCloudCommandDecision("pin", "PIN 码操作")
    if mmi_code in _PUK_MMI_CODES:
        return SensitiveCloudCommandDecision("puk", "PUK 码操作")
    if mmi_code in _CALL_CONTROL_MMI_CODES:
        return SensitiveCloudCommandDecision("call_control", "设置呼叫转移或呼叫限制")
    if compact.startswith("AT+CUSD") or _is_ussd_dial_command(compact):
        return SensitiveCloudCommandDecision("ussd", "使用 USSD 服务")
    if compact.startswith("AT+CCFC"):
        return SensitiveCloudCommandDecision("call_control", "设置呼叫转移或呼叫限制")
    if compact.startswith("ATD"):
        return SensitiveCloudCommandDecision("call", "拨打电话")
    if compact.startswith(_DELETE_DATA_PREFIXES) or _is_phonebook_delete_command(compact):
        return SensitiveCloudCommandDecision("delete_data", "删除设备数据")
    if compact.startswith(_PHONE_NUMBER_CHANGE_PREFIXES):
        return SensitiveCloudCommandDecision("phone_number", "修改本机号码")
    if compact.startswith("AT+CPBW"):
        return SensitiveCloudCommandDecision("phone_number", "修改本机号码")
    if compact.startswith(_SN_CHANGE_PREFIXES):
        return SensitiveCloudCommandDecision("sn", "修改 SN 码")
    if compact.startswith(_DEVICE_POWER_PREFIXES):
        return SensitiveCloudCommandDecision("device_power", "重置或关闭设备")
    return SensitiveCloudCommandDecision()


def sensitive_cloud_command_reason(command, command_meta=None):
    return sensitive_cloud_command_decision(command, command_meta).reason


def is_sensitive_cloud_command_allowed(decision, permissions):
    if not getattr(decision, "sensitive", False):
        return True
    normalized = normalize_cloud_command_permissions(permissions)
    return bool(normalized.get(decision.category, False))


def cloud_sensitive_command_block_message(reason):
    reason = str(reason or "敏感操作").strip()
    return (
        f"安全设置已拦截云端敏感指令：{reason}。"
        "如确认需要，请在客户端“设置 > 云端控制 > 安全设置”中开启对应权限。"
    )


def cloud_sensitive_commands_status(permissions):
    normalized = normalize_cloud_command_permissions(permissions)
    allowed = sum(1 for value in normalized.values() if value)
    total = len(CLOUD_COMMAND_PERMISSION_SPECS)
    if allowed == 0:
        return "云端敏感指令：全部关闭"
    if allowed == total:
        return "云端敏感指令：全部开启"
    return f"云端敏感指令：已开启 {allowed}/{total} 项"
