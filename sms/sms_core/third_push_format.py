import re
import time

from sms_core.config_schema import THIRD_PUSH_SMS_TEMPLATE


def template_vars(raw_msg: str, port: str = ""):
    return {
        "{msg}": raw_msg,
        "{raw_msg}": raw_msg,
        "{time}": time.strftime("%Y-%m-%d %H:%M:%S"),
        "{port}": port,
    }


def apply_vars(value, variables):
    if isinstance(value, dict):
        return {key: apply_vars(item, variables) for key, item in value.items()}
    if isinstance(value, list):
        return [apply_vars(item, variables) for item in value]
    if isinstance(value, str):
        return re.sub(
            r"\{(?:msg|raw_msg|time|port)\}",
            lambda match: variables.get(match.group(0), match.group(0)),
            value,
        )
    return value


def format_message(raw_msg: str, template: str = None, port: str = "") -> str:
    text = template if template is not None else THIRD_PUSH_SMS_TEMPLATE
    if not str(text or "").strip():
        text = "{msg}"
    return str(apply_vars(str(text), template_vars(raw_msg, port))).strip()
