import re
import time

from sms_core.config_schema import THIRD_PUSH_SMS_TEMPLATE


def template_vars(raw_msg: str, port: str = "", variables=None):
    values = {
        "{msg}": raw_msg,
        "{raw_msg}": raw_msg,
        "{time}": time.strftime("%Y-%m-%d %H:%M:%S"),
        "{port}": port,
        "{sender}": "",
        "{from}": "",
        "{phone}": "",
        "{local_number}": "",
        "{self_number}": "",
        "{sms_time}": "",
        "{caller}": "",
        "{call_time}": "",
    }
    for key, value in dict(variables or {}).items():
        name = str(key or "").strip()
        if not name:
            continue
        if not (name.startswith("{") and name.endswith("}")):
            name = "{" + name + "}"
        values[name] = str(value or "")
    if not values["{sms_time}"]:
        values["{sms_time}"] = values["{time}"]
    if not values["{call_time}"]:
        values["{call_time}"] = values["{time}"]
    return values


def apply_vars(value, variables):
    if isinstance(value, dict):
        return {key: apply_vars(item, variables) for key, item in value.items()}
    if isinstance(value, list):
        return [apply_vars(item, variables) for item in value]
    if isinstance(value, str):
        return re.sub(
            r"\{(?:msg|raw_msg|time|port|sender|from|phone|local_number|self_number|sms_time|caller|call_time)\}",
            lambda match: variables.get(match.group(0), match.group(0)),
            value,
        )
    return value


def format_message(raw_msg: str, template: str = None, port: str = "", variables=None) -> str:
    text = template if template is not None else THIRD_PUSH_SMS_TEMPLATE
    if not str(text or "").strip():
        text = "{msg}"
    return str(apply_vars(str(text), template_vars(raw_msg, port, variables))).strip()
