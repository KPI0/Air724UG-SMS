import os
import re
from datetime import datetime


SMS_CALLBACK_META_RE = re.compile(
    r"^\s*(?P<sender>\+?\d+)\s+"
    r"(?P<year>\d{2})/(?P<month>\d{2})/(?P<day>\d{2}),"
    r"(?P<hour>\d{2}):(?P<minute>\d{2}):(?P<second>\d{2})\+\d+"
)


def sms_keyword_hit(full_msg: str, keywords) -> bool:
    if not keywords:
        return True
    msg_lower = str(full_msg or "").lower()
    return any(keyword and str(keyword).lower() in msg_lower for keyword in keywords)


def build_unmatched_sms_log_entries(log_dir, log_prefix, full_msg, now=None):
    body = str(full_msg or "")
    if not body:
        return []

    current = now or datetime.now()
    today = current.strftime("%Y-%m-%d")
    time_prefix = current.strftime("%Y-%m-%d %H:%M:%S")
    prefix = str(log_prefix or "system").replace(":", "_")
    path = os.path.join(log_dir, f"sms_{prefix}_{today}.txt")
    return [
        (path, f"{time_prefix} 🚫 [未匹配拦截] 📩 收到短信：\n"),
        (path, f"{time_prefix} {body}\n"),
    ]


def repeat_count_message(state: dict, key: str, message: str, limit: int, suppressed_message=None):
    last_key = state.get("key")
    count = int(state.get("count") or 0)
    if last_key == key:
        count += 1
    else:
        last_key = key
        count = 1

    state["key"] = last_key
    state["count"] = count

    if count < limit:
        return message
    if count == limit:
        return suppressed_message or f"{message}（后续同类消息已忽略）"
    return None


def sms_year_from_short_year(value: str) -> int:
    year = int(value)
    return 1900 + year if year >= 70 else 2000 + year


def parse_sms_callback_metadata(callback_head: str, now=None):
    text = str(callback_head or "").strip()
    match = SMS_CALLBACK_META_RE.search(text)
    fallback_time = (now or datetime.now()).strftime("%Y-%m-%d %H:%M:%S")
    if not match:
        return {
            "sender": "",
            "from": "",
            "phone": "",
            "local_number": "",
            "self_number": "",
            "sms_time": fallback_time,
        }

    sender = match.group("sender")
    try:
        sms_time = datetime(
            sms_year_from_short_year(match.group("year")),
            int(match.group("month")),
            int(match.group("day")),
            int(match.group("hour")),
            int(match.group("minute")),
            int(match.group("second")),
        ).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        sms_time = fallback_time

    return {
        "sender": sender,
        "from": sender,
        "phone": sender,
        "local_number": "",
        "self_number": "",
        "sms_time": sms_time,
    }


def enqueue_third_push_with_variables(enqueue_third_push, full_msg, variables):
    try:
        return enqueue_third_push(full_msg, variables=variables)
    except TypeError:
        return enqueue_third_push(full_msg)


def process_pending_sms(
    pending,
    keywords,
    log_unmatched_sms,
    log_dir,
    log_prefix,
    repeat_state,
    repeat_limit,
    enqueue_third_push,
    send_cloud_sms_event,
    port_ui,
    play_alert,
    show_sms_popup,
    file_log_put,
    system_ui,
):
    if pending is None:
        return "empty"

    full_msg = pending.full_msg

    if full_msg:
        enqueue_third_push_with_variables(
            enqueue_third_push,
            full_msg,
            parse_sms_callback_metadata(pending.callback_head),
        )
        send_cloud_sms_event(pending.callback_head, full_msg)

    if full_msg and sms_keyword_hit(full_msg, keywords):
        if pending.display_lines:
            first = True
            for line in pending.display_lines:
                if first:
                    port_ui(line, "normal")
                    first = False
                else:
                    port_ui(line, "sms")
        else:
            port_ui("📩 收到短信：", "normal")
            port_ui(full_msg, "sms")

        play_alert()
        show_sms_popup(full_msg)
        return "shown"

    if log_unmatched_sms and full_msg:
        try:
            for item in build_unmatched_sms_log_entries(log_dir, log_prefix, full_msg):
                file_log_put(item)
        except Exception:
            pass

    try:
        msg_text = "🚫 短信未命中关键词，已忽略"
        ui_msg = repeat_count_message(repeat_state, msg_text, msg_text, repeat_limit)
        if ui_msg is not None:
            system_ui(ui_msg, "normal")
    except Exception:
        try:
            system_ui("🚫 短信未命中关键词，已忽略", "normal")
        except Exception:
            pass

    return "ignored"
