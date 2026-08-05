import hashlib
import os
import re
from datetime import datetime


SMS_CALLBACK_META_RE = re.compile(
    r"^\s*(?P<sender>[^\r\n]+?)\s+"
    r"(?P<year>\d{2})/(?P<month>\d{2})/(?P<day>\d{2}),"
    r"(?P<hour>\d{2}):(?P<minute>\d{2}):(?P<second>\d{2})"
    r"(?P<tz_sign>[+-])(?P<tz>\d+)"
)
SMS_CALLBACK_SENDER_RE = re.compile(r"^\s*(?P<sender>\+?\d+)\b")


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
            "sms_time_identity": "",
            "sms_time_valid": False,
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
        sms_time_identity = (
            f"{sms_time}{match.group('tz_sign')}{match.group('tz')}"
        )
        sms_time_valid = True
    except Exception:
        sms_time = fallback_time
        sms_time_identity = ""
        sms_time_valid = False

    return {
        "sender": sender,
        "from": sender,
        "phone": sender,
        "local_number": "",
        "self_number": "",
        "sms_time": sms_time,
        "sms_time_identity": sms_time_identity,
        "sms_time_valid": sms_time_valid,
    }


def build_sms_ui_display_lines(callback_head: str, full_msg: str, now=None):
    text = str(callback_head or "").strip()
    metadata = parse_sms_callback_metadata(text, now=now)
    sender = str(metadata.get("sender") or "").strip()
    sms_time = str(metadata.get("sms_time") or "").strip()

    if not SMS_CALLBACK_META_RE.search(text):
        sender_match = SMS_CALLBACK_SENDER_RE.search(text)
        sender = sender_match.group("sender") if sender_match else sender
        sms_time = ""

    body_lines = _sms_display_body_lines(full_msg)
    if len(body_lines) <= 1:
        content_lines = [f"内容：{body_lines[0] if body_lines else ''}"]
    else:
        content_lines = ["内容："] + body_lines

    return [
        f"号码：{sender or '未知号码'}",
        f"时间：{sms_time or '未知时间'}",
        *content_lines,
    ]


def _sms_display_body_lines(full_msg: str):
    text = str(full_msg or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in text.split("\n")]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return lines


def sms_message_id(
    sender: str,
    sms_time: str,
    body: str,
    concat_reference=None,
    concat_reference_bits=None,
    concat_total=None,
) -> str:
    if concat_reference is None:
        text = "\n".join([str(sender or ""), str(sms_time or ""), str(body or "")])
    else:
        concat_key = ":".join([
            str(concat_reference_bits or 8),
            str(concat_reference),
            str(concat_total or ""),
        ])
        text = "\n".join([
            str(sender or ""),
            str(sms_time or ""),
            concat_key,
            str(body or ""),
        ])
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


def _mark_sms_message_seen(state: dict, message_id: str, limit: int = 500) -> bool:
    if not isinstance(state, dict) or not message_id:
        return False

    seen = state.setdefault("_sms_message_ids", set())
    order = state.setdefault("_sms_message_id_order", [])
    if message_id in seen:
        return True

    seen.add(message_id)
    order.append(message_id)
    while len(order) > limit:
        old = order.pop(0)
        seen.discard(old)
    return False


def enqueue_third_push_with_variables(enqueue_third_push, full_msg, variables):
    try:
        return enqueue_third_push(full_msg, variables=variables)
    except TypeError:
        return enqueue_third_push(full_msg)


def send_cloud_sms_event_with_metadata(send_cloud_sms_event, callback_head, full_msg, metadata):
    if metadata:
        try:
            return send_cloud_sms_event(callback_head, full_msg, metadata=metadata)
        except TypeError:
            return send_cloud_sms_event(callback_head, full_msg)
    return send_cloud_sms_event(callback_head, full_msg)


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
        variables = parse_sms_callback_metadata(pending.callback_head)
        message_trace_id = str(getattr(pending, "message_trace_id", "") or "").strip()
        message_id = sms_message_id(
            variables.get("sender"),
            variables.get("sms_time_identity") or variables.get("sms_time"),
            full_msg,
            getattr(pending, "concat_reference", None),
            getattr(pending, "concat_reference_bits", None),
            getattr(pending, "concat_total", None),
        )
        if message_trace_id and _mark_sms_message_seen(
            repeat_state,
            f"trace:{message_trace_id}:{message_id}",
        ):
            return "duplicate"
        variables["message_id"] = message_id
        if message_trace_id:
            variables["message_trace_id"] = message_trace_id
        enqueue_third_push_with_variables(
            enqueue_third_push,
            full_msg,
            variables,
        )
        cloud_metadata = {}
        if variables.get("sms_time_valid"):
            cloud_metadata["sms_time"] = variables.get("sms_time")
            cloud_metadata["sms_time_identity"] = variables.get("sms_time_identity")
        if message_trace_id:
            cloud_metadata["message_trace_id"] = message_trace_id
        send_cloud_sms_event_with_metadata(
            send_cloud_sms_event,
            pending.callback_head,
            full_msg,
            cloud_metadata,
        )

    if full_msg and sms_keyword_hit(full_msg, keywords):
        port_ui("📩 收到短信：", "normal")
        for line in build_sms_ui_display_lines(pending.callback_head, full_msg):
            port_ui(line, "sms")

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
