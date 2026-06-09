import os
from datetime import datetime


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
        enqueue_third_push(full_msg)
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
