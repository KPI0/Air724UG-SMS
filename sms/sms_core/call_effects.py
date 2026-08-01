from dataclasses import dataclass

from sms_core.status_text import format_connected_status


@dataclass(frozen=True)
class CallEffectResult:
    stop_processing: bool = False


def enqueue_call_push(enqueue_push, message, variables):
    try:
        return enqueue_push(message, event_type="call", variables=variables)
    except TypeError:
        return enqueue_push(message, event_type="call")


def _show_missed_call(missed_call, port_ui, show_missed_call_popup):
    if missed_call is None:
        return False
    caller_num = str(getattr(missed_call, "caller_num", "") or "未知号码")
    port_ui(f"📵 未接来电：{caller_num}", "warning")
    show_missed_call_popup(missed_call)
    return True


def apply_ring_timeout_expired(
    current_port,
    port_ui,
    set_status,
    close_call_popup,
    finish_incoming_call=lambda: None,
    show_missed_call_popup=lambda _missed_call: None,
):
    port_ui("📞 呼叫已取消或未接听", "normal")
    set_status(format_connected_status(current_port), "green")
    missed_call = finish_incoming_call()
    close_call_popup()
    _show_missed_call(missed_call, port_ui, show_missed_call_popup)


def apply_call_answer_result(
    result,
    caller_num,
    restore_answer,
    mark_connected,
    port_ui,
    set_status,
    ui_post,
    set_ring_timeout,
):
    if not result.ok:
        port_ui(f"📞 接听失败：{result.error}", "warning")
        ui_post(restore_answer)
        return False

    port_ui("📞 已发送接听指令 (ATA)", "normal")
    set_status(f"📞 通话中：{caller_num}", "blue")
    set_ring_timeout(-1.0)
    ui_post(mark_connected)
    return True


def apply_call_hangup_result(
    result,
    restore_hangup,
    port_ui,
    ui_post,
    close_call_popup,
):
    if not result.ok:
        port_ui(f"📞 挂断失败：{result.error}", "warning")
        ui_post(restore_hangup)
        return False

    port_ui("📞 已发送挂机指令 (ATH)", "normal")
    close_call_popup()
    return True


def apply_call_decision(
    decision,
    current_port,
    send_hangup,
    enqueue_push,
    port_ui,
    set_status,
    show_call_popup,
    close_call_popup,
    start_incoming_call=lambda _caller_num: True,
    finish_incoming_call=lambda: None,
    show_missed_call_popup=lambda _missed_call: None,
):
    incoming_started = True
    incoming_replaced = False
    replaced_missed_call = None
    if decision.incoming_number:
        start_result = start_incoming_call(decision.incoming_number)
        incoming_started = bool(start_result)
        incoming_replaced = getattr(start_result, "replaced", False) is True
        replaced_missed_call = getattr(start_result, "replaced_missed_call", None)

    if incoming_replaced:
        close_call_popup()
        _show_missed_call(replaced_missed_call, port_ui, show_missed_call_popup)

    if decision.push_message and (not decision.incoming_number or incoming_started):
        caller = decision.incoming_number or decision.blocked_number or decision.show_popup_number
        enqueue_call_push(
            enqueue_push,
            decision.push_message,
            {
                "caller": caller,
                "from": caller,
                "phone": caller,
            },
        )

    if decision.blocked_number:
        port_ui(
            f"🚫 防骚扰拦截：拒接 {decision.blocked_number} ({decision.block_reason})",
            "warning",
        )

    if decision.stop_processing:
        send_hangup()
        return CallEffectResult(stop_processing=True)

    if decision.incoming_number and incoming_started:
        port_ui(f"📞 收到来电：来自 {decision.incoming_number}", "normal")
        set_status(f"🔔 响铃中：{decision.incoming_number}", "blue")
        show_call_popup(decision.show_popup_number)

    if decision.call_ended or decision.hangup_notify:
        if decision.hangup_notify:
            port_ui("📞 语音通话已结束", "normal")
        set_status(format_connected_status(current_port), "green")
        missed_call = finish_incoming_call()
        close_call_popup()
        _show_missed_call(missed_call, port_ui, show_missed_call_popup)

    if decision.connected_number:
        port_ui(f"📞 对方已接听：{decision.connected_number}", "normal")
        set_status(f"📞 通话中：{decision.connected_number}", "blue")

    return CallEffectResult(stop_processing=False)
