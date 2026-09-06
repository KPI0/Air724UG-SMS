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


def _send_call_state(send_cloud_call_state, phone, phase, reason, direction, call_session_id=""):
    """Send a transient call state while keeping legacy callback compatibility."""
    kwargs = {"direction": direction}
    if call_session_id:
        kwargs["call_session_id"] = str(call_session_id)
    return send_cloud_call_state(phone, phase, reason, **kwargs)


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

    # ATA only confirms that the modem accepted the answer command.  The
    # remote party may still be ringing, so do not mark the call connected or
    # start its duration timer until a real CALL=1/CLCC/CONNECT indication is
    # observed by the serial call state machine.
    port_ui("📞 已发送接听指令 (ATA)，等待对方接通", "normal")
    set_status(f"📞 正在接听：{caller_num}", "blue")
    # Keep the existing ringing deadline while the modem is still negotiating
    # the answer. The call state machine switches it to the connected sentinel
    # only after CALL=1/CLCC/CONNECT is observed.
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
    send_cloud_call_event=lambda *_args, **_kwargs: None,
    send_cloud_call_state=lambda *_args, **_kwargs: None,
    finish_dial_call=lambda _message="": None,
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
        try:
            event_kwargs = {
                "blocked": bool(decision.blocked_number),
                "block_reason": decision.block_reason,
            }
            if getattr(decision, "call_session_id", ""):
                event_kwargs["call_session_id"] = decision.call_session_id
            send_cloud_call_event(caller, decision.push_message, **event_kwargs)
        except Exception:
            pass

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
        if decision.call_ended and decision.end_direction:
            try:
                _send_call_state(
                    send_cloud_call_state,
                    decision.end_number,
                    decision.end_phase or "ended",
                    decision.end_reason,
                    decision.end_direction,
                    getattr(decision, "call_session_id", ""),
                )
            except Exception:
                pass
        end_message = getattr(decision, "end_message", "") or "📞 语音通话已结束"
        if decision.hangup_notify:
            port_ui(end_message, "normal")
        set_status(format_connected_status(current_port), "green")
        if getattr(decision, "outgoing_call_ended", False):
            finish_dial_call(end_message)
        else:
            missed_call = finish_incoming_call()
            close_call_popup()
            _show_missed_call(missed_call, port_ui, show_missed_call_popup)

    if decision.connected_number:
        port_ui(f"📞 对方已接听：{decision.connected_number}", "normal")
        set_status(f"📞 通话中：{decision.connected_number}", "blue")

    if decision.incoming_connected_number:
        try:
            _send_call_state(
                send_cloud_call_state,
                decision.incoming_connected_number,
                "connected",
                "",
                "incoming",
                getattr(decision, "call_session_id", ""),
            )
        except Exception:
            pass

    return CallEffectResult(stop_processing=False)
