import re
from dataclasses import dataclass, replace

from sms_core.serial_parsers import (
    evaluate_call_filter,
    is_clip_line,
    is_call_connected_event,
    is_hangup_event,
    is_new_clip,
    is_ring_line,
    parse_clip_number,
)


@dataclass
class ClipDecision:
    caller_num: str
    blocked: bool
    block_reason: str
    new_clip: bool
    last_clip_num: str
    last_clip_time: float
    ring_timeout_target: float


@dataclass
class HangupDecision:
    matched: bool
    should_notify: bool
    ring_timeout_target: float
    current_dial_num: str
    last_clip_num: str
    last_hangup_time: float


@dataclass
class CallState:
    ring_timeout_target: float = 0.0
    current_dial_num: str = ""
    last_clip_num: str = ""
    last_clip_time: float = 0.0
    last_hangup_time: float = 0.0
    # A desktop connection must identify each modem call generation just like
    # the firmware channel does.  Without this, a late terminal frame from a
    # previous same-number call can be mistaken for the current call by the
    # web console when both channels are connected.
    call_session_id: str = ""
    call_session_sequence: int = 0


@dataclass
class CallLineDecision:
    state: CallState
    push_message: str = ""
    blocked_number: str = ""
    block_reason: str = ""
    incoming_number: str = ""
    show_popup_number: str = ""
    call_ended: bool = False
    hangup_notify: bool = False
    connected_number: str = ""
    incoming_connected_number: str = ""
    outgoing_call_ended: bool = False
    end_message: str = ""
    end_direction: str = ""
    end_number: str = ""
    end_phase: str = ""
    end_reason: str = ""
    call_session_id: str = ""
    stop_processing: bool = False


def _begin_call_session(state: CallState, direction: str, number: str, now: float):
    """Start a new local call generation and return its stable identifier."""
    state.call_session_sequence = int(state.call_session_sequence or 0) + 1
    try:
        stamp = int(float(now) * 1000)
    except (TypeError, ValueError, OverflowError):
        stamp = 0
    # Keep the session token opaque to the phone number.  The number is already
    # carried by the event payload and embedding it here would duplicate
    # sensitive data in logs and WebSocket traces without improving matching.
    state.call_session_id = f"{direction}:{stamp}:{state.call_session_sequence}"
    return state.call_session_id


def handle_clip_line(
    line,
    last_clip_num,
    last_clip_time,
    now,
    filter_mode,
    whitelist,
    blacklist,
    ring_timeout_seconds=12.0,
):
    caller_num = parse_clip_number(line)
    blocked, block_reason = evaluate_call_filter(
        caller_num,
        filter_mode,
        whitelist,
        blacklist,
    )
    new_clip = is_new_clip(caller_num, last_clip_num, now, last_clip_time)

    next_last_clip_num = last_clip_num
    next_last_clip_time = last_clip_time
    if new_clip:
        next_last_clip_num = caller_num
        next_last_clip_time = now

    ring_timeout_target = 0.0 if blocked else now + ring_timeout_seconds
    return ClipDecision(
        caller_num=caller_num,
        blocked=blocked,
        block_reason=block_reason,
        new_clip=new_clip,
        last_clip_num=next_last_clip_num,
        last_clip_time=next_last_clip_time,
        ring_timeout_target=ring_timeout_target,
    )


def call_push_message(caller_num, blocked=False, block_reason=""):
    if blocked:
        return f"收到来电：来自 {caller_num}（已拦截：{block_reason}）"
    return f"收到来电：来自 {caller_num}"


def refresh_ring_timeout(line, current_target, now, ring_timeout_seconds=12.0):
    if not is_ring_line(line):
        return current_target
    if current_target == -1.0:
        return current_target
    return now + ring_timeout_seconds


def ring_timeout_expired(ring_timeout_target, now):
    return ring_timeout_target > 0 and now > ring_timeout_target


def line_confirms_call_presence(line):
    return (
        is_ring_line(line)
        or is_clip_line(line)
        or is_call_connected_event(line)
    )


def is_call_active(ring_timeout_target, current_dial_num, popup_active):
    return ring_timeout_target != 0.0 or bool(current_dial_num) or bool(popup_active)


def handle_hangup_line(
    line,
    ring_timeout_target,
    current_dial_num,
    popup_active,
    last_clip_num,
    last_hangup_time,
    now,
    debounce_seconds=3.0,
):
    if not is_hangup_event(line) or not is_call_active(ring_timeout_target, current_dial_num, popup_active):
        return HangupDecision(
            matched=False,
            should_notify=False,
            ring_timeout_target=ring_timeout_target,
            current_dial_num=current_dial_num,
            last_clip_num=last_clip_num,
            last_hangup_time=last_hangup_time,
        )

    should_notify = now - last_hangup_time > debounce_seconds
    return HangupDecision(
        matched=True,
        should_notify=should_notify,
        ring_timeout_target=0.0,
        current_dial_num="",
        last_clip_num="",
        last_hangup_time=now if should_notify else last_hangup_time,
    )


def connected_call_number(line, current_dial_num):
    if is_call_connected_event(line) and current_dial_num:
        return current_dial_num
    return ""


def call_end_message(line):
    """Return a user-facing reason for an outbound call terminal response."""
    text = str(line or "").upper()
    if "BUSY" in text:
        return "📞 对方忙线"
    if "NO ANSWER" in text:
        return "📞 对方未接听"
    return "📞 对方已挂断"


def call_end_state(line):
    text = str(line or "").upper()
    if re.search(r"(?:^|\s)\+(?:CME|CMS)\s+ERROR\b", text):
        return "failed", text.strip()
    if "BUSY" in text:
        return "busy", "BUSY"
    if "NO ANSWER" in text:
        return "no_answer", "NO ANSWER"
    if "NO CARRIER" in text:
        return "ended", "NO CARRIER"
    if "CALL" in text and ",0" in text.replace(" ", ""):
        return "ended", "CALL=0"
    return "ended", ""


def is_call_dial_failure_event(line):
    """Return true for an explicit Modem rejection of the current ATD."""
    return bool(re.search(r"(?:^|\s)\+(?:CME|CMS)\s+ERROR\b", str(line or "").upper()))


def handle_call_line(
    line,
    state,
    now,
    filter_mode,
    whitelist,
    blacklist,
    popup_active,
):
    next_state = CallState(
        ring_timeout_target=state.ring_timeout_target,
        current_dial_num=state.current_dial_num,
        last_clip_num=state.last_clip_num,
        last_clip_time=state.last_clip_time,
        last_hangup_time=state.last_hangup_time,
        call_session_id=state.call_session_id,
        call_session_sequence=state.call_session_sequence,
    )
    decision = CallLineDecision(state=next_state)

    # The serial UI stores the dial number before writing ATD.  Establish the
    # outgoing generation at the first modem line so CONNECT/NO CARRIER cannot
    # be associated with a previous incoming or same-number session.
    if next_state.current_dial_num and not next_state.call_session_id.startswith("outgoing:"):
        _begin_call_session(next_state, "outgoing", next_state.current_dial_num, now)

    if next_state.current_dial_num and is_call_dial_failure_event(line):
        decision.call_session_id = next_state.call_session_id
        decision.call_ended = True
        decision.outgoing_call_ended = True
        decision.end_direction = "outgoing"
        decision.end_number = next_state.current_dial_num
        decision.end_phase, decision.end_reason = call_end_state(line)
        decision.end_message = "📞 拨号失败：" + str(line or "").strip()
        next_state.current_dial_num = ""
        next_state.call_session_id = ""
        return decision

    if is_clip_line(line):
        clip = handle_clip_line(
            line,
            next_state.last_clip_num,
            next_state.last_clip_time,
            now,
            filter_mode,
            whitelist,
            blacklist,
        )
        if (
            not clip.blocked
            and state.ring_timeout_target == -1.0
            and clip.caller_num == state.last_clip_num
        ):
            # Once this caller is connected, a late/repeated +CLIP is still
            # the same call even if it arrives outside the normal duplicate
            # debounce window.
            clip = replace(
                clip,
                new_clip=False,
                last_clip_num=state.last_clip_num,
                last_clip_time=state.last_clip_time,
                ring_timeout_target=-1.0,
            )
        if not clip.new_clip and state.ring_timeout_target == -1.0:
            # Some basebands repeat +CLIP after ATA/connection.  Do not turn
            # an already connected call back into a ringing timeout window.
            next_state.ring_timeout_target = -1.0
        else:
            next_state.ring_timeout_target = clip.ring_timeout_target

        if clip.new_clip:
            next_state.last_clip_num = clip.last_clip_num
            next_state.last_clip_time = clip.last_clip_time
            next_state.last_hangup_time = 0.0
            _begin_call_session(next_state, "incoming", clip.caller_num, now)
            decision.call_session_id = next_state.call_session_id
            decision.push_message = call_push_message(clip.caller_num, clip.blocked, clip.block_reason)

        if clip.blocked:
            if clip.new_clip:
                decision.blocked_number = clip.caller_num
                decision.block_reason = clip.block_reason
            decision.stop_processing = True
            return decision

        if clip.new_clip:
            decision.incoming_number = clip.caller_num
            decision.show_popup_number = clip.caller_num
    else:
        next_state.ring_timeout_target = refresh_ring_timeout(
            line,
            next_state.ring_timeout_target,
            now,
        )

    hangup = handle_hangup_line(
        line,
        next_state.ring_timeout_target,
        next_state.current_dial_num,
        popup_active,
        next_state.last_clip_num,
        next_state.last_hangup_time,
        now,
    )
    if hangup.matched:
        decision.call_session_id = next_state.call_session_id
        decision.outgoing_call_ended = bool(next_state.current_dial_num)
        decision.end_direction = "outgoing" if decision.outgoing_call_ended else "incoming"
        decision.end_number = next_state.current_dial_num or next_state.last_clip_num
        decision.end_phase, decision.end_reason = call_end_state(line)
        if decision.outgoing_call_ended:
            decision.end_message = call_end_message(line)
        decision.call_ended = True
        next_state.ring_timeout_target = hangup.ring_timeout_target
        next_state.current_dial_num = hangup.current_dial_num
        next_state.last_clip_num = hangup.last_clip_num
        next_state.call_session_id = ""
        if hangup.should_notify:
            decision.hangup_notify = True
            next_state.last_hangup_time = hangup.last_hangup_time

    connected_num = connected_call_number(line, next_state.current_dial_num)
    if connected_num:
        decision.call_session_id = next_state.call_session_id
        decision.connected_number = connected_num
        next_state.ring_timeout_target = 0.0
    elif is_call_connected_event(line) and next_state.last_clip_num:
        decision.call_session_id = next_state.call_session_id
        decision.incoming_connected_number = next_state.last_clip_num
        next_state.ring_timeout_target = -1.0

    return decision
