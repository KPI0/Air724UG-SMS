from dataclasses import dataclass

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
    stop_processing: bool = False


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
    )
    decision = CallLineDecision(state=next_state)

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
        next_state.ring_timeout_target = clip.ring_timeout_target

        if clip.new_clip:
            next_state.last_clip_num = clip.last_clip_num
            next_state.last_clip_time = clip.last_clip_time
            next_state.last_hangup_time = 0.0
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
        decision.call_ended = True
        next_state.ring_timeout_target = hangup.ring_timeout_target
        next_state.current_dial_num = hangup.current_dial_num
        next_state.last_clip_num = hangup.last_clip_num
        if hangup.should_notify:
            decision.hangup_notify = True
            next_state.last_hangup_time = hangup.last_hangup_time

    connected_num = connected_call_number(line, next_state.current_dial_num)
    if connected_num:
        decision.connected_number = connected_num
        next_state.ring_timeout_target = 0.0

    return decision
