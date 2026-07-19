import queue

from sms_core.config_schema import THIRD_PUSH_CALL_TEMPLATE, THIRD_PUSH_SMS_TEMPLATE
from sms_core.third_push import (
    THIRD_PUSH_CHANNEL_LABELS,
    dispatch_push_item,
    push_result_status_message,
)


def _exception_text(prefix, exc):
    detail = str(exc).strip() or exc.__class__.__name__
    return f"{prefix}：{detail}"


def _safe_system_ui(system_ui, message):
    try:
        system_ui(message, "normal")
        return True
    except Exception:
        return False


def _safe_show_result(show_result, ok_channels, fail_infos):
    try:
        show_result(list(ok_channels or []), list(fail_infos or []))
        return True
    except Exception:
        return False


def resolve_push_channels(
    channels,
    *,
    enabled,
    sms_enabled,
    call_enabled,
    configured_channels,
    event_type="sms",
    valid_channels=None,
):
    valid_channels = valid_channels or THIRD_PUSH_CHANNEL_LABELS
    if channels is None:
        if not enabled:
            return []
        if event_type == "sms" and not sms_enabled:
            return []
        if event_type == "call" and not call_enabled:
            return []
        return list(configured_channels or [])
    return [channel for channel in channels if channel in valid_channels]


def build_third_push_payload(
    raw_msg,
    *,
    channels,
    settings,
    template=None,
    variables=None,
    event_type="sms",
    show_success=False,
    show_result=False,
    sms_template=THIRD_PUSH_SMS_TEMPLATE,
    call_template=THIRD_PUSH_CALL_TEMPLATE,
):
    if event_type == "call" and template is None:
        template = call_template
    return {
        "message": str(raw_msg or ""),
        "channels": list(channels or []),
        "settings": dict(settings or {}),
        "template": sms_template if template is None else template,
        "variables": dict(variables or {}),
        "show_success": show_success,
        "show_result": show_result,
    }


def enqueue_third_push_runtime(
    raw_msg,
    *,
    push_queue,
    enabled,
    sms_enabled,
    call_enabled,
    configured_channels,
    current_settings,
    channels=None,
    settings=None,
    template=None,
    variables=None,
    event_type="sms",
    show_success=False,
    show_result=False,
    system_ui=None,
    valid_channels=None,
):
    selected = resolve_push_channels(
        channels,
        enabled=enabled,
        sms_enabled=sms_enabled,
        call_enabled=call_enabled,
        configured_channels=configured_channels,
        event_type=event_type,
        valid_channels=valid_channels,
    )
    if not selected:
        return False

    payload = build_third_push_payload(
        raw_msg,
        channels=selected,
        settings=settings if settings is not None else current_settings,
        template=template,
        variables=variables,
        event_type=event_type,
        show_success=show_success,
        show_result=show_result,
    )

    try:
        push_queue.put_nowait(payload)
        return True
    except queue.Full:
        if system_ui is not None:
            system_ui("📡 三方推送队列已满，本条通知未推送", "normal")
        return False


def third_push_worker_runtime(
    *,
    stop_event,
    push_queue,
    send_channel_func,
    system_ui,
    show_result,
    format_message_func=None,
    dispatch_func=dispatch_push_item,
    result_status_message=push_result_status_message,
    poll_timeout=0.5,
    should_emit_results=None,
):
    should_emit_results = should_emit_results or (lambda: True)
    while not stop_event.is_set():
        try:
            item = push_queue.get(timeout=poll_timeout)
        except queue.Empty:
            continue

        try:
            try:
                result = dispatch_func(
                    item,
                    send_channel_func,
                    format_message_func=format_message_func,
                )
            except Exception as exc:
                fail_message = _exception_text("三方推送任务异常", exc)
                if should_emit_results():
                    _safe_system_ui(system_ui, fail_message)
                    if isinstance(item, dict) and item.get("show_result"):
                        _safe_show_result(show_result, [], [fail_message])
                continue

            if not should_emit_results():
                continue

            try:
                status_message = result_status_message(result)
            except Exception as exc:
                status_message = _exception_text("三方推送结果处理异常", exc)

            try:
                if status_message:
                    _safe_system_ui(system_ui, status_message)
                if getattr(result, "show_result", False):
                    _safe_show_result(
                        show_result,
                        getattr(result, "ok_channels", []),
                        getattr(result, "fail_infos", []),
                    )
            except Exception as exc:
                fail_message = _exception_text("三方推送结果回调异常", exc)
                if should_emit_results():
                    _safe_system_ui(system_ui, fail_message)
        finally:
            try:
                push_queue.task_done()
            except Exception:
                pass
