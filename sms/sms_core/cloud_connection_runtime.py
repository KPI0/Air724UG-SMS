import json


def serialize_cloud_payload(payload):
    return json.dumps(payload, ensure_ascii=False)


def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def _close_unawaited_coro(coro):
    close = getattr(coro, "close", None)
    if close is not None:
        close()


async def send_cloud_payload_runtime(
    ws,
    payload,
    *,
    serialize_payload=serialize_cloud_payload,
    log_error=None,
):
    try:
        await ws.send(serialize_payload(payload))
        return "sent"
    except Exception as exc:
        _safe_log(log_error, f"Send cloud payload failed: {exc!r}")
        return "error"


def schedule_cloud_unregister_runtime(
    *,
    reason,
    get_loop,
    get_ws,
    is_connected,
    send_unregister,
    run_coroutine_threadsafe,
    log_error=None,
):
    coro = None
    try:
        loop = get_loop()
        ws = get_ws()
        if loop is not None and loop.is_running() and ws is not None and is_connected():
            coro = send_unregister(ws, reason)
            run_coroutine_threadsafe(coro, loop)
            return True
    except Exception as exc:
        _safe_log(log_error, f"Schedule cloud unregister failed: {exc!r}")
        if coro is not None:
            _close_unawaited_coro(coro)
    return False


async def unregister_then_close_cloud_connection_runtime(
    ws,
    *,
    reason,
    auto_upload,
    send_unregister,
    send_session_revoke=None,
    log_error=None,
):
    unregistered = False
    closed = False
    try:
        if auto_upload:
            await send_unregister(ws, reason)
            unregistered = True
    except Exception as exc:
        _safe_log(log_error, f"Send cloud unregister failed: {exc!r}")

    if send_session_revoke is not None:
        try:
            await send_session_revoke(ws, reason)
        except Exception as exc:
            _safe_log(log_error, f"Revoke cloud device session failed: {exc!r}")

    try:
        await ws.close()
        closed = True
    except Exception as exc:
        _safe_log(log_error, f"Close cloud websocket failed: {exc!r}")

    return unregistered, closed


async def reply_cloud_payload_runtime(
    ws,
    payload,
    *,
    identity_payload,
    log,
    serialize_payload=serialize_cloud_payload,
):
    try:
        if isinstance(payload, dict):
            payload = {**identity_payload(), **payload}
        await ws.send(serialize_payload(payload))
        return "sent"
    except Exception as exc:
        log(f"回复云端失败：{exc}")
        return "error"
