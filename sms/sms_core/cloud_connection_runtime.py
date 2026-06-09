import json


def serialize_cloud_payload(payload):
    return json.dumps(payload, ensure_ascii=False)


async def send_cloud_payload_runtime(ws, payload, *, serialize_payload=serialize_cloud_payload):
    try:
        await ws.send(serialize_payload(payload))
        return "sent"
    except Exception:
        return "error"


def schedule_cloud_unregister_runtime(
    *,
    reason,
    get_loop,
    get_ws,
    is_connected,
    send_unregister,
    run_coroutine_threadsafe,
):
    try:
        loop = get_loop()
        ws = get_ws()
        if loop is not None and loop.is_running() and ws is not None and is_connected():
            run_coroutine_threadsafe(send_unregister(ws, reason), loop)
            return True
    except Exception:
        pass
    return False


async def unregister_then_close_cloud_connection_runtime(
    ws,
    *,
    reason,
    auto_upload,
    send_unregister,
):
    unregistered = False
    closed = False
    try:
        if auto_upload:
            await send_unregister(ws, reason)
            unregistered = True
    except Exception:
        pass

    try:
        await ws.close()
        closed = True
    except Exception:
        pass

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
