"""Persistence ACKs for queued events, independent of command execution."""
import asyncio
from contextlib import nullcontext
from dataclasses import dataclass


@dataclass
class PendingCloudEventAck:
    websocket: object
    imei: str
    event_id: str
    is_current: object
    loop: object
    wake: object
    acknowledged: bool = False


def interrupt_cloud_sms_event_drain(state, *, lock_held=False):
    """Invalidate delivery, retaining queued events until an explicit clear."""
    if state is None:
        return
    with nullcontext() if lock_held else state.lock:
        state.delivery_generation += 1
        state.pending_ack = None
        for task in tuple(state.active_ack_tasks):
            try:
                task.get_loop().call_soon_threadsafe(task.cancel)
            except RuntimeError:
                pass


async def stop_cloud_sms_event_drain(state):
    if state is None:
        return
    with state.lock:
        interrupt_cloud_sms_event_drain(state, lock_held=True)
        tasks = set(state.active_ack_tasks)
        futures = tuple(state.drain_futures)
    # Scheduled drains may not have entered their coroutine yet. Join them
    # too, before the cloud thread closes its event loop.
    tasks.update(future if asyncio.isfuture(future) else asyncio.wrap_future(future)
                 for future in futures)
    if tasks:
        await asyncio.gather(*tasks, return_exceptions=True)


def track_cloud_sms_event_drain(state, future):
    if future is None or not callable(getattr(future, "add_done_callback", None)):
        return
    with state.lock:
        state.drain_futures.add(future)

    def finished(done):
        with state.lock:
            state.drain_futures.discard(done)
    future.add_done_callback(finished)


def handle_cloud_sms_event_ack(websocket, data, *, state):
    if state is None or not isinstance(data, dict) or data.get("type") != "device_event_ack":
        return False
    if data.get("ok") is not True:
        return False
    identities = [str(data.get(key) or "").strip() for key in ("imei", "device_imei") if key in data]
    event_ids = [str(data.get(key) or "").strip() for key in ("source_event_id", "event_id") if key in data]
    with state.lock:
        pending = state.pending_ack
        if (
            pending is None or pending.acknowledged
            or websocket is not pending.websocket or not pending.is_current()
            or not identities or any(value != pending.imei for value in identities)
            or not event_ids or any(value != pending.event_id for value in event_ids)
        ):
            return False
        pending.acknowledged = True
        pending.loop.call_soon_threadsafe(pending.wake.set)
        return True


async def send_cloud_event_with_ack(
    websocket, payload, *, state, generation, delivery_generation,
    connection_ready, send_payload, log_error=None,
    ack_timeout=10.0, retry_base=2.0, retry_max=30.0,
):
    event_id = str(payload.get("source_event_id") or "").strip()
    imei = str(payload.get("imei") or "").strip()
    if not event_id or len(event_id.encode("utf-8")) > 64 or not imei:
        # Never silently downgrade a malformed reliable event to ws.send().
        return "error"
    loop = asyncio.get_running_loop()
    task = asyncio.current_task()

    def is_current():
        return (state.generation == generation
                and state.delivery_generation == delivery_generation
                and state.pending_ack is pending and connection_ready())

    pending = PendingCloudEventAck(websocket, imei, event_id, is_current, loop, asyncio.Event())
    with state.lock:
        if state.generation != generation or not connection_ready():
            return "not_connected"
        state.pending_ack = pending
        state.active_ack_tasks.add(task)
    delay = retry_base
    warned = False
    try:
        while True:
            with state.lock:
                if pending.acknowledged:
                    return "sent"
                if not is_current():
                    return "not_connected"
            result = "error"
            try:
                result = await asyncio.wait_for(send_payload(websocket, payload), ack_timeout)
            except Exception:
                # No raw event, device identifier, or transport exception is
                # logged here; retry only the immutable event payload.
                pass
            if result == "sent" and not pending.acknowledged:
                try:
                    await asyncio.wait_for(pending.wake.wait(), ack_timeout)
                except asyncio.TimeoutError:
                    pass
            with state.lock:
                if pending.acknowledged:
                    return "sent"
                if not is_current():
                    return "not_connected"
            if not warned and log_error is not None:
                warned = True
                try:
                    log_error("云端事件未收到保存确认，已保留并退避重试；请确认控制台已更新且存储可用。")
                except Exception:
                    pass
            # This also handles older servers without ACK support: retain
            # the event, use bounded backoff, and never report false success.
            try:
                await asyncio.wait_for(pending.wake.wait(), delay)
            except asyncio.TimeoutError:
                pass
            delay = min(retry_max, delay * 2)
    finally:
        with state.lock:
            if state.pending_ack is pending:
                state.pending_ack = None
            state.active_ack_tasks.discard(task)
