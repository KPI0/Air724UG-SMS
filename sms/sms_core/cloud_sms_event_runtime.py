import asyncio
import queue
import threading
from dataclasses import dataclass, field


@dataclass
class CloudSmsEventDrainState:
    lock: object = field(default_factory=threading.Lock)
    drain_scheduled: bool = False
    generation: int = 0


def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def _task_done_safely(event_queue):
    try:
        event_queue.task_done()
    except Exception:
        pass


def _close_unawaited_coro(coro):
    close = getattr(coro, "close", None)
    if close is not None:
        close()


def put_sms_event_drop_oldest(event_queue, payload):
    try:
        event_queue.put_nowait(payload)
        return True
    except queue.Full:
        try:
            event_queue.get_nowait()
            _task_done_safely(event_queue)
        except queue.Empty:
            pass
        try:
            event_queue.put_nowait(payload)
            return True
        except queue.Full:
            return False


def clear_cloud_sms_event_state(event_queue, state, *, log_error=None):
    try:
        with state.lock:
            while True:
                try:
                    event_queue.get_nowait()
                except queue.Empty:
                    break
                _task_done_safely(event_queue)
            state.drain_scheduled = False
            state.generation += 1
        return True
    except Exception as exc:
        _safe_log(log_error, f"Clear cloud SMS event queue failed: {exc!r}")
        return False


async def drain_cloud_sms_event_queue(
    ws,
    *,
    event_queue,
    batch_size,
    state,
    is_current_connection,
    is_connected,
    is_authorized,
    send_payload,
    create_task=None,
    log_error=None,
    generation=None,
):
    create_task = create_task or asyncio.create_task
    if generation is None:
        with state.lock:
            generation = state.generation
    sent = 0
    send_failed = False

    try:
        while sent < batch_size:
            if (
                not is_current_connection(ws)
                or not is_connected()
                or not is_authorized()
            ):
                return "not_connected"
            with state.lock:
                if generation != state.generation:
                    return "stale"
                try:
                    payload = event_queue.get_nowait()
                except queue.Empty:
                    return "empty"

            try:
                result = await send_payload(ws, payload)
            except Exception as exc:
                _safe_log(log_error, f"Send queued cloud SMS event failed: {exc!r}")
                result = "error"

            if result != "sent":
                # Preserve the failed payload for the next authorized
                # connection.  If producers filled the bounded queue while
                # this send was in flight, evict the oldest queued item to
                # make room; never silently discard the item that just failed.
                with state.lock:
                    state_is_current = generation == state.generation
                    requeued = bool(
                        state_is_current
                        and put_sms_event_drop_oldest(event_queue, payload)
                    )
                if state_is_current and not requeued:
                    _safe_log(log_error, "Requeue cloud SMS event failed: queue full")
                _task_done_safely(event_queue)
                send_failed = True
                return "error"

            _task_done_safely(event_queue)
            sent += 1
    finally:
        with state.lock:
            state_is_current = generation == state.generation
            should_continue = bool(
                state_is_current
                and not send_failed
                and not event_queue.empty()
                and is_current_connection(ws)
                and is_connected()
                and is_authorized()
            )
            if state_is_current and not should_continue:
                state.drain_scheduled = False

        if should_continue:
            coro = None
            try:
                coro = drain_cloud_sms_event_queue(
                    ws,
                    event_queue=event_queue,
                    batch_size=batch_size,
                    state=state,
                    is_current_connection=is_current_connection,
                    is_connected=is_connected,
                    is_authorized=is_authorized,
                    send_payload=send_payload,
                    create_task=create_task,
                    log_error=log_error,
                    generation=generation,
                )
                create_task(coro)
            except Exception as exc:
                _safe_log(log_error, f"Schedule next cloud SMS event drain failed: {exc!r}")
                if coro is not None:
                    _close_unawaited_coro(coro)
                with state.lock:
                    if generation == state.generation:
                        state.drain_scheduled = False

    return "sent"


def schedule_cloud_sms_event_drain(
    loop,
    ws,
    *,
    state,
    drain_coro_factory,
    run_coroutine_threadsafe=None,
    log_error=None,
):
    run_coroutine_threadsafe = run_coroutine_threadsafe or asyncio.run_coroutine_threadsafe
    with state.lock:
        if state.drain_scheduled:
            return False
        generation = state.generation
        state.drain_scheduled = True

    coro = None
    try:
        coro = drain_coro_factory(ws, generation)
        run_coroutine_threadsafe(coro, loop)
        return True
    except Exception as exc:
        _safe_log(log_error, f"Schedule cloud SMS event drain failed: {exc!r}")
        if coro is not None:
            _close_unawaited_coro(coro)
        with state.lock:
            if generation == state.generation:
                state.drain_scheduled = False
        return False


def enqueue_cloud_sms_event_runtime(
    payload,
    *,
    event_queue,
    can_send,
    loop,
    ws,
    schedule_drain,
    log_error=None,
    state=None,
    is_enabled=None,
):
    if state is None:
        queued = put_sms_event_drop_oldest(event_queue, payload)
    else:
        with state.lock:
            if is_enabled is not None and not is_enabled():
                return "disabled"
            queued = put_sms_event_drop_oldest(event_queue, payload)
    if not queued:
        _safe_log(log_error, "Cloud SMS event queue is full; event was not queued")
        return "queue_full"
    if can_send and loop is not None and ws is not None:
        schedule_drain(loop, ws)
        return "queued"
    return "queued_offline"
