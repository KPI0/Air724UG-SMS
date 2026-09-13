import asyncio
import queue
import threading
from dataclasses import dataclass, field

from sms_core.cloud_event_ack_runtime import (
    handle_cloud_sms_event_ack,
    interrupt_cloud_sms_event_drain,
    send_cloud_event_with_ack,
    stop_cloud_sms_event_drain,
    track_cloud_sms_event_drain,
)


@dataclass
class CloudSmsEventDrainState:
    lock: object = field(default_factory=threading.Lock)
    drain_scheduled: bool = False
    generation: int = 0
    delivery_generation: int = 0
    pending_ack: object = None
    active_ack_tasks: set = field(default_factory=set)
    drain_futures: set = field(default_factory=set)


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


def put_sms_event_drop_oldest(event_queue, payload, *, older_items=None):
    if older_items is not None:
        # A failed send must go back before later events, even when it was
        # selected from the middle of this FIFO Queue. Producers may have
        # evicted some of its predecessors while the send was in flight.
        older_ids = {id(item) for item in older_items}
        with event_queue.mutex:
            if event_queue.maxsize > 0 and len(event_queue.queue) >= event_queue.maxsize:
                event_queue.queue.popleft()
                event_queue.unfinished_tasks -= 1
            position = sum(id(item) in older_ids for item in event_queue.queue)
            event_queue.queue.insert(position, payload)
            event_queue.unfinished_tasks += 1
            event_queue.not_empty.notify()
        return True
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


def _event_matches_imei(payload, imei):
    return imei is None or (
        bool(imei)
        and isinstance(payload, dict)
        and str(payload.get("imei") or "").strip() == imei
    )


def has_cloud_sms_event_for_imei(event_queue, imei):
    with event_queue.mutex:
        return any(_event_matches_imei(payload, imei) for payload in event_queue.queue)


def _take_cloud_sms_event_for_imei(event_queue, imei):
    # Select in place: rotating unrelated devices to the tail would change
    # their delivery order and the global drop-oldest policy at a batch edge.
    with event_queue.mutex:
        older_items = []
        for position, payload in enumerate(event_queue.queue):
            if _event_matches_imei(payload, imei):
                del event_queue.queue[position]
                event_queue.not_full.notify()
                return payload, older_items
            older_items.append(payload)
    raise queue.Empty


def clear_cloud_sms_event_state(event_queue, state, *, log_error=None):
    try:
        with state.lock:
            interrupt_cloud_sms_event_drain(state, lock_held=True)
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
        _safe_log(log_error, f"Clear cloud device event queue failed: {exc!r}")
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
    runtime_imei=None,
    drain_imei=None,
    schedule_pending=None,
    ack_timeout=10.0,
    retry_base=2.0,
    retry_max=30.0,
):
    create_task = create_task or asyncio.create_task
    with state.lock:
        if generation is None:
            generation = state.generation
        delivery_generation = state.delivery_generation
        if runtime_imei is not None and drain_imei is None:
            drain_imei = runtime_imei()
    sent = 0
    send_failed = False

    def connection_ready():
        return bool(
            delivery_generation == state.delivery_generation
            and is_current_connection(ws)
            and is_connected()
            and is_authorized()
            and (runtime_imei is None or (drain_imei and runtime_imei() == drain_imei))
        )

    try:
        while sent < batch_size:
            with state.lock:
                if generation != state.generation:
                    return "stale"
                if not connection_ready():
                    return "not_connected"
                try:
                    payload, older_items = _take_cloud_sms_event_for_imei(event_queue, drain_imei)
                except queue.Empty:
                    return "empty"

            result = None
            try:
                if isinstance(payload, dict) and payload.get("ack_required") is True:
                    result = await send_cloud_event_with_ack(
                        ws, payload, state=state, generation=generation,
                        delivery_generation=delivery_generation,
                        connection_ready=connection_ready, send_payload=send_payload,
                        log_error=log_error, ack_timeout=ack_timeout,
                        retry_base=retry_base, retry_max=retry_max,
                    )
                else:
                    result = await send_payload(ws, payload)
            except Exception as exc:
                _safe_log(log_error, f"Send queued cloud device event failed: {exc!r}")
                result = "error"
            finally:
                if result != "sent":
                    # Identity changes retain the queue generation; only an
                    # explicit clear invalidates this in-flight event.
                    with state.lock:
                        state_is_current = generation == state.generation
                        requeued = bool(
                            state_is_current
                            and put_sms_event_drop_oldest(
                                event_queue, payload, older_items=older_items
                            )
                        )
                    if state_is_current and not requeued:
                        _safe_log(log_error, "Requeue cloud device event failed: queue full")
                    send_failed = True
                _task_done_safely(event_queue)

            if result != "sent":
                return "error"

            sent += 1
    finally:
        with state.lock:
            state_is_current = generation == state.generation
            should_continue = bool(
                state_is_current
                and not send_failed
                and connection_ready()
                and has_cloud_sms_event_for_imei(event_queue, drain_imei)
            )
            if state_is_current and not should_continue:
                state.drain_scheduled = False
            changed_connection = bool(
                delivery_generation != state.delivery_generation
                or not is_current_connection(ws)
                or (runtime_imei is not None and runtime_imei() != drain_imei)
            )
            reschedule_current = state_is_current and changed_connection and not should_continue

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
                    runtime_imei=runtime_imei,
                    drain_imei=drain_imei,
                    schedule_pending=schedule_pending,
                    ack_timeout=ack_timeout,
                    retry_base=retry_base,
                    retry_max=retry_max,
                )
                track_cloud_sms_event_drain(state, create_task(coro))
            except Exception as exc:
                _safe_log(log_error, f"Schedule next cloud device event drain failed: {exc!r}")
                if coro is not None:
                    _close_unawaited_coro(coro)
                with state.lock:
                    if generation == state.generation:
                        state.drain_scheduled = False
        elif reschedule_current and schedule_pending is not None:
            # A new ACK may have arrived while this old send was awaited.
            # Release the old slot before scheduling the current connection.
            schedule_pending()

    return "sent"


def schedule_cloud_sms_event_drain(
    loop,
    ws,
    *,
    state,
    drain_coro_factory,
    run_coroutine_threadsafe=None,
    log_error=None,
    runtime_imei=None,
    can_schedule=None,
):
    run_coroutine_threadsafe = run_coroutine_threadsafe or asyncio.run_coroutine_threadsafe
    coro = None
    generation = None
    try:
        with state.lock:
            if can_schedule is not None and not can_schedule():
                return False
            if state.drain_scheduled:
                return False
            generation = state.generation
            drain_imei = runtime_imei() if runtime_imei is not None else None
            state.drain_scheduled = True
            # Shutdown must see either no submission or its published future.
            # Releasing the lock after claiming the slot would let stop return
            # before this producer queues a task on the closing event loop.
            coro = drain_coro_factory(ws, generation, drain_imei)
            future = run_coroutine_threadsafe(coro, loop)
            if callable(getattr(future, "add_done_callback", None)):
                state.drain_futures.add(future)
        # A completed concurrent Future may invoke its callback immediately;
        # install that callback only after releasing the non-reentrant lock.
        track_cloud_sms_event_drain(state, future)
        return True
    except Exception as exc:
        _safe_log(log_error, f"Schedule cloud device event drain failed: {exc!r}")
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
        _safe_log(log_error, "Cloud device event queue is full; event was not queued")
        return "queue_full"
    if can_send and loop is not None and ws is not None:
        schedule_drain(loop, ws)
        return "queued"
    return "queued_offline"
