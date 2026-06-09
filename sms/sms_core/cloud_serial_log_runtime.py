import asyncio
import json
import queue
import threading
from dataclasses import dataclass, field


@dataclass
class CloudSerialLogDrainState:
    lock: object = field(default_factory=threading.Lock)
    drain_scheduled: bool = False


def put_drop_oldest(log_queue, payload):
    try:
        log_queue.put_nowait(payload)
        return True
    except queue.Full:
        try:
            log_queue.get_nowait()
        except queue.Empty:
            pass
        try:
            log_queue.put_nowait(payload)
            return True
        except queue.Full:
            return False


def clear_cloud_serial_log_queue(log_queue):
    try:
        while True:
            log_queue.get_nowait()
    except queue.Empty:
        pass
    except Exception:
        pass


def reset_cloud_serial_log_state(log_queue, state):
    clear_cloud_serial_log_queue(log_queue)
    try:
        with state.lock:
            state.drain_scheduled = False
    except Exception:
        pass


async def drain_cloud_serial_log_queue(
    ws,
    *,
    log_queue,
    batch_size,
    state,
    is_current_connection,
    is_connected,
    serialize_payload=None,
    create_task=None,
):
    serialize_payload = serialize_payload or (lambda payload: json.dumps(payload, ensure_ascii=False))
    create_task = create_task or asyncio.create_task
    should_continue = False

    try:
        sent = 0
        while sent < batch_size:
            if not is_current_connection(ws) or not is_connected():
                clear_cloud_serial_log_queue(log_queue)
                return
            try:
                payload = log_queue.get_nowait()
            except queue.Empty:
                return

            await ws.send(serialize_payload(payload))
            sent += 1
    except Exception:
        clear_cloud_serial_log_queue(log_queue)
    finally:
        with state.lock:
            should_continue = (
                not log_queue.empty()
                and is_current_connection(ws)
                and is_connected()
            )
            if not should_continue:
                state.drain_scheduled = False

        if should_continue:
            try:
                create_task(drain_cloud_serial_log_queue(
                    ws,
                    log_queue=log_queue,
                    batch_size=batch_size,
                    state=state,
                    is_current_connection=is_current_connection,
                    is_connected=is_connected,
                    serialize_payload=serialize_payload,
                    create_task=create_task,
                ))
            except Exception:
                with state.lock:
                    state.drain_scheduled = False


def schedule_cloud_serial_log_drain(
    loop,
    ws,
    *,
    state,
    drain_coro_factory,
    run_coroutine_threadsafe=None,
):
    run_coroutine_threadsafe = run_coroutine_threadsafe or asyncio.run_coroutine_threadsafe

    with state.lock:
        if state.drain_scheduled:
            return False
        state.drain_scheduled = True

    try:
        run_coroutine_threadsafe(drain_coro_factory(ws), loop)
        return True
    except Exception:
        with state.lock:
            state.drain_scheduled = False
        return False


def send_cloud_serial_log_runtime(
    line,
    *,
    authorized,
    get_loop,
    get_ws,
    is_connected,
    runtime_imei,
    build_payload,
    log_queue,
    schedule_drain,
):
    if not authorized:
        return "unauthorized"

    try:
        loop = get_loop()
        ws = get_ws()
        if loop is None or not loop.is_running() or ws is None or not is_connected():
            return "not_connected"
        if not runtime_imei():
            return "missing_imei"

        payload = build_payload(line)
        if payload is None:
            return "empty"

        if not put_drop_oldest(log_queue, payload):
            return "queue_full"

        schedule_drain(loop, ws)
        return "queued"
    except Exception:
        return "error"
