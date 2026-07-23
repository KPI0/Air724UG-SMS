import asyncio
import json
import queue
import threading
from dataclasses import dataclass, field


@dataclass
class CloudSerialLogDrainState:
    lock: object = field(default_factory=threading.Lock)
    drain_scheduled: bool = False


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


def _task_done_safely(log_queue):
    task_done = getattr(log_queue, "task_done", None)
    if task_done is None:
        return
    try:
        task_done()
    except Exception:
        pass


def put_drop_oldest(log_queue, payload):
    try:
        log_queue.put_nowait(payload)
        return True
    except queue.Full:
        try:
            log_queue.get_nowait()
            _task_done_safely(log_queue)
        except queue.Empty:
            pass
        try:
            log_queue.put_nowait(payload)
            return True
        except queue.Full:
            return False


def clear_cloud_serial_log_queue(log_queue, *, log_error=None):
    try:
        while True:
            try:
                log_queue.get_nowait()
            except queue.Empty:
                break
            _task_done_safely(log_queue)
    except Exception as exc:
        _safe_log(log_error, f"Clear cloud serial log queue failed: {exc!r}")


def reset_cloud_serial_log_state(log_queue, state, *, log_error=None):
    clear_cloud_serial_log_queue(log_queue, log_error=log_error)
    try:
        with state.lock:
            state.drain_scheduled = False
    except Exception as exc:
        _safe_log(log_error, f"Reset cloud serial log state failed: {exc!r}")


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
    log_error=None,
):
    serialize_payload = serialize_payload or (lambda payload: json.dumps(payload, ensure_ascii=False))
    create_task = create_task or asyncio.create_task
    should_continue = False

    try:
        sent = 0
        while sent < batch_size:
            if not is_current_connection(ws) or not is_connected():
                clear_cloud_serial_log_queue(log_queue, log_error=log_error)
                return
            try:
                payload = log_queue.get_nowait()
            except queue.Empty:
                return

            try:
                await ws.send(serialize_payload(payload))
                sent += 1
            finally:
                _task_done_safely(log_queue)
    except Exception as exc:
        _safe_log(log_error, f"Drain cloud serial log queue failed: {exc!r}")
        clear_cloud_serial_log_queue(log_queue, log_error=log_error)
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
            coro = None
            try:
                coro = drain_cloud_serial_log_queue(
                    ws,
                    log_queue=log_queue,
                    batch_size=batch_size,
                    state=state,
                    is_current_connection=is_current_connection,
                    is_connected=is_connected,
                    serialize_payload=serialize_payload,
                    create_task=create_task,
                    log_error=log_error,
                )
                create_task(coro)
            except Exception as exc:
                _safe_log(log_error, f"Schedule next cloud serial log drain failed: {exc!r}")
                if coro is not None:
                    _close_unawaited_coro(coro)
                with state.lock:
                    state.drain_scheduled = False


def schedule_cloud_serial_log_drain(
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
        state.drain_scheduled = True

    coro = None
    try:
        coro = drain_coro_factory(ws)
        run_coroutine_threadsafe(coro, loop)
        return True
    except Exception as exc:
        _safe_log(log_error, f"Schedule cloud serial log drain failed: {exc!r}")
        if coro is not None:
            _close_unawaited_coro(coro)
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
    log_error=None,
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
    except Exception as exc:
        _safe_log(log_error, f"Send cloud serial log failed: {exc!r}")
        return "error"
