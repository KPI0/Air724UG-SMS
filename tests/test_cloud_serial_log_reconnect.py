"""A delayed producer/drain must not consume another connection's logs."""
import asyncio
import json
import queue
import threading
import unittest
from types import SimpleNamespace

from sms_core.cloud_serial_log_runtime import (
    CloudSerialLogDrainState,
    drain_cloud_serial_log_queue,
    reset_cloud_serial_log_state,
    schedule_cloud_serial_log_drain,
    send_cloud_serial_log_runtime,
)


class CloudSerialLogReconnectTests(unittest.IsolatedAsyncioTestCase):
    async def test_preempted_submission_cannot_clear_new_connection_logs(self):
        loop = asyncio.get_running_loop()
        log_queue = queue.Queue(maxsize=4)
        state = CloudSerialLogDrainState()
        entered = threading.Event()
        resume = threading.Event()
        send_started = asyncio.Event()
        send_resume = asyncio.Event()
        submitted = {}

        class Socket:
            def __init__(self, gated=False):
                self.gated = gated
                self.sent = []

            async def send(self, payload):
                if self.gated and not self.sent:
                    send_started.set()
                    await send_resume.wait()
                self.sent.append(json.loads(payload))

        old_ws, new_ws = Socket(), Socket(gated=True)
        current = [old_ws]

        def drain(ws, generation):
            return drain_cloud_serial_log_queue(
                ws, log_queue=log_queue, batch_size=100, state=state,
                is_current_connection=lambda candidate: candidate is current[0],
                is_connected=lambda: True, generation=generation,
            )

        def submit(coro, target_loop):
            old_producer = threading.current_thread() is worker
            if old_producer:
                entered.set()
                if not resume.wait(3):
                    raise TimeoutError("producer submission gate")
            future = asyncio.run_coroutine_threadsafe(coro, target_loop)
            submitted["old" if old_producer else "new"] = future
            return future

        def schedule(target_loop, ws, *, generation):
            return schedule_cloud_serial_log_drain(
                target_loop, ws, state=state, drain_coro_factory=drain,
                run_coroutine_threadsafe=submit, generation=generation,
            )

        def produce(line):
            return send_cloud_serial_log_runtime(
                line, authorized=True, get_loop=lambda: loop, get_ws=lambda: current[0],
                is_connected=lambda: True, runtime_imei=lambda: "000000000000001",
                build_payload=lambda value: {"type": "log", "data": value},
                log_queue=log_queue, schedule_drain=schedule, state=state,
            )

        worker = threading.Thread(target=lambda: produce("old"), daemon=True)
        worker.start()
        try:
            self.assertTrue(await asyncio.to_thread(entered.wait, 3))
            current[0] = new_ws
            reset_cloud_serial_log_state(log_queue, state)
            self.assertEqual(produce("new-1"), "queued")
            await asyncio.wait_for(send_started.wait(), 3)
            self.assertEqual(produce("new-2"), "queued")
            self.assertEqual(log_queue.qsize(), 1)
            resume.set()
            await asyncio.to_thread(worker.join, 3)
            self.assertFalse(worker.is_alive())
            await asyncio.wait_for(asyncio.wrap_future(submitted["old"]), 3)
            self.assertEqual(log_queue.qsize(), 1)
            self.assertTrue(state.drain_scheduled)
            send_resume.set()
            await asyncio.wait_for(asyncio.wrap_future(submitted["new"]), 3)
            self.assertEqual([item["data"] for item in new_ws.sent], ["new-1", "new-2"])
            self.assertEqual(old_ws.sent, [])
            self.assertEqual(log_queue.unfinished_tasks, 0)
            self.assertFalse(state.drain_scheduled)
        finally:
            resume.set()
            send_resume.set()
            await asyncio.to_thread(worker.join, 3)

    async def test_old_send_failure_preserves_new_generation(self):
        for fail in (False, True):
            with self.subTest(fail=fail):
                log_queue = queue.Queue(maxsize=2)
                state = CloudSerialLogDrainState(drain_scheduled=True)
                log_queue.put_nowait({"id": "old"})
                current = []

                class Socket:
                    async def send(self, _payload):
                        reset_cloud_serial_log_state(log_queue, state)
                        current[:] = [object()]
                        log_queue.put_nowait({"id": "new"})
                        state.drain_scheduled = True
                        if fail:
                            raise RuntimeError("old send failed")

                ws = Socket()
                current[:] = [ws]
                await drain_cloud_serial_log_queue(
                    ws, log_queue=log_queue, batch_size=2, state=state,
                    is_current_connection=lambda candidate: candidate is current[0],
                    is_connected=lambda: True,
                )
                self.assertEqual(list(log_queue.queue), [{"id": "new"}])
                self.assertEqual(log_queue.unfinished_tasks, 1)
                self.assertTrue(state.drain_scheduled)

    async def test_reset_invalidates_queued_continuation_on_same_socket(self):
        log_queue = queue.Queue(maxsize=2)
        state = CloudSerialLogDrainState(drain_scheduled=True)
        log_queue.put_nowait({"id": 1})
        log_queue.put_nowait({"id": 2})
        sent, scheduled = [], []

        async def send(payload):
            sent.append(payload)

        ws = SimpleNamespace(send=send)
        await drain_cloud_serial_log_queue(
            ws, log_queue=log_queue, batch_size=1, state=state,
            is_current_connection=lambda candidate: candidate is ws,
            is_connected=lambda: True, create_task=scheduled.append,
        )
        self.assertEqual(len(scheduled), 1)
        reset_cloud_serial_log_state(log_queue, state)
        log_queue.put_nowait({"id": "new"})
        state.drain_scheduled = True
        await scheduled[0]
        self.assertEqual(len(sent), 1)
        self.assertEqual(list(log_queue.queue), [{"id": "new"}])
        self.assertEqual(log_queue.unfinished_tasks, 1)
        self.assertTrue(state.drain_scheduled)

    async def test_old_scheduler_failure_cannot_clear_new_flag(self):
        for failure_point in ("factory", "submit", "continuation"):
            with self.subTest(failure_point=failure_point):
                log_queue = queue.Queue(maxsize=2)
                state = CloudSerialLogDrainState()

                def fail_after_reset(*_args):
                    reset_cloud_serial_log_state(log_queue, state)
                    log_queue.put_nowait({"id": "new"})
                    state.drain_scheduled = True
                    raise RuntimeError("old scheduler failed")

                async def no_op(_ws, _generation):
                    pass

                if failure_point == "continuation":
                    log_queue.put_nowait({"id": 1})
                    log_queue.put_nowait({"id": 2})
                    async def send(_payload):
                        pass
                    ws = SimpleNamespace(send=send)
                    await drain_cloud_serial_log_queue(
                        ws, log_queue=log_queue, batch_size=1, state=state,
                        is_current_connection=lambda candidate: candidate is ws,
                        is_connected=lambda: True, create_task=fail_after_reset,
                    )
                else:
                    self.assertFalse(schedule_cloud_serial_log_drain(
                        "loop", object(), state=state,
                        drain_coro_factory=fail_after_reset if failure_point == "factory" else no_op,
                        run_coroutine_threadsafe=fail_after_reset,
                    ))
                self.assertTrue(state.drain_scheduled)
                self.assertEqual(list(log_queue.queue), [{"id": "new"}])
                self.assertEqual(log_queue.unfinished_tasks, 1)

    async def test_late_payload_builder_cannot_enqueue_after_reset_or_reauthorization(self):
        for change in ("reset", "connection", "authorization", "identity"):
            with self.subTest(change=change):
                log_queue = queue.Queue(maxsize=2)
                state = CloudSerialLogDrainState()
                loop = asyncio.get_running_loop()
                current = [object()]
                authorized, imei = [True], ["000000000000001"]

                def build_payload(_line):
                    if change == "reset":
                        reset_cloud_serial_log_state(log_queue, state)
                    elif change == "connection":
                        current[0] = object()
                    elif change == "authorization":
                        authorized[0] = False
                    else:
                        imei[0] = "000000000000002"
                    log_queue.put_nowait({"id": "new"})
                    state.drain_scheduled = True
                    return {"id": "late"}

                result = send_cloud_serial_log_runtime(
                    "synthetic", authorized=True,
                    get_loop=lambda: loop, get_ws=lambda: current[0],
                    is_connected=lambda: True, runtime_imei=lambda: imei[0],
                    build_payload=build_payload, log_queue=log_queue, state=state,
                    is_authorized=lambda: authorized[0],
                    schedule_drain=lambda *_args, **_kwargs: self.fail("late producer scheduled"),
                )
                self.assertIn(result, ("stale", "not_connected", "unauthorized"))
                self.assertEqual(list(log_queue.queue), [{"id": "new"}])
                self.assertEqual(log_queue.unfinished_tasks, 1)
                self.assertTrue(state.drain_scheduled)

    async def test_reset_before_scheduling_rejects_old_generation(self):
        state = CloudSerialLogDrainState()
        log_queue = queue.Queue()
        generation = state.generation
        reset_cloud_serial_log_state(log_queue, state)
        self.assertFalse(schedule_cloud_serial_log_drain(
            "loop", object(), state=state, generation=generation,
            drain_coro_factory=lambda *_args: self.fail("stale generation scheduled"),
        ))
        self.assertFalse(state.drain_scheduled)


if __name__ == "__main__":
    unittest.main()
