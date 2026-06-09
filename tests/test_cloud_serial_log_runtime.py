import asyncio
import queue
import unittest

from sms_core.cloud_serial_log_runtime import (
    CloudSerialLogDrainState,
    clear_cloud_serial_log_queue,
    drain_cloud_serial_log_queue,
    put_drop_oldest,
    reset_cloud_serial_log_state,
    schedule_cloud_serial_log_drain,
    send_cloud_serial_log_runtime,
)


class FakeLoop:
    def __init__(self, running=True):
        self.running = running

    def is_running(self):
        return self.running


class FakeWs:
    def __init__(self):
        self.sent = []

    async def send(self, payload):
        self.sent.append(payload)


class CloudSerialLogRuntimeTests(unittest.TestCase):
    def test_put_drop_oldest_adds_payload_when_queue_has_space(self):
        log_queue = queue.Queue(maxsize=2)

        self.assertTrue(put_drop_oldest(log_queue, {"id": 1}))
        self.assertEqual(log_queue.get_nowait(), {"id": 1})

    def test_put_drop_oldest_drops_oldest_when_full(self):
        log_queue = queue.Queue(maxsize=1)
        log_queue.put_nowait({"id": "old"})

        self.assertTrue(put_drop_oldest(log_queue, {"id": "new"}))

        self.assertEqual(log_queue.get_nowait(), {"id": "new"})

    def test_send_cloud_serial_log_runtime_skips_unavailable_states(self):
        base = {
            "authorized": True,
            "get_loop": lambda: FakeLoop(True),
            "get_ws": lambda: object(),
            "is_connected": lambda: True,
            "runtime_imei": lambda: "imei",
            "build_payload": lambda line: {"line": line},
            "log_queue": queue.Queue(),
            "schedule_drain": lambda loop, ws: None,
        }

        self.assertEqual(send_cloud_serial_log_runtime("AT", **{**base, "authorized": False}), "unauthorized")
        self.assertEqual(send_cloud_serial_log_runtime("AT", **{**base, "get_loop": lambda: None}), "not_connected")
        self.assertEqual(send_cloud_serial_log_runtime("AT", **{**base, "get_loop": lambda: FakeLoop(False)}), "not_connected")
        self.assertEqual(send_cloud_serial_log_runtime("AT", **{**base, "get_ws": lambda: None}), "not_connected")
        self.assertEqual(send_cloud_serial_log_runtime("AT", **{**base, "is_connected": lambda: False}), "not_connected")
        self.assertEqual(send_cloud_serial_log_runtime("AT", **{**base, "runtime_imei": lambda: ""}), "missing_imei")
        self.assertEqual(send_cloud_serial_log_runtime("AT", **{**base, "build_payload": lambda line: None}), "empty")

    def test_send_cloud_serial_log_runtime_queues_payload_and_schedules_drain(self):
        log_queue = queue.Queue()
        loop = FakeLoop()
        ws = object()
        scheduled = []

        result = send_cloud_serial_log_runtime(
            "AT",
            authorized=True,
            get_loop=lambda: loop,
            get_ws=lambda: ws,
            is_connected=lambda: True,
            runtime_imei=lambda: "imei",
            build_payload=lambda line: {"line": line},
            log_queue=log_queue,
            schedule_drain=lambda next_loop, next_ws: scheduled.append((next_loop, next_ws)),
        )

        self.assertEqual(result, "queued")
        self.assertEqual(log_queue.get_nowait(), {"line": "AT"})
        self.assertEqual(scheduled, [(loop, ws)])

    def test_reset_cloud_serial_log_state_clears_queue_and_flag(self):
        log_queue = queue.Queue()
        log_queue.put_nowait({"id": 1})
        state = CloudSerialLogDrainState()
        state.drain_scheduled = True

        reset_cloud_serial_log_state(log_queue, state)

        self.assertTrue(log_queue.empty())
        self.assertFalse(state.drain_scheduled)

    def test_clear_cloud_serial_log_queue_empties_pending_items(self):
        log_queue = queue.Queue()
        log_queue.put_nowait({"id": 1})
        log_queue.put_nowait({"id": 2})

        clear_cloud_serial_log_queue(log_queue)

        self.assertTrue(log_queue.empty())

    def test_drain_cloud_serial_log_queue_sends_batch_and_reschedules(self):
        log_queue = queue.Queue()
        for item_id in (1, 2, 3):
            log_queue.put_nowait({"id": item_id})
        state = CloudSerialLogDrainState()
        state.drain_scheduled = True
        ws = FakeWs()
        scheduled = []

        def create_task(coro):
            scheduled.append(coro)
            coro.close()

        asyncio.run(drain_cloud_serial_log_queue(
            ws,
            log_queue=log_queue,
            batch_size=2,
            state=state,
            is_current_connection=lambda current_ws: current_ws is ws,
            is_connected=lambda: True,
            create_task=create_task,
        ))

        self.assertEqual(ws.sent, ['{"id": 1}', '{"id": 2}'])
        self.assertEqual(log_queue.get_nowait(), {"id": 3})
        self.assertTrue(state.drain_scheduled)
        self.assertEqual(len(scheduled), 1)

    def test_drain_cloud_serial_log_queue_clears_when_connection_stale(self):
        log_queue = queue.Queue()
        log_queue.put_nowait({"id": 1})
        state = CloudSerialLogDrainState()
        state.drain_scheduled = True
        ws = FakeWs()

        asyncio.run(drain_cloud_serial_log_queue(
            ws,
            log_queue=log_queue,
            batch_size=2,
            state=state,
            is_current_connection=lambda _current_ws: False,
            is_connected=lambda: True,
        ))

        self.assertTrue(log_queue.empty())
        self.assertFalse(state.drain_scheduled)
        self.assertEqual(ws.sent, [])

    def test_schedule_cloud_serial_log_drain_sets_and_reuses_flag(self):
        state = CloudSerialLogDrainState()
        loop = FakeLoop()
        ws = object()
        calls = []

        async def drain(_ws):
            pass

        def run_coroutine_threadsafe(coro, next_loop):
            calls.append(next_loop)
            coro.close()

        self.assertTrue(schedule_cloud_serial_log_drain(
            loop,
            ws,
            state=state,
            drain_coro_factory=drain,
            run_coroutine_threadsafe=run_coroutine_threadsafe,
        ))
        self.assertFalse(schedule_cloud_serial_log_drain(
            loop,
            ws,
            state=state,
            drain_coro_factory=drain,
            run_coroutine_threadsafe=run_coroutine_threadsafe,
        ))
        self.assertTrue(state.drain_scheduled)
        self.assertEqual(calls, [loop])

    def test_schedule_cloud_serial_log_drain_resets_flag_on_failure(self):
        state = CloudSerialLogDrainState()

        async def drain(_ws):
            pass

        def run_coroutine_threadsafe(coro, _loop):
            coro.close()
            raise RuntimeError("boom")

        self.assertFalse(schedule_cloud_serial_log_drain(
            FakeLoop(),
            object(),
            state=state,
            drain_coro_factory=drain,
            run_coroutine_threadsafe=run_coroutine_threadsafe,
        ))
        self.assertFalse(state.drain_scheduled)


if __name__ == "__main__":
    unittest.main()
