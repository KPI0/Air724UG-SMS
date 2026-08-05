import asyncio
import queue
import threading
import unittest
from unittest.mock import patch

from sms_core.cloud_sms_event_runtime import (
    CloudSmsEventDrainState,
    clear_cloud_sms_event_state,
    drain_cloud_sms_event_queue,
    enqueue_cloud_sms_event_runtime,
    schedule_cloud_sms_event_drain,
)


class CloudSmsEventRuntimeTests(unittest.TestCase):
    def test_offline_event_is_retained_and_sent_after_authorized_reconnect(self):
        event_queue = queue.Queue(maxsize=4)
        state = CloudSmsEventDrainState()
        payload = {"type": "sms_event", "content": "body"}

        result = enqueue_cloud_sms_event_runtime(
            payload,
            event_queue=event_queue,
            can_send=False,
            loop=None,
            ws=None,
            schedule_drain=lambda *_args: self.fail("offline enqueue must not schedule"),
        )

        sent = []

        async def send_payload(ws, next_payload):
            sent.append((ws, next_payload))
            return "sent"

        drain_result = asyncio.run(drain_cloud_sms_event_queue(
            "ws",
            event_queue=event_queue,
            batch_size=4,
            state=state,
            is_current_connection=lambda ws: ws == "ws",
            is_connected=lambda: True,
            is_authorized=lambda: True,
            send_payload=send_payload,
        ))

        self.assertEqual(result, "queued_offline")
        self.assertEqual(drain_result, "empty")
        self.assertEqual(sent, [("ws", payload)])
        self.assertTrue(event_queue.empty())

    def test_send_failure_requeues_payload_even_if_queue_refills(self):
        event_queue = queue.Queue(maxsize=2)
        state = CloudSmsEventDrainState()
        failed_payload = {"id": "failed"}
        event_queue.put_nowait(failed_payload)

        async def send_payload(_ws, _payload):
            event_queue.put_nowait({"id": "new-1"})
            event_queue.put_nowait({"id": "new-2"})
            return "error"

        result = asyncio.run(drain_cloud_sms_event_queue(
            "ws",
            event_queue=event_queue,
            batch_size=1,
            state=state,
            is_current_connection=lambda _ws: True,
            is_connected=lambda: True,
            is_authorized=lambda: True,
            send_payload=send_payload,
        ))

        remaining = [event_queue.get_nowait(), event_queue.get_nowait()]
        self.assertEqual(result, "error")
        self.assertIn(failed_payload, remaining)
        self.assertFalse(state.drain_scheduled)

    def test_clear_removes_pending_events_and_resets_schedule_state(self):
        event_queue = queue.Queue(maxsize=2)
        event_queue.put_nowait({"id": 1})
        state = CloudSmsEventDrainState(
            lock=threading.Lock(),
            drain_scheduled=True,
        )

        self.assertTrue(clear_cloud_sms_event_state(event_queue, state))
        self.assertTrue(event_queue.empty())
        self.assertFalse(state.drain_scheduled)

    def test_clear_during_failed_send_prevents_old_payload_from_returning(self):
        event_queue = queue.Queue(maxsize=2)
        state = CloudSmsEventDrainState()
        old_payload = {"id": "old"}
        new_payload = {"id": "new"}
        event_queue.put_nowait(old_payload)

        async def send_payload(_ws, _payload):
            clear_cloud_sms_event_state(event_queue, state)
            event_queue.put_nowait(new_payload)
            return "error"

        result = asyncio.run(drain_cloud_sms_event_queue(
            "ws",
            event_queue=event_queue,
            batch_size=1,
            state=state,
            is_current_connection=lambda _ws: True,
            is_connected=lambda: True,
            is_authorized=lambda: True,
            send_payload=send_payload,
        ))

        self.assertEqual(result, "error")
        self.assertEqual(event_queue.get_nowait(), new_payload)
        self.assertTrue(event_queue.empty())

    def test_clear_waits_for_failed_payload_requeue_and_removes_it(self):
        event_queue = queue.Queue(maxsize=2)
        state = CloudSmsEventDrainState()
        event_queue.put_nowait({"id": "old"})
        requeue_started = threading.Event()
        allow_requeue = threading.Event()
        clear_done = threading.Event()
        drain_results = []
        original_put = __import__(
            "sms_core.cloud_sms_event_runtime",
            fromlist=["put_sms_event_drop_oldest"],
        ).put_sms_event_drop_oldest

        def blocking_put(target_queue, payload):
            requeue_started.set()
            self.assertTrue(allow_requeue.wait(timeout=2.0))
            return original_put(target_queue, payload)

        async def send_payload(_ws, _payload):
            return "error"

        def run_drain():
            drain_results.append(asyncio.run(drain_cloud_sms_event_queue(
                "ws",
                event_queue=event_queue,
                batch_size=1,
                state=state,
                is_current_connection=lambda _ws: True,
                is_connected=lambda: True,
                is_authorized=lambda: True,
                send_payload=send_payload,
            )))

        def run_clear():
            clear_cloud_sms_event_state(event_queue, state)
            clear_done.set()

        with patch(
            "sms_core.cloud_sms_event_runtime.put_sms_event_drop_oldest",
            side_effect=blocking_put,
        ):
            drain_thread = threading.Thread(target=run_drain)
            drain_thread.start()
            self.assertTrue(requeue_started.wait(timeout=2.0))
            clear_thread = threading.Thread(target=run_clear)
            clear_thread.start()
            self.assertFalse(clear_done.wait(timeout=0.05))
            allow_requeue.set()
            drain_thread.join(timeout=2.0)
            clear_thread.join(timeout=2.0)

        self.assertEqual(drain_results, ["error"])
        self.assertTrue(clear_done.is_set())
        self.assertTrue(event_queue.empty())

    def test_old_drain_does_not_clear_new_generation_schedule_flag(self):
        event_queue = queue.Queue(maxsize=2)
        state = CloudSmsEventDrainState(drain_scheduled=True)
        event_queue.put_nowait({"id": "old"})

        async def send_payload(_ws, _payload):
            clear_cloud_sms_event_state(event_queue, state)
            with state.lock:
                state.drain_scheduled = True
            return "error"

        asyncio.run(drain_cloud_sms_event_queue(
            "ws",
            event_queue=event_queue,
            batch_size=1,
            state=state,
            is_current_connection=lambda _ws: True,
            is_connected=lambda: True,
            is_authorized=lambda: True,
            send_payload=send_payload,
        ))

        self.assertTrue(state.drain_scheduled)

    def test_scheduled_drain_captures_generation_before_coroutine_starts(self):
        event_queue = queue.Queue(maxsize=2)
        event_queue.put_nowait({"id": "old"})
        state = CloudSmsEventDrainState()
        submitted = []
        sent = []
        ws = object()

        async def send_payload(_ws, payload):
            sent.append(payload)
            return "sent"

        def drain_factory(current_ws, generation):
            return drain_cloud_sms_event_queue(
                current_ws,
                event_queue=event_queue,
                batch_size=1,
                state=state,
                is_current_connection=lambda current_ws: current_ws is ws,
                is_connected=lambda: True,
                is_authorized=lambda: True,
                send_payload=send_payload,
                generation=generation,
            )

        self.assertTrue(schedule_cloud_sms_event_drain(
            "loop",
            ws,
            state=state,
            drain_coro_factory=drain_factory,
            run_coroutine_threadsafe=lambda coro, _loop: submitted.append(coro),
        ))
        clear_cloud_sms_event_state(event_queue, state)
        event_queue.put_nowait({"id": "new"})
        self.assertTrue(schedule_cloud_sms_event_drain(
            "loop",
            ws,
            state=state,
            drain_coro_factory=drain_factory,
            run_coroutine_threadsafe=lambda coro, _loop: submitted.append(coro),
        ))

        self.assertEqual(asyncio.run(submitted[0]), "stale")
        self.assertEqual(sent, [])
        self.assertEqual(list(event_queue.queue), [{"id": "new"}])
        self.assertTrue(state.drain_scheduled)
        self.assertEqual(asyncio.run(submitted[1]), "sent")
        self.assertEqual(sent, [{"id": "new"}])
        self.assertTrue(event_queue.empty())
        self.assertEqual(event_queue.unfinished_tasks, 0)
        self.assertFalse(state.drain_scheduled)

    def test_enqueue_checks_enabled_state_under_same_lock_as_clear(self):
        event_queue = queue.Queue(maxsize=2)
        state = CloudSmsEventDrainState()
        enabled = [True]

        result = enqueue_cloud_sms_event_runtime(
            {"id": "late"},
            event_queue=event_queue,
            can_send=False,
            loop=None,
            ws=None,
            schedule_drain=lambda *_args: None,
            state=state,
            is_enabled=lambda: enabled[0],
        )
        self.assertEqual(result, "queued_offline")

        enabled[0] = False
        clear_cloud_sms_event_state(event_queue, state)
        result = enqueue_cloud_sms_event_runtime(
            {"id": "disabled"},
            event_queue=event_queue,
            can_send=False,
            loop=None,
            ws=None,
            schedule_drain=lambda *_args: None,
            state=state,
            is_enabled=lambda: enabled[0],
        )
        self.assertEqual(result, "disabled")
        self.assertTrue(event_queue.empty())


if __name__ == "__main__":
    unittest.main()
