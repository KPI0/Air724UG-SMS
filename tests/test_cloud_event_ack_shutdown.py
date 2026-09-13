"""Shutdown must wait for a producer that already claimed the drain slot."""
import asyncio
import queue
import threading
import unittest

from sms_core.cloud_sms_event_runtime import (
    CloudSmsEventDrainState, drain_cloud_sms_event_queue,
    schedule_cloud_sms_event_drain, stop_cloud_sms_event_drain,
    handle_cloud_sms_event_ack,
)


class CloudEventAckShutdownTests(unittest.TestCase):
    def check_shutdown_during_submission(self, pause_at):
        loop = asyncio.new_event_loop()
        state = CloudSmsEventDrainState()
        events = queue.Queue()
        events.put({"imei": "000000000000001", "type": "sms_event", "content": "synthetic"})
        entered, release = threading.Event(), threading.Event()
        stopping, stopped = threading.Event(), threading.Event()
        connected = True
        ws = object()
        sent, producer_results, errors = [], [], []

        def pause_producer():
            entered.set()
            if not release.wait(3):
                raise AssertionError("Producer was not released")

        async def send(*args):
            sent.append(args)
            return "sent"

        def factory(socket, generation, imei):
            if pause_at == "factory":
                pause_producer()
            return drain_cloud_sms_event_queue(
                socket, event_queue=events, state=state, batch_size=100,
                generation=generation, runtime_imei=lambda: "000000000000001", drain_imei=imei,
                is_current_connection=lambda value: value is ws,
                is_connected=lambda: connected, is_authorized=lambda: connected,
                send_payload=send,
            )

        def submit(coro, target_loop):
            if pause_at == "submit":
                pause_producer()
            return asyncio.run_coroutine_threadsafe(coro, target_loop)

        def producer():
            producer_results.append(schedule_cloud_sms_event_drain(
                loop, ws, state=state, drain_coro_factory=factory,
                runtime_imei=lambda: "000000000000001", can_schedule=lambda: connected,
                run_coroutine_threadsafe=submit, log_error=errors.append,
            ))

        def resume_producer():
            # The old implementation finishes stop before the submission is
            # visible. With atomic publication, stop waits on the producer;
            # release it after a short, bounded scheduling pause instead.
            stopping.wait(3)
            stopped.wait(0.1)
            release.set()

        producer_thread = threading.Thread(target=producer)
        release_thread = threading.Thread(target=resume_producer)

        async def run_main():
            nonlocal connected
            producer_thread.start()
            release_thread.start()
            while not entered.is_set():
                await asyncio.sleep(0)
            connected = False
            stopping.set()
            await stop_cloud_sms_event_drain(state)
            stopped.set()
            # Do not yield again after the shutdown boundary: the real cloud
            # thread closes its loop as soon as run_until_complete returns.
            producer_thread.join(3)
            release_thread.join(3)

        try:
            loop.run_until_complete(run_main())
            observed = {
                "scheduled": state.drain_scheduled,
                "futures": len(state.drain_futures),
                "tasks": len(asyncio.all_tasks(loop)),
                "retained": events.unfinished_tasks,
            }
        finally:
            release.set()
            stopping.set()
            stopped.set()
            producer_thread.join(3)
            release_thread.join(3)
            # Clean up any orphan created by a failing implementation without
            # leaving warnings or pending tasks in the test runner itself.
            loop.run_until_complete(stop_cloud_sms_event_drain(state))
            loop.close()
        self.assertFalse(producer_thread.is_alive())
        self.assertFalse(release_thread.is_alive())
        self.assertEqual(errors, [])
        self.assertEqual(producer_results, [True])
        self.assertEqual(sent, [])
        self.assertEqual(observed, {"scheduled": False, "futures": 0, "tasks": 0, "retained": 1})

    def test_shutdown_joins_producer_paused_during_coro_creation(self):
        self.check_shutdown_during_submission("factory")

    def test_shutdown_joins_producer_paused_during_loop_submission(self):
        self.check_shutdown_during_submission("submit")


class CloudEventAckBatchTests(unittest.TestCase):
    def test_multiple_ack_batches_keep_device_order_and_cleanup_all_futures(self):
        loop = asyncio.new_event_loop()
        state = CloudSmsEventDrainState()
        events = queue.Queue(maxsize=1000)
        imei, other_imei = "000000000000001", "000000000000002"
        ws = object()
        connected = True
        original, unrelated = [], []
        for index in range(205):
            if index in (50, 150):
                other = {"imei": other_imei, "sequence": index}
                events.put(other)
                unrelated.append(other)
            payload = {"imei": imei, "type": "sms_event", "ack_required": True,
                       "source_event_id": f"batch-{index}", "content": f"synthetic {index}"}
            original.append(payload)
            events.put(payload)
        sent = []

        async def send(socket, payload):
            sent.append(payload)
            # Lose one ACK at the first batch boundary, then accept its retry.
            if payload is original[99] and sent.count(payload) == 1:
                return "sent"
            self.assertTrue(handle_cloud_sms_event_ack(socket, {
                "type": "device_event_ack", "ok": True, "imei": imei,
                "source_event_id": payload["source_event_id"],
            }, state=state))
            return "sent"

        def factory(socket, generation, drain_imei):
            return drain_cloud_sms_event_queue(
                socket, event_queue=events, state=state, batch_size=100,
                generation=generation, drain_imei=drain_imei, runtime_imei=lambda: imei,
                is_current_connection=lambda value: value is ws,
                is_connected=lambda: connected, is_authorized=lambda: connected, send_payload=send,
                ack_timeout=0.005, retry_base=0.005, retry_max=0.01,
            )

        async def finished():
            while state.drain_scheduled:
                await asyncio.sleep(0.001)

        async def run_main():
            nonlocal connected
            try:
                self.assertTrue(schedule_cloud_sms_event_drain(
                    loop, ws, state=state, drain_coro_factory=factory,
                    runtime_imei=lambda: imei, can_schedule=lambda: connected,
                ))
                await asyncio.wait_for(finished(), 5)
                self.assertEqual(sent, original[:100] + original[99:])
                self.assertEqual(list(events.queue), unrelated)
                self.assertEqual(events.unfinished_tasks, 2)
                self.assertIsNone(state.pending_ack)
            finally:
                connected = False
                await stop_cloud_sms_event_drain(state)

        try:
            # Inspect at the same boundary as the cloud thread. Completed
            # tasks' done callbacks run before run_until_complete returns,
            # but need not run synchronously inside stop's gather call.
            loop.run_until_complete(run_main())
            self.assertFalse(state.drain_scheduled)
            self.assertFalse(state.drain_futures)
            self.assertFalse(state.active_ack_tasks)
            self.assertFalse(asyncio.all_tasks(loop))
        finally:
            connected = False
            loop.run_until_complete(stop_cloud_sms_event_drain(state))
            loop.close()
