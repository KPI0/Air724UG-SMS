"""Reliable desktop events must survive transport success without a save ACK."""
import asyncio
import json
import queue
import unittest
from types import SimpleNamespace
from unittest import mock

from sms_core import cloud_event_ack_runtime as ack_runtime
from sms_core import cloud_sms_event_runtime as events
from sms_core.cloud_payloads import (
    build_call_event_payload, build_sent_sms_event_payload, build_sms_event_payload,
)
from tests import test_cloud_event_device_switch as device_switch

A, B = device_switch.A, device_switch.B


def payloads():
    identity = {"imei": A, "device_imei": A}
    return [
        build_sms_event_payload("", "synthetic SMS", 1700000000, identity),
        build_sent_sms_event_payload("10086", "synthetic SMS", 1700000000, identity, "desktop-send-test"),
        build_call_event_payload("10086", "synthetic call", 1700000000, identity, call_session_id="call-test"),
    ]


def ack(payload, **overrides):
    return {
        "type": "device_event_ack", "ok": True,
        "imei": payload["imei"], "device_imei": payload["imei"],
        "event_id": payload["source_event_id"],
        "source_event_id": payload["source_event_id"], **overrides,
    }


class CloudEventAckTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.queue = queue.Queue(maxsize=4)
        self.state = events.CloudSmsEventDrainState()
        self.ws = object()
        self.current_ws = self.ws
        self.imei = A
        self.authorized = True
        self.writes = []

    async def write(self, ws, payload):
        self.writes.append((ws, dict(payload)))
        return "sent"

    def drain(self, **kwargs):
        return events.drain_cloud_sms_event_queue(
            self.ws, event_queue=self.queue, state=self.state, batch_size=100,
            is_current_connection=lambda ws: ws is self.current_ws,
            is_connected=lambda: self.current_ws is not None,
            is_authorized=lambda: self.authorized, runtime_imei=lambda: self.imei,
            send_payload=kwargs.pop("send_payload", self.write), **kwargs,
        )

    async def wait_until(self, predicate):
        async def wait():
            while not predicate():
                await asyncio.sleep(0.001)
        await asyncio.wait_for(wait(), 1)

    def confirm(self, payload, ws=None, **overrides):
        return events.handle_cloud_sms_event_ack(
            self.ws if ws is None else ws, ack(payload, **overrides), state=self.state,
        )

    async def cancel(self, task):
        task.cancel()
        await asyncio.gather(task, return_exceptions=True)

    def test_each_event_builder_requests_ack_with_a_stable_bounded_id(self):
        first, second = payloads(), payloads()
        for index, payload in enumerate(first):
            with self.subTest(kind=index):
                self.assertIs(payload.get("ack_required"), True)
                self.assertTrue(payload.get("source_event_id"))
                self.assertLessEqual(len(payload["source_event_id"].encode()), 64)
        self.assertNotEqual(first[0]["source_event_id"], second[0]["source_event_id"])
        self.assertEqual(first[1]["source_event_id"], second[1]["source_event_id"])
        self.assertNotEqual(first[2]["source_event_id"], second[2]["source_event_id"])

    async def test_transport_success_without_ack_is_not_completion(self):
        for payload in payloads():
            with self.subTest(kind=payload["type"], direction=payload.get("direction")):
                self.queue.put(payload)
                self.writes.clear()
                task = asyncio.create_task(self.drain())
                try:
                    await self.wait_until(lambda: self.writes)
                    self.assertFalse(task.done(), "A WebSocket write is not a persistence ACK")
                    self.assertEqual(self.queue.unfinished_tasks, 1)
                finally:
                    await self.cancel(task)
                self.assertIs(self.queue.get_nowait(), payload)
                self.queue.task_done()

    async def test_only_matching_current_positive_ack_retires_event(self):
        payload = payloads()[0]
        self.queue.put(payload)
        task = asyncio.create_task(self.drain())
        try:
            await self.wait_until(lambda: self.writes)
            for changes in ({"ok": False}, {"ok": "true"}, {"imei": B},
                            {"device_imei": B}, {"event_id": "other"},
                            {"source_event_id": "other"}, {"imei": "", "device_imei": ""}):
                self.assertFalse(self.confirm(payload, **changes))
            self.assertFalse(self.confirm(payload, ws=object()))
            self.authorized = False
            self.assertFalse(self.confirm(payload))
            self.authorized = True
            self.assertTrue(self.confirm(payload))
            self.assertFalse(self.confirm(payload), "A duplicate ACK must be inert")
            await asyncio.wait_for(task, 1)
            self.assertTrue(self.queue.empty())
            self.assertEqual(self.queue.unfinished_tasks, 0)
            self.assertFalse(self.confirm(payload))
        finally:
            await self.cancel(task)

    async def test_ack_during_send_is_not_missed(self):
        payload = payloads()[0]
        self.queue.put(payload)
        async def send(ws, outgoing):
            await self.write(ws, outgoing)
            self.assertTrue(self.confirm(outgoing))
            return "sent"
        await asyncio.wait_for(self.drain(send_payload=send), 1)
        self.assertEqual(self.queue.unfinished_tasks, 0)

    async def test_lost_ack_retries_same_event_with_backoff(self):
        payload = payloads()[1]
        self.queue.put(payload)
        warnings = []
        waits = asyncio.Queue()

        async def controlled_wait_for(awaitable, timeout):
            operation = asyncio.ensure_future(awaitable)
            expiry = asyncio.get_running_loop().create_future()
            waits.put_nowait((timeout, expiry))
            try:
                done, _ = await asyncio.wait(
                    (operation, expiry), return_when=asyncio.FIRST_COMPLETED,
                )
                if operation in done:
                    return operation.result()
                raise asyncio.TimeoutError
            finally:
                expiry.cancel()
                if not operation.done():
                    operation.cancel()
                await asyncio.gather(operation, return_exceptions=True)

        async def next_wait(expected_timeout):
            actual_timeout, expiry = await asyncio.wait_for(waits.get(), 1)
            self.assertEqual(actual_timeout, expected_timeout)
            return expiry

        # Windows Python 3.11 has an approximately 15.6 ms monotonic clock.
        # Control timeout expiry instead of measuring tiny real intervals;
        # keep the actual send, ACK wake-up, cancellation and queue handling.
        controlled_asyncio = SimpleNamespace(**vars(asyncio))
        controlled_asyncio.wait_for = controlled_wait_for
        with mock.patch.object(ack_runtime, "asyncio", controlled_asyncio):
            task = asyncio.create_task(self.drain(log_error=warnings.append))
            try:
                backoffs = (2.0, 4.0, 8.0, 16.0, 30.0, 30.0)
                for attempt, delay in enumerate(backoffs, start=1):
                    await next_wait(10.0)  # The transport send completes normally.
                    ack_expiry = await next_wait(10.0)
                    self.assertEqual(len(self.writes), attempt)
                    self.assertTrue(all(item == payload for _ws, item in self.writes))
                    self.assertEqual(self.queue.unfinished_tasks, 1)
                    self.assertFalse(task.done())
                    ack_expiry.set_result(None)

                    retry_expiry = await next_wait(delay)
                    self.assertEqual(len(self.writes), attempt,
                                     "Do not resend before the backoff expires")
                    self.assertEqual(len(warnings), 1)
                    if attempt < len(backoffs):
                        retry_expiry.set_result(None)
                    else:
                        # A late ACK interrupts even the capped retry wait.
                        self.assertTrue(self.confirm(payload))

                await asyncio.wait_for(task, 1)
                self.assertEqual(len(self.writes), len(backoffs))
                self.assertEqual(self.queue.unfinished_tasks, 0)
                self.assertTrue(self.queue.empty())
                self.assertIsNone(self.state.pending_ack)
                self.assertFalse(self.state.active_ack_tasks)
            finally:
                await self.cancel(task)

    async def test_transport_error_retries_event_without_rebuilding_payload(self):
        payload = payloads()[0]
        self.queue.put(payload)
        async def send(ws, outgoing):
            await self.write(ws, outgoing)
            if len(self.writes) == 1:
                raise OSError("synthetic transport failure")
            self.confirm(outgoing)
            return "sent"
        await asyncio.wait_for(self.drain(send_payload=send, retry_base=0.01), 1)
        self.assertEqual(len(self.writes), 2)
        self.assertEqual(self.writes[0][1], self.writes[1][1])

    async def test_disconnect_retains_event_and_old_socket_ack_cannot_finish_retry(self):
        payload = payloads()[2]
        self.queue.put(payload)
        task = asyncio.create_task(self.drain())
        try:
            await self.wait_until(lambda: self.writes)
            old_ws = self.ws
            self.current_ws = None
            events.interrupt_cloud_sms_event_drain(self.state)
            await asyncio.gather(task, return_exceptions=True)
            self.assertIs(self.queue.queue[0], payload)
            self.assertEqual(self.queue.unfinished_tasks, 1)
            self.ws = self.current_ws = object()
            task = asyncio.create_task(self.drain())
            await self.wait_until(lambda: len(self.writes) == 2)
            self.assertFalse(self.confirm(payload, ws=old_ws))
            self.assertTrue(self.confirm(payload))
            await asyncio.wait_for(task, 1)
        finally:
            await self.cancel(task)

    async def test_clear_during_wait_drops_only_old_generation(self):
        payload = payloads()[0]
        self.queue.put(payload)
        task = asyncio.create_task(self.drain())
        try:
            await self.wait_until(lambda: self.writes)
            events.clear_cloud_sms_event_state(self.queue, self.state)
            new = payloads()[1]
            self.queue.put(new)
            self.assertFalse(self.confirm(payload))
            await asyncio.gather(task, return_exceptions=True)
            self.assertEqual(list(self.queue.queue), [new])
            self.assertEqual(self.queue.unfinished_tasks, 1)
        finally:
            await self.cancel(task)

    async def test_shutdown_cancels_hung_send_and_retains_unconfirmed_event(self):
        payload = payloads()[0]
        self.queue.put(payload)
        async def hung_send(ws, outgoing):
            await self.write(ws, outgoing)
            await asyncio.Future()
        task = asyncio.create_task(self.drain(send_payload=hung_send))
        try:
            await self.wait_until(lambda: self.writes)
            self.current_ws = None
            await asyncio.wait_for(events.stop_cloud_sms_event_drain(self.state), 1)
            self.assertTrue(task.done())
            self.assertEqual(list(self.queue.queue), [payload])
            self.assertEqual(self.queue.unfinished_tasks, 1)
            self.assertIsNone(self.state.pending_ack)
            self.assertFalse(self.state.drain_scheduled)
        finally:
            await self.cancel(task)

    async def test_unconfirmed_event_keeps_position_when_bounded_queue_refills(self):
        payload = payloads()[0]
        other = {"type": "sms_event", "imei": B, "content": "older other device"}
        later = payloads()[1]
        for item in (other, payload, later):
            self.queue.put(item)
        task = asyncio.create_task(self.drain())
        try:
            await self.wait_until(lambda: self.writes)
            additions = [{"imei": B, "content": str(index)} for index in range(3)]
            for item in additions:
                events.enqueue_cloud_sms_event_runtime(
                    item, event_queue=self.queue, state=self.state, can_send=False,
                    loop=None, ws=None, schedule_drain=lambda *_args: None,
                )
            await self.cancel(task)
            self.assertEqual(list(self.queue.queue), [payload, *additions])
            self.assertEqual(self.queue.unfinished_tasks, self.queue.maxsize)
        finally:
            await self.cancel(task)

    async def test_shutdown_joins_scheduled_drain_before_it_starts(self):
        payload = payloads()[0]
        self.queue.put(payload)
        self.assertTrue(events.schedule_cloud_sms_event_drain(
            asyncio.get_running_loop(), self.ws, state=self.state,
            drain_coro_factory=lambda _ws, generation, _imei: self.drain(generation=generation),
        ))
        self.current_ws = None
        await asyncio.wait_for(events.stop_cloud_sms_event_drain(self.state), 1)
        self.assertEqual(self.writes, [])
        self.assertEqual(list(self.queue.queue), [payload])
        self.assertFalse(self.state.drain_scheduled)
        self.assertFalse(self.state.drain_futures)


class CloudEventAckNamespaceTests(unittest.IsolatedAsyncioTestCase):
    async def test_identity_switch_wakes_waiter_and_new_login_can_schedule_current_device(self):
        from sms_core.cloud_message_namespace_runtime import (
            drain_cloud_sms_event_queue_namespace_runtime, handle_cloud_message_namespace_runtime,
            schedule_cloud_sms_event_drain_namespace_runtime,
        )
        from sms_core.cloud_state_namespace_runtime import set_cloud_device_imei_namespace_runtime
        namespace = device_switch.CloudEventDeviceSwitchTests().namespace()
        namespace.update(cloud_connected=True, cloud_device_authorized=True, cloud_ws_conn=object())
        state = namespace["CLOUD_SMS_EVENT_DRAIN_STATE"]
        event_queue = namespace["CLOUD_SMS_EVENT_Q"]
        original, other = payloads()[:2]
        other.update(imei=B, device_imei=B)
        event_queue.put(original)
        event_queue.put(other)
        writes = []
        async def send(ws, payload):
            writes.append(payload)
            if payload is other:
                await handle_cloud_message_namespace_runtime(namespace, ws, json.dumps(ack(payload)))
            return "sent"
        namespace["_cloud_send_payload"] = send
        namespace["_cloud_drain_sms_event_queue"] = lambda ws, **kwargs: drain_cloud_sms_event_queue_namespace_runtime(namespace, ws, **kwargs)
        namespace["_schedule_cloud_sms_event_drain"] = lambda: schedule_cloud_sms_event_drain_namespace_runtime(namespace)
        namespace["_schedule_cloud_sms_event_drain"]()
        async def wait_for(predicate):
            while not predicate():
                await asyncio.sleep(0.001)
        try:
            await asyncio.wait_for(wait_for(lambda: writes), 1)
            set_cloud_device_imei_namespace_runtime(namespace, B)
            await handle_cloud_message_namespace_runtime(namespace, namespace["cloud_ws_conn"], json.dumps({
                "type": "device_login_ack", "ok": True, "auth_status": "authorized", "imei": B,
            }))
            # This can race with cancellation of the previous delivery.
            namespace["_schedule_cloud_sms_event_drain"]()
            await asyncio.wait_for(wait_for(lambda: len(writes) == 2 and not state.drain_scheduled), 1)
            self.assertEqual(writes, [original, other])
            self.assertEqual(list(event_queue.queue), [original])
            self.assertEqual(event_queue.unfinished_tasks, 1)
            self.assertEqual(state.generation, 0)
        finally:
            namespace["cloud_connected"] = False
            await events.stop_cloud_sms_event_drain(state)

    async def test_ack_routes_without_command_password_replay_or_serial_execution(self):
        fixture = device_switch.CloudEventDeviceSwitchTests()
        namespace = fixture.namespace()
        namespace.update(cloud_connected=True, cloud_device_authorized=True, cloud_ws_conn=object())
        from sms_core.cloud_message_namespace_runtime import (
            drain_cloud_sms_event_queue_namespace_runtime, handle_cloud_message_namespace_runtime,
        )
        payload = payloads()[0]
        namespace["CLOUD_SMS_EVENT_Q"].put(payload)
        sent = asyncio.Event()
        async def send(_ws, _payload):
            sent.set()
            return "sent"
        namespace["_cloud_send_payload"] = send
        task = asyncio.create_task(drain_cloud_sms_event_queue_namespace_runtime(namespace, namespace["cloud_ws_conn"]))
        try:
            await asyncio.wait_for(sent.wait(), 1)
            result = await handle_cloud_message_namespace_runtime(
                namespace, namespace["cloud_ws_conn"], json.dumps(ack(payload)),
            )
            self.assertEqual(result, "device_event_ack")
            await asyncio.wait_for(task, 1)
            self.assertEqual(namespace["CLOUD_SMS_EVENT_Q"].unfinished_tasks, 0)
        finally:
            task.cancel()
            await asyncio.gather(task, return_exceptions=True)
