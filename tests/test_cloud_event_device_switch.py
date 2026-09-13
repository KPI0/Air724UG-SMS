"""Device queues retain the identity and order assigned by their producer."""
import asyncio
import json
import queue
import threading
import unittest

from sms_core.cloud_message_namespace_runtime import (
    drain_cloud_sms_event_queue_namespace_runtime,
    handle_cloud_message_namespace_runtime,
    schedule_cloud_sms_event_drain_namespace_runtime,
    send_cloud_call_event_namespace_runtime,
    send_cloud_sms_event_namespace_runtime,
)
from sms_core.cloud_payloads import build_call_event_payload, build_sms_event_payload
from sms_core.cloud_protocol import cloud_login_ack_matches_imei, normalize_imei
from sms_core.cloud_serial_log_runtime import CloudSerialLogDrainState, reset_cloud_serial_log_state
from sms_core.cloud_sms_event_runtime import (
    CloudSmsEventDrainState,
    clear_cloud_sms_event_state,
    drain_cloud_sms_event_queue,
    enqueue_cloud_sms_event_runtime,
    handle_cloud_sms_event_ack,
    schedule_cloud_sms_event_drain,
)
from sms_core.cloud_state_namespace_runtime import (
    cloud_runtime_imei_namespace_runtime,
    set_cloud_device_imei_namespace_runtime,
)
from sms_core.cloud_ws_namespace_runtime import wait_cloud_login_ack_namespace_runtime


A, B, C = "000000000000001", "000000000000002", "000000000000003"


def event(imei, sequence, event_type="sms_event"):
    return {"imei": imei, "sequence": sequence, "type": event_type}


class CloudEventDeviceSwitchTests(unittest.IsolatedAsyncioTestCase):
    def namespace(self):
        namespace = {
            "CLOUD_DEVICE_IMEI": A, "cloud_imei_verified": True,
            "CLOUD_CONTROL_ENABLED": True,
            "CLOUD_SMS_EVENT_Q": queue.Queue(maxsize=8),
            "CLOUD_SMS_EVENT_DRAIN_STATE": CloudSmsEventDrainState(),
            "CLOUD_SMS_EVENT_DRAIN_BATCH": 100,
            "CLOUD_SERIAL_LOG_Q": queue.Queue(maxsize=8),
            "CLOUD_SERIAL_LOG_DRAIN_STATE": CloudSerialLogDrainState(),
            "cloud_ws_loop": asyncio.get_running_loop(),
            "cloud_ws_conn": None, "cloud_connected": False,
            "cloud_device_authorized": False, "cloud_stop_event": threading.Event(),
            "_normalize_imei": normalize_imei,
            "_cloud_now_ts": lambda: 1700000000,
            "_cloud_log": lambda *_args, **_kwargs: None,
            "_cloud_safe_preview": lambda _text: "synthetic",
            "_notify_cloud_identity_changed": lambda: None,
            "_cloud_build_sms_event_payload": build_sms_event_payload,
            "_cloud_build_call_event_payload": build_call_event_payload,
            "set_cloud_auth_status_from_ack": lambda _data: None,
            "set_cloud_status": lambda *_args: None,
            "_cloud_auth_matches": lambda _data: False,
            "_cloud_check_replay_window": lambda *_args, **_kwargs: None,
            "_cloud_send_status_payload": lambda: {},
            "show_window": lambda: None, "hide_window": lambda: None,
        }
        namespace["_cloud_runtime_imei"] = lambda: cloud_runtime_imei_namespace_runtime(namespace)
        namespace["_cloud_identity_payload"] = lambda: {"imei": namespace["_cloud_runtime_imei"]()}
        namespace["_reset_cloud_serial_log_state"] = lambda: reset_cloud_serial_log_state(
            namespace["CLOUD_SERIAL_LOG_Q"], namespace["CLOUD_SERIAL_LOG_DRAIN_STATE"]
        )
        # Most cases drive the drain explicitly; callback arguments are still
        # the real namespace producers' arguments.
        namespace["_schedule_cloud_sms_event_drain"] = lambda *_args, **_kwargs: False

        async def reply(_ws, _payload):
            self.fail("an authorization ACK must not trigger a command reply")

        async def send(_ws, _payload):
            self.fail("test must install its transport before draining")

        namespace["_cloud_reply"] = reply
        namespace["_cloud_send_payload"] = send
        return namespace

    async def authorize(self, namespace, imei):
        return await handle_cloud_message_namespace_runtime(
            namespace, namespace["cloud_ws_conn"],
            json.dumps({"type": "device_login_ack", "ok": True, "auth_status": "authorized", "imei": imei}),
        )

    async def test_namespace_offline_sms_and_call_survive_a_b_a_switch(self):
        namespace = self.namespace()
        state = namespace["CLOUD_SMS_EVENT_DRAIN_STATE"]
        event_queue = namespace["CLOUD_SMS_EVENT_Q"]
        self.assertEqual(send_cloud_sms_event_namespace_runtime(
            namespace, "synthetic callback", "synthetic body", {"message_trace_id": "synthetic-trace"}
        ), "queued_offline")
        self.assertEqual(send_cloud_call_event_namespace_runtime(
            namespace, "synthetic-caller", "synthetic incoming call", call_session_id="synthetic-call"
        ), "queued_offline")
        original_events = list(event_queue.queue)
        namespace["cloud_device_authorized"] = True
        namespace["CLOUD_SERIAL_LOG_Q"].put_nowait({"imei": A, "data": "old log"})
        self.assertTrue(set_cloud_device_imei_namespace_runtime(namespace, B))
        self.assertFalse(namespace["cloud_device_authorized"])
        self.assertEqual(state.generation, 0)
        self.assertEqual(list(event_queue.queue), original_events)
        self.assertTrue(namespace["CLOUD_SERIAL_LOG_Q"].empty())

        ws = object()
        namespace["cloud_ws_conn"] = ws
        namespace["cloud_connected"] = True
        accepted, rejected = [], []

        async def send(current_ws, payload):
            self.assertIs(current_ws, ws)
            (accepted if payload["imei"] == namespace["CLOUD_DEVICE_IMEI"] else rejected).append(payload)
            handle_cloud_sms_event_ack(current_ws, {
                "type": "device_event_ack", "ok": True, "imei": payload["imei"],
                "source_event_id": payload["source_event_id"],
            }, state=state)
            return "sent"

        namespace["_cloud_send_payload"] = send
        self.assertEqual(await self.authorize(namespace, A), "stale_ack")
        self.assertFalse(namespace["cloud_device_authorized"])
        self.assertFalse(schedule_cloud_sms_event_drain_namespace_runtime(namespace))
        self.assertEqual(await drain_cloud_sms_event_queue_namespace_runtime(namespace, ws), "not_connected")
        self.assertEqual(send_cloud_call_event_namespace_runtime(
            namespace, "synthetic-caller", "synthetic B call", call_session_id="synthetic-B"
        ), "queued_offline")
        await self.authorize(namespace, B)
        await drain_cloud_sms_event_queue_namespace_runtime(namespace, ws)
        self.assertEqual([payload["imei"] for payload in accepted], [B])
        self.assertEqual(list(event_queue.queue), original_events)
        self.assertFalse(schedule_cloud_sms_event_drain_namespace_runtime(namespace))

        self.assertTrue(set_cloud_device_imei_namespace_runtime(namespace, A))
        self.assertFalse(namespace["cloud_device_authorized"])
        await self.authorize(namespace, A)
        await drain_cloud_sms_event_queue_namespace_runtime(namespace, ws)
        self.assertEqual(accepted[1:], original_events)
        self.assertEqual(rejected, [])
        self.assertEqual(event_queue.unfinished_tasks, 0)
        self.assertEqual(state.generation, 0)
        clear_cloud_sms_event_state(event_queue, state)
        self.assertEqual(state.generation, 1)

    async def test_batch_continuations_leave_other_devices_in_original_order(self):
        event_queue = queue.Queue(maxsize=8)
        original = [event(B, 1), event(A, 1), event(C, 1), event(A, 2), event(B, 2), event(A, 3)]
        for payload in original:
            event_queue.put_nowait(payload)
        state = CloudSmsEventDrainState(drain_scheduled=True)
        sent, scheduled = [], []

        async def send(_ws, payload):
            sent.append(payload)
            return "sent"

        await drain_cloud_sms_event_queue(
            "ws", event_queue=event_queue, batch_size=1, state=state,
            is_current_connection=lambda _ws: True, is_connected=lambda: True,
            is_authorized=lambda: True, runtime_imei=lambda: A,
            send_payload=send, create_task=scheduled.append,
        )
        while scheduled:
            await scheduled.pop(0)
        self.assertEqual(sent, [original[1], original[3], original[5]])
        self.assertEqual(list(event_queue.queue), [original[0], original[2], original[4]])
        self.assertEqual(event_queue.unfinished_tasks, 3)
        self.assertFalse(state.drain_scheduled)

    async def test_unmatched_queue_does_not_spin_or_change_oldest_eviction(self):
        event_queue = queue.Queue(maxsize=3)
        original = [event(A, 1), event(B, 1), event(A, 2)]
        for payload in original:
            event_queue.put_nowait(payload)
        state = CloudSmsEventDrainState(drain_scheduled=True)
        result = await drain_cloud_sms_event_queue(
            "ws", event_queue=event_queue, batch_size=1, state=state,
            is_current_connection=lambda _ws: True, is_connected=lambda: True,
            is_authorized=lambda: True, runtime_imei=lambda: C,
            send_payload=lambda *_args: self.fail("unmatched event sent"),
            create_task=lambda _coro: self.fail("unmatched queue rescheduled"),
        )
        self.assertEqual(result, "empty")
        self.assertEqual(list(event_queue.queue), original)
        self.assertFalse(state.drain_scheduled)
        newest = event(C, 1)
        enqueue_cloud_sms_event_runtime(
            newest, event_queue=event_queue, state=state, can_send=False, loop=None, ws=None,
            schedule_drain=lambda *_args: None,
        )
        self.assertEqual(list(event_queue.queue), original[1:] + [newest])
        self.assertEqual(event_queue.unfinished_tasks, 3)

    async def test_failed_event_keeps_order_when_producers_refill_bounded_queue(self):
        for refill in (False, True):
            with self.subTest(refill=refill):
                event_queue = queue.Queue(maxsize=4)
                original = [event(B, 1), event(A, 1), event(C, 1), event(A, 2)]
                for payload in original:
                    event_queue.put_nowait(payload)
                state = CloudSmsEventDrainState(drain_scheduled=True)
                newcomers = [event(B, 2), event(C, 2)]

                async def fail(_ws, payload):
                    self.assertIs(payload, original[1])
                    if refill:
                        for newcomer in newcomers:
                            enqueue_cloud_sms_event_runtime(
                                newcomer, event_queue=event_queue, state=state,
                                can_send=False, loop=None, ws=None, schedule_drain=lambda *_args: None,
                            )
                    raise RuntimeError("synthetic transport failure")

                self.assertEqual(await drain_cloud_sms_event_queue(
                    "ws", event_queue=event_queue, batch_size=4, state=state,
                    is_current_connection=lambda _ws: True, is_connected=lambda: True,
                    is_authorized=lambda: True, runtime_imei=lambda: A, send_payload=fail,
                ), "error")
                expected = [original[1], original[3], *newcomers] if refill else original
                self.assertEqual(list(event_queue.queue), expected)
                self.assertEqual(event_queue.unfinished_tasks, 4)
                self.assertFalse(state.drain_scheduled)
                clear_cloud_sms_event_state(event_queue, state)
                self.assertEqual(event_queue.unfinished_tasks, 0)

    async def test_identity_or_socket_change_during_send_stops_old_drain(self):
        for change in ("identity", "socket"):
            for outcome in ("sent", "error"):
                with self.subTest(change=change, outcome=outcome):
                    event_queue = queue.Queue(maxsize=4)
                    first, second, other = event(A, 1), event(A, 2), event(B, 1)
                    for payload in (first, second, other):
                        event_queue.put_nowait(payload)
                    state = CloudSmsEventDrainState(drain_scheduled=True)
                    old_ws = object()
                    current_ws, current_imei = [old_ws], [A]
                    sent, reschedules = [], []

                    async def send(_ws, payload):
                        sent.append(payload)
                        if change == "identity":
                            current_imei[0] = B
                        else:
                            current_ws[0] = object()
                        return outcome

                    await drain_cloud_sms_event_queue(
                        old_ws, event_queue=event_queue, batch_size=4, state=state,
                        is_current_connection=lambda ws: ws is current_ws[0],
                        is_connected=lambda: True, is_authorized=lambda: True,
                        runtime_imei=lambda: current_imei[0], send_payload=send,
                        schedule_pending=lambda: reschedules.append(state.drain_scheduled),
                    )
                    self.assertEqual(sent, [first])
                    self.assertEqual(list(event_queue.queue), [second, other] if outcome == "sent" else [first, second, other])
                    self.assertEqual(event_queue.unfinished_tasks, 2 if outcome == "sent" else 3)
                    self.assertEqual(reschedules, [False])
                    self.assertEqual(state.generation, 0)

    async def test_scheduler_captures_identity_before_coroutine_starts(self):
        state = CloudSmsEventDrainState()
        event_queue = queue.Queue(maxsize=2)
        event_queue.put_nowait(event(A, 1))
        event_queue.put_nowait(event(B, 1))
        current_imei, submitted = [A], []

        def factory(ws, generation, drain_imei):
            return drain_cloud_sms_event_queue(
                ws, event_queue=event_queue, batch_size=2, state=state,
                is_current_connection=lambda _ws: True, is_connected=lambda: True,
                is_authorized=lambda: True, runtime_imei=lambda: current_imei[0],
                generation=generation, drain_imei=drain_imei,
                send_payload=lambda *_args: self.fail("old task sent with new identity"),
            )

        self.assertTrue(schedule_cloud_sms_event_drain(
            "loop", "ws", state=state, drain_coro_factory=factory,
            runtime_imei=lambda: current_imei[0],
            run_coroutine_threadsafe=lambda coro, _loop: submitted.append(coro),
        ))
        current_imei[0] = B
        self.assertEqual(await submitted[0], "not_connected")
        self.assertEqual(event_queue.qsize(), 2)
        self.assertEqual(event_queue.unfinished_tasks, 2)
        self.assertFalse(state.drain_scheduled)

    async def test_cancelled_send_retains_event_and_balances_task_count(self):
        state = CloudSmsEventDrainState(drain_scheduled=True)
        event_queue = queue.Queue(maxsize=2)
        payload = event(A, 1)
        event_queue.put_nowait(payload)

        async def cancel(_ws, _payload):
            raise asyncio.CancelledError

        with self.assertRaises(asyncio.CancelledError):
            await drain_cloud_sms_event_queue(
                "ws", event_queue=event_queue, batch_size=2, state=state,
                is_current_connection=lambda _ws: True, is_connected=lambda: True,
                is_authorized=lambda: True, runtime_imei=lambda: A, send_payload=cancel,
            )
        self.assertEqual(list(event_queue.queue), [payload])
        self.assertEqual(event_queue.unfinished_tasks, 1)
        self.assertFalse(state.drain_scheduled)

    async def test_old_ack_cannot_authorize_or_revoke_new_identity(self):
        namespace = self.namespace()
        namespace["cloud_ws_conn"] = object()
        namespace["cloud_connected"] = True
        self.assertTrue(set_cloud_device_imei_namespace_runtime(namespace, B))
        applied = []

        def apply_status(data):
            self.assertTrue(namespace["CLOUD_SMS_EVENT_DRAIN_STATE"].lock.locked())
            applied.append(data)

        namespace["set_cloud_auth_status_from_ack"] = apply_status
        for authorized in (False, True):
            for status in ("authorized", "failed", "waiting"):
                with self.subTest(authorized=authorized, status=status):
                    namespace["cloud_device_authorized"] = authorized
                    result = await handle_cloud_message_namespace_runtime(
                        namespace, namespace["cloud_ws_conn"],
                        json.dumps({"type": "device_login_ack", "imei": A, "auth_status": status}),
                    )
                    self.assertEqual(result, "stale_ack")
                    self.assertEqual(namespace["cloud_device_authorized"], authorized)
        self.assertEqual(applied, [])
        await self.authorize(namespace, B)
        self.assertTrue(namespace["cloud_device_authorized"])
        self.assertEqual(len(applied), 1)
        self.assertTrue(set_cloud_device_imei_namespace_runtime(namespace, B))
        self.assertTrue(namespace["cloud_device_authorized"], "unchanged identity must keep authorization")

    async def test_initial_handshake_waits_for_current_imei_ack(self):
        for old_status in ("authorized", "failed", "waiting"):
            with self.subTest(old_status=old_status):
                namespace = self.namespace()
                set_cloud_device_imei_namespace_runtime(namespace, B)
                messages = [
                    {"type": "device_login_ack", "auth_status": old_status, "imei": A},
                    {"type": "device_login_ack", "auth_status": "authorized", "imei": B},
                ]
                applied = []

                class Socket:
                    async def recv(self):
                        return json.dumps(messages.pop(0))

                ws = Socket()
                namespace["cloud_ws_conn"] = ws
                namespace["set_cloud_auth_status_from_ack"] = applied.append
                self.assertTrue(await wait_cloud_login_ack_namespace_runtime(namespace, ws, timeout=1))
                self.assertTrue(namespace["cloud_device_authorized"])
                self.assertEqual([ack["imei"] for ack in applied], [B])
                self.assertEqual(messages, [])

    async def test_ack_identity_aliases_and_legacy_missing_imei(self):
        for data, expected in (
            ({}, True), ({"imei": "", "device_imei": B}, True),
            ({"imei": B, "device_imei": B}, True),
            ({"imei": B, "device_imei": A}, False),
            ({"target_imei": A}, False), ({"imei": "invalid"}, False),
        ):
            with self.subTest(data=data):
                self.assertEqual(cloud_login_ack_matches_imei(data, B), expected)
        self.assertFalse(cloud_login_ack_matches_imei({}, ""))
        namespace = self.namespace()
        namespace["cloud_ws_conn"] = object()
        await handle_cloud_message_namespace_runtime(
            namespace, namespace["cloud_ws_conn"],
            json.dumps({"type": "device_login_ack", "auth_status": "authorized"}),
        )
        self.assertTrue(namespace["cloud_device_authorized"])


if __name__ == "__main__":
    unittest.main()
