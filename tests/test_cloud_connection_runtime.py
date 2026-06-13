import asyncio
import unittest

from sms_core.cloud_connection_runtime import (
    reply_cloud_payload_runtime,
    schedule_cloud_unregister_runtime,
    send_cloud_payload_runtime,
    serialize_cloud_payload,
    unregister_then_close_cloud_connection_runtime,
)


class FakeLoop:
    def __init__(self, running=True):
        self.running = running

    def is_running(self):
        return self.running


class FakeWebSocket:
    def __init__(self, *, send_error=False, close_error=False):
        self.sent = []
        self.closed = False
        self.send_error = send_error
        self.close_error = close_error

    async def send(self, payload):
        if self.send_error:
            raise RuntimeError("send failed")
        self.sent.append(payload)

    async def close(self):
        if self.close_error:
            raise RuntimeError("close failed")
        self.closed = True


class CloudConnectionRuntimeTests(unittest.TestCase):
    def test_serialize_cloud_payload_keeps_unicode(self):
        self.assertEqual(serialize_cloud_payload({"text": "中文"}), '{"text": "中文"}')

    def test_send_cloud_payload_runtime_sends_json(self):
        ws = FakeWebSocket()

        result = asyncio.run(send_cloud_payload_runtime(ws, {"ok": True}))

        self.assertEqual(result, "sent")
        self.assertEqual(ws.sent, ['{"ok": true}'])

    def test_send_cloud_payload_runtime_swallows_send_errors(self):
        ws = FakeWebSocket(send_error=True)
        logs = []

        result = asyncio.run(send_cloud_payload_runtime(ws, {"ok": True}, log_error=logs.append))

        self.assertEqual(result, "error")
        self.assertTrue(any("send failed" in message for message in logs))

    def test_schedule_cloud_unregister_runtime_schedules_when_connected(self):
        calls = []

        async def send_unregister(ws, reason):
            return ws, reason

        result = schedule_cloud_unregister_runtime(
            reason="hidden",
            get_loop=lambda: FakeLoop(),
            get_ws=lambda: "ws",
            is_connected=lambda: True,
            send_unregister=send_unregister,
            run_coroutine_threadsafe=lambda coro, loop: calls.append((coro, loop)),
        )

        self.assertTrue(result)
        self.assertIsInstance(calls[0][1], FakeLoop)
        calls[0][0].close()

    def test_schedule_cloud_unregister_runtime_skips_unavailable_states(self):
        async def send_unregister(ws, reason):
            return ws, reason

        self.assertFalse(schedule_cloud_unregister_runtime(
            reason="hidden",
            get_loop=lambda: None,
            get_ws=lambda: "ws",
            is_connected=lambda: True,
            send_unregister=send_unregister,
            run_coroutine_threadsafe=lambda coro, loop: None,
        ))
        self.assertFalse(schedule_cloud_unregister_runtime(
            reason="hidden",
            get_loop=lambda: FakeLoop(running=False),
            get_ws=lambda: "ws",
            is_connected=lambda: True,
            send_unregister=send_unregister,
            run_coroutine_threadsafe=lambda coro, loop: None,
        ))
        self.assertFalse(schedule_cloud_unregister_runtime(
            reason="hidden",
            get_loop=lambda: FakeLoop(),
            get_ws=lambda: None,
            is_connected=lambda: True,
            send_unregister=send_unregister,
            run_coroutine_threadsafe=lambda coro, loop: None,
        ))
        self.assertFalse(schedule_cloud_unregister_runtime(
            reason="hidden",
            get_loop=lambda: FakeLoop(),
            get_ws=lambda: "ws",
            is_connected=lambda: False,
            send_unregister=send_unregister,
            run_coroutine_threadsafe=lambda coro, loop: None,
        ))

    def test_schedule_cloud_unregister_runtime_closes_coro_on_submit_failure(self):
        created = []
        logs = []

        async def send_unregister(ws, reason):
            return ws, reason

        def tracked_send_unregister(ws, reason):
            coro = send_unregister(ws, reason)
            created.append(coro)
            return coro

        def run_coroutine_threadsafe(_coro, _loop):
            raise RuntimeError("loop rejected")

        result = schedule_cloud_unregister_runtime(
            reason="hidden",
            get_loop=lambda: FakeLoop(),
            get_ws=lambda: "ws",
            is_connected=lambda: True,
            send_unregister=tracked_send_unregister,
            run_coroutine_threadsafe=run_coroutine_threadsafe,
            log_error=logs.append,
        )

        self.assertFalse(result)
        self.assertEqual(created[0].cr_frame, None)
        self.assertTrue(any("loop rejected" in message for message in logs))

    def test_unregister_then_close_cloud_connection_runtime_unregisters_when_enabled(self):
        ws = FakeWebSocket()
        calls = []

        async def send_unregister(seen_ws, reason):
            calls.append((seen_ws, reason))

        result = asyncio.run(unregister_then_close_cloud_connection_runtime(
            ws,
            reason="disconnect",
            auto_upload=True,
            send_unregister=send_unregister,
        ))

        self.assertEqual(result, (True, True))
        self.assertEqual(calls, [(ws, "disconnect")])
        self.assertTrue(ws.closed)

    def test_unregister_then_close_cloud_connection_runtime_closes_even_after_errors(self):
        ws = FakeWebSocket()
        logs = []

        async def send_unregister(seen_ws, reason):
            raise RuntimeError("unregister failed")

        result = asyncio.run(unregister_then_close_cloud_connection_runtime(
            ws,
            reason="disconnect",
            auto_upload=True,
            send_unregister=send_unregister,
            log_error=logs.append,
        ))

        self.assertEqual(result, (False, True))
        self.assertTrue(ws.closed)
        self.assertTrue(any("unregister failed" in message for message in logs))

    def test_unregister_then_close_cloud_connection_runtime_logs_close_errors(self):
        ws = FakeWebSocket(close_error=True)
        logs = []

        async def send_unregister(seen_ws, reason):
            return None

        result = asyncio.run(unregister_then_close_cloud_connection_runtime(
            ws,
            reason="disconnect",
            auto_upload=True,
            send_unregister=send_unregister,
            log_error=logs.append,
        ))

        self.assertEqual(result, (True, False))
        self.assertTrue(any("close failed" in message for message in logs))

    def test_reply_cloud_payload_runtime_attaches_identity_to_dict_payload(self):
        ws = FakeWebSocket()
        calls = []

        result = asyncio.run(reply_cloud_payload_runtime(
            ws,
            {"type": "pong"},
            identity_payload=lambda: {"imei": "861"},
            log=lambda message: calls.append(message),
        ))

        self.assertEqual(result, "sent")
        self.assertEqual(ws.sent, ['{"imei": "861", "type": "pong"}'])
        self.assertEqual(calls, [])

    def test_reply_cloud_payload_runtime_logs_send_errors(self):
        ws = FakeWebSocket(send_error=True)
        calls = []

        result = asyncio.run(reply_cloud_payload_runtime(
            ws,
            {"type": "pong"},
            identity_payload=lambda: {"imei": "861"},
            log=lambda message: calls.append(message),
        ))

        self.assertEqual(result, "error")
        self.assertTrue(any("send failed" in item for item in calls))


if __name__ == "__main__":
    unittest.main()
