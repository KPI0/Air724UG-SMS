import asyncio
import unittest

from sms_core.cloud_connection_runtime import (
    reply_cloud_payload_runtime,
    schedule_cloud_payload_runtime,
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
    def test_payload_and_reply_errors_hide_transport_echo(self):
        class FailingWebSocket:
            async def send(self, _payload):
                raise RuntimeError("synthetic-payload-private-value")

        ws, logs = FailingWebSocket(), []
        sent = asyncio.run(send_cloud_payload_runtime(
            ws, {"secret": "synthetic-payload-private-value"}, log_error=logs.append,
        ))
        replied = asyncio.run(reply_cloud_payload_runtime(
            ws, {"type": "pong"}, identity_payload=lambda: {"imei": "000000000000001"},
            log=logs.append,
        ))

        self.assertEqual((sent, replied), ("error", "error"))
        self.assertEqual(len(logs), 2)
        self.assertTrue(all("RuntimeError" in message for message in logs))
        self.assertFalse(any("synthetic-payload-private-value" in message for message in logs))

    def test_teardown_and_scheduling_errors_hide_transport_echo_and_still_cleanup(self):
        calls, logs, created = [], [], []

        class FailingWebSocket:
            async def close(self):
                calls.append("close")
                raise RuntimeError("synthetic-close-private-value")

        async def unregister(_ws, _reason):
            calls.append("unregister")
            raise RuntimeError("synthetic-unregister-private-value")

        async def revoke(_ws, _reason):
            calls.append("revoke")
            raise RuntimeError("synthetic-revoke-private-value")

        ws = FailingWebSocket()
        result = asyncio.run(unregister_then_close_cloud_connection_runtime(
            ws, reason="disconnect", auto_upload=True, send_unregister=unregister,
            send_session_revoke=revoke, log_error=logs.append,
        ))
        self.assertEqual(result, (False, False))
        self.assertEqual(calls, ["unregister", "revoke", "close"])

        async def send(_ws, _payload):
            return None

        def tracked_send(*args):
            coro = send(*args)
            created.append(coro)
            return coro

        def rejected_schedule(_coro, _loop):
            raise RuntimeError("synthetic-schedule-private-value")

        dependencies = {
            "get_loop": lambda: FakeLoop(), "get_ws": lambda: ws,
            "is_connected": lambda: True, "run_coroutine_threadsafe": rejected_schedule,
            "log_error": logs.append,
        }
        self.assertFalse(schedule_cloud_payload_runtime(
            {"ok": True}, send_payload=tracked_send, **dependencies,
        ))
        self.assertFalse(schedule_cloud_unregister_runtime(
            reason="disconnect", send_unregister=tracked_send, **dependencies,
        ))
        self.assertEqual(len(created), 2)
        self.assertTrue(all(coro.cr_frame is None for coro in created))
        self.assertEqual(len(logs), 5)
        self.assertTrue(all("RuntimeError" in message for message in logs))
        self.assertFalse(any("private-value" in message for message in logs))

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
        self.assertIn("Send cloud payload failed: RuntimeError", logs)

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
        self.assertIn("Schedule cloud unregister failed: RuntimeError", logs)

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

    def test_unregister_then_close_revokes_session_even_when_device_is_hidden(self):
        ws = FakeWebSocket()
        calls = []

        async def send_unregister(_seen_ws, _reason):
            calls.append("unregister")

        async def send_session_revoke(seen_ws, reason):
            calls.append(("revoke", seen_ws, reason))

        result = asyncio.run(unregister_then_close_cloud_connection_runtime(
            ws,
            reason="password_changed",
            auto_upload=False,
            send_unregister=send_unregister,
            send_session_revoke=send_session_revoke,
        ))

        self.assertEqual(result, (False, True))
        self.assertEqual(calls, [("revoke", ws, "password_changed")])
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
        self.assertIn("Send cloud unregister failed: RuntimeError", logs)

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
        self.assertIn("Close cloud websocket failed: RuntimeError", logs)

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
        self.assertIn("回复云端失败：RuntimeError", calls)


if __name__ == "__main__":
    unittest.main()
