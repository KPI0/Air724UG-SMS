import asyncio
import unittest

from sms_core.cloud_ws_runtime import (
    base_cloud_backoff,
    cloud_thread_main_runtime,
    cloud_ws_main_app_runtime,
    cloud_ws_main_runtime,
    wait_cloud_login_ack_runtime,
    wait_for_cloud_imei,
)


class FakeStopEvent:
    def __init__(self):
        self.stopped = False

    def is_set(self):
        return self.stopped

    def set(self):
        self.stopped = True


class FakeWebSocket:
    def __init__(self, messages=None):
        self.messages = list(messages or [])
        self.closed = False

    async def recv(self):
        if self.messages:
            return self.messages.pop(0)
        raise asyncio.TimeoutError()

    async def close(self):
        self.closed = True


class FakeConnectContext:
    def __init__(self, ws):
        self.ws = ws

    async def __aenter__(self):
        return self.ws

    async def __aexit__(self, exc_type, exc, tb):
        return False


class FakeLock:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


class FakeLoop:
    def __init__(self, *, fail=False, close_fail=False):
        self.fail = fail
        self.close_fail = close_fail
        self.closed = False
        self.runs = []

    def run_until_complete(self, value):
        self.runs.append(value)
        if self.fail:
            raise RuntimeError("loop failed")
        return value

    def close(self):
        self.closed = True
        if self.close_fail:
            raise RuntimeError("close failed")


class CloudWsRuntimeTests(unittest.TestCase):
    def test_base_cloud_backoff_normalizes_bad_values(self):
        self.assertEqual(base_cloud_backoff(5), 5.0)
        self.assertEqual(base_cloud_backoff(0), 1.0)
        self.assertEqual(base_cloud_backoff("bad"), 1.0)

    def test_wait_for_cloud_imei_requests_periodically(self):
        stop_event = FakeStopEvent()
        calls = []
        imei_ready = [False, True]

        async def sleep(delay):
            calls.append(("sleep", delay))

        last = asyncio.run(wait_for_cloud_imei(
            stop_event=stop_event,
            runtime_imei=lambda: imei_ready.pop(0),
            set_cloud_status=lambda *args: calls.append(("status", args)),
            request_cloud_device_imei=lambda: calls.append(("request",)),
            last_imei_request=0.0,
            monotonic=lambda: 6.0,
            sleep=sleep,
        ))

        self.assertEqual(last, 6.0)
        self.assertIn(("request",), calls)
        self.assertIn(("sleep", 0.5), calls)

    def test_wait_cloud_login_ack_runtime_authorizes_device(self):
        stop_event = FakeStopEvent()
        ws = FakeWebSocket([b'{"type":"device_login_ack","auth_status":"ok","message":"ok"}'])
        calls = []

        result = asyncio.run(wait_cloud_login_ack_runtime(
            ws,
            stop_event=stop_event,
            set_authorized=lambda value: calls.append(("authorized", value)),
            set_auth_status_from_ack=lambda data: calls.append(("ack", data["type"])),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            safe_preview=lambda value: f"safe:{value}",
            timeout=1.0,
            monotonic=lambda: 0.0,
        ))

        self.assertTrue(result)
        self.assertIn(("authorized", True), calls)
        self.assertIn(("ack", "device_login_ack"), calls)
        self.assertTrue(any(item[0] == "log" and item[2].get("show_main") for item in calls))

    def test_wait_cloud_login_ack_runtime_rejects_device(self):
        stop_event = FakeStopEvent()
        ws = FakeWebSocket(['{"type":"device_auth_result","status":"auth_failed","message":"bad"}'])
        calls = []

        result = asyncio.run(wait_cloud_login_ack_runtime(
            ws,
            stop_event=stop_event,
            set_authorized=lambda value: calls.append(("authorized", value)),
            set_auth_status_from_ack=lambda data: calls.append(("ack", data["status"])),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            safe_preview=lambda value: value,
            timeout=1.0,
            monotonic=lambda: 0.0,
        ))

        self.assertTrue(result)
        self.assertIn(("authorized", False), calls)
        self.assertIn(("ack", "auth_failed"), calls)
        self.assertTrue(any(item[0] == "log" and item[1][0] == "bad" for item in calls))

    def test_wait_cloud_login_ack_runtime_ignores_non_ack_and_invalid_json(self):
        stop_event = FakeStopEvent()
        ws = FakeWebSocket([
            "{bad",
            '{"type":"ping","value":1}',
            '{"type":"device_auth","auth_status":"ok"}',
        ])
        calls = []

        result = asyncio.run(wait_cloud_login_ack_runtime(
            ws,
            stop_event=stop_event,
            set_authorized=lambda value: calls.append(("authorized", value)),
            set_auth_status_from_ack=lambda data: calls.append(("ack", data["type"])),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            safe_preview=lambda value: f"preview:{value}",
            timeout=1.0,
            monotonic=lambda: 0.0,
        ))

        self.assertTrue(result)
        self.assertTrue(any(item[0] == "log" and "preview:" in item[1][0] for item in calls))
        self.assertIn(("authorized", True), calls)

    def test_wait_cloud_login_ack_runtime_times_out(self):
        stop_event = FakeStopEvent()
        ws = FakeWebSocket([])
        calls = []
        times = iter([0.0, 2.0])

        result = asyncio.run(wait_cloud_login_ack_runtime(
            ws,
            stop_event=stop_event,
            set_authorized=lambda value: calls.append(("authorized", value)),
            set_auth_status_from_ack=lambda data: calls.append(("ack", data)),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            safe_preview=lambda value: value,
            timeout=1.0,
            monotonic=lambda: next(times),
        ))

        self.assertFalse(result)
        self.assertEqual(calls[0], ("authorized", False))
        self.assertTrue(calls[1][2].get("show_main"))

    def test_cloud_ws_main_success_handles_message_then_stops(self):
        stop_event = FakeStopEvent()
        ws = FakeWebSocket(["hello"])
        calls = []

        async def sleep(delay):
            calls.append(("sleep", delay))

        async def wait_for(awaitable, timeout):
            return await awaitable

        async def send_register(seen_ws):
            calls.append(("register", seen_ws))

        async def wait_login_ack(seen_ws):
            return True

        async def handle_message(seen_ws, message):
            calls.append(("message", seen_ws, message))
            stop_event.set()

        asyncio.run(cloud_ws_main_runtime(
            "ws://server",
            2,
            stop_event=stop_event,
            runtime_imei=lambda: "861234567890123",
            request_cloud_device_imei=lambda: calls.append(("request",)),
            set_cloud_status=lambda *args: calls.append(("status", args)),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            connect=lambda *args, **kwargs: FakeConnectContext(ws),
            set_connection_state=lambda *args, **kwargs: calls.append(("state", args, kwargs)),
            reset_serial_log_state=lambda: calls.append(("reset",)),
            send_register=send_register,
            wait_login_ack=wait_login_ack,
            handle_message=handle_message,
            cloud_control_enabled=True,
            monotonic=lambda: 0.0,
            sleep=sleep,
            wait_for=wait_for,
            jitter=lambda start, end: 0.0,
        ))

        self.assertIn(("register", ws), calls)
        self.assertIn(("message", ws, "hello"), calls)
        self.assertIn(("reset",), calls)
        self.assertTrue(calls[-1][0] == "log")
        self.assertGreaterEqual(
            sum(1 for item in calls if item[0] == "log" and item[2].get("show_main")),
            2,
        )
        self.assertTrue(any(item[0] == "state" and item[2] == {"connected": True, "authorized": False} for item in calls))
        self.assertTrue(any(item[0] == "state" and item[2] == {"connected": False, "authorized": False} for item in calls))

    def test_cloud_ws_main_connection_error_backs_off_then_stops(self):
        stop_event = FakeStopEvent()
        calls = []

        def connect(*args, **kwargs):
            raise RuntimeError("down")

        async def sleep(delay):
            calls.append(("sleep", delay))
            if delay == 0.1:
                stop_event.set()

        asyncio.run(cloud_ws_main_runtime(
            "ws://server",
            1,
            stop_event=stop_event,
            runtime_imei=lambda: "861234567890123",
            request_cloud_device_imei=lambda: calls.append(("request",)),
            set_cloud_status=lambda *args: calls.append(("status", args)),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            connect=connect,
            set_connection_state=lambda *args, **kwargs: calls.append(("state", args, kwargs)),
            reset_serial_log_state=lambda: calls.append(("reset",)),
            send_register=lambda ws: calls.append(("register", ws)),
            wait_login_ack=lambda ws: True,
            handle_message=lambda ws, message: calls.append(("message", message)),
            cloud_control_enabled=False,
            monotonic=lambda: 0.0,
            sleep=sleep,
            jitter=lambda start, end: 0.0,
        ))

        self.assertIn(("sleep", 0.1), calls)
        self.assertTrue(any(item[0] == "log" and "down" in item[1][0] for item in calls))
        self.assertGreaterEqual(calls.count(("reset",)), 2)

    def test_cloud_thread_main_runtime_runs_and_clears_current_thread(self):
        calls = []
        thread = object()
        loop = FakeLoop()

        result = cloud_thread_main_runtime(
            "ws://server",
            3,
            lock=FakeLock(),
            set_loop=lambda value: calls.append(("loop", value)),
            get_thread=lambda: thread,
            set_thread=lambda value: calls.append(("thread", value)),
            run_main=lambda url, interval: ("main", url, interval),
            log=lambda message: calls.append(("log", message)),
            new_event_loop=lambda: loop,
            set_event_loop=lambda value: calls.append(("event_loop", value)),
            current_thread=lambda: thread,
        )

        self.assertEqual(result, "stopped")
        self.assertEqual(loop.runs, [("main", "ws://server", 3)])
        self.assertTrue(loop.closed)
        self.assertIn(("loop", loop), calls)
        self.assertIn(("loop", None), calls)
        self.assertIn(("thread", None), calls)

    def test_cloud_thread_main_runtime_logs_errors_and_ignores_close_errors(self):
        calls = []
        existing_thread = object()
        loop = FakeLoop(fail=True, close_fail=True)

        result = cloud_thread_main_runtime(
            "ws://server",
            3,
            lock=FakeLock(),
            set_loop=lambda value: calls.append(("loop", value)),
            get_thread=lambda: existing_thread,
            set_thread=lambda value: calls.append(("thread", value)),
            run_main=lambda url, interval: ("main", url, interval),
            log=lambda message: calls.append(("log", message)),
            new_event_loop=lambda: loop,
            set_event_loop=lambda value: calls.append(("event_loop", value)),
            current_thread=lambda: object(),
        )

        self.assertEqual(result, "error")
        self.assertTrue(any(item[0] == "log" and "loop failed" in item[1] for item in calls))
        self.assertTrue(loop.closed)
        self.assertNotIn(("thread", None), calls)

    def test_cloud_ws_main_app_runtime_forwards_dependencies_and_sets_state(self):
        calls = []
        stop_event = FakeStopEvent()

        async def run_ws_main(url, reconnect_interval, **kwargs):
            calls.append(("run", url, reconnect_interval, kwargs))
            kwargs["set_connection_state"]("ws", connected=1, authorized="")

        asyncio.run(cloud_ws_main_app_runtime(
            "ws://server",
            4,
            stop_event=stop_event,
            runtime_imei=lambda: "imei",
            request_cloud_device_imei=lambda: calls.append(("request",)),
            set_cloud_status=lambda *args: calls.append(("status", args)),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            connect=lambda *args, **kwargs: calls.append(("connect", args, kwargs)),
            set_ws=lambda value: calls.append(("ws", value)),
            set_connected=lambda value: calls.append(("connected", value)),
            set_authorized=lambda value: calls.append(("authorized", value)),
            reset_serial_log_state=lambda: calls.append(("reset",)),
            send_register=lambda ws: calls.append(("register", ws)),
            wait_login_ack=lambda ws: True,
            handle_message=lambda ws, message: calls.append(("message", ws, message)),
            cloud_control_enabled=True,
            monotonic=lambda: 1.0,
            run_ws_main=run_ws_main,
        ))

        self.assertEqual(calls[0][0:3], ("run", "ws://server", 4))
        forwarded = calls[0][3]
        self.assertIs(forwarded["stop_event"], stop_event)
        self.assertTrue(forwarded["cloud_control_enabled"])
        self.assertIn(("ws", "ws"), calls)
        self.assertIn(("connected", True), calls)
        self.assertIn(("authorized", False), calls)


if __name__ == "__main__":
    unittest.main()
