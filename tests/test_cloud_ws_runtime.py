import asyncio
import unittest

from sms_core.cloud_ws_runtime import (
    base_cloud_backoff,
    build_cloud_ssl_context,
    cloud_thread_main_runtime,
    cloud_websocket_connect_kwargs,
    cloud_ws_main_app_runtime,
    cloud_ws_main_runtime,
    wait_cloud_login_ack_runtime,
    wait_for_cloud_auth_retry_runtime,
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


async def _async_none():
    return None


async def _async_true():
    return True


async def _async_false():
    return False


class CloudWsRuntimeTests(unittest.TestCase):
    def test_build_cloud_ssl_context_skips_plain_websocket(self):
        calls = []

        result = build_cloud_ssl_context(
            "ws://server/ws/device",
            create_default_context=lambda: calls.append("context"),
            certifi_where=lambda: calls.append("certifi"),
        )

        self.assertIsNone(result)
        self.assertEqual(calls, [])

    def test_build_cloud_ssl_context_merges_system_and_certifi_roots(self):
        calls = []

        class FakeContext:
            def load_verify_locations(self, *, cafile):
                calls.append(("load", cafile))

        context = FakeContext()
        result = build_cloud_ssl_context(
            "wss://server/ws/device",
            create_default_context=lambda: calls.append(("context",)) or context,
            certifi_where=lambda: "certifi-cacert.pem",
        )

        self.assertIs(result, context)
        self.assertEqual(calls, [("context",), ("load", "certifi-cacert.pem")])

    def test_cloud_websocket_connect_kwargs_adds_ssl_only_for_wss(self):
        self.assertEqual(
            cloud_websocket_connect_kwargs(
                "ws://server/ws/device",
                ssl_context_factory=lambda _url: None,
            ),
            {"ping_interval": 30, "ping_timeout": 30},
        )
        self.assertEqual(
            cloud_websocket_connect_kwargs(
                "wss://server/ws/device",
                ssl_context_factory=lambda _url: "tls-context",
            ),
            {
                "ping_interval": 30,
                "ping_timeout": 30,
                "ssl": "tls-context",
            },
        )

    def test_base_cloud_backoff_normalizes_bad_values(self):
        self.assertEqual(base_cloud_backoff(5), 5.0)
        self.assertEqual(base_cloud_backoff(0), 1.0)
        self.assertEqual(base_cloud_backoff("bad"), 1.0)

    def test_cloud_ws_main_uses_current_enabled_state_when_stopping(self):
        stop_event = FakeStopEvent()
        stop_event.set()
        statuses = []
        enabled = {"value": True}
        enabled["value"] = False

        asyncio.run(cloud_ws_main_runtime(
            "ws://server",
            1,
            stop_event=stop_event,
            runtime_imei=lambda: "imei",
            request_cloud_device_imei=lambda: None,
            set_cloud_status=lambda text, color: statuses.append((text, color)),
            log=lambda *_args, **_kwargs: None,
            connect=lambda *_args, **_kwargs: None,
            set_connection_state=lambda *_args, **_kwargs: None,
            reset_serial_log_state=lambda: None,
            send_register=lambda _ws: None,
            wait_login_ack=lambda _ws: True,
            handle_message=lambda _ws, _message: None,
            cloud_control_enabled=lambda: enabled["value"],
            monotonic=lambda: 0.0,
        ))

        self.assertEqual(statuses[-1], ("🌐 已关闭", "#666666"))

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

        self.assertEqual(result, "auth_failed")
        self.assertIn(("authorized", False), calls)
        self.assertIn(("ack", "auth_failed"), calls)
        self.assertTrue(any(item[0] == "log" and item[1][0] == "bad" for item in calls))

    def test_wait_cloud_login_ack_runtime_keeps_waiting_device_connected(self):
        stop_event = FakeStopEvent()
        ws = FakeWebSocket([
            '{"type":"device_login_ack","ok":false,"status":"waiting",'
            '"auth_status":"waiting","message":"waiting"}'
        ])
        calls = []

        result = asyncio.run(wait_cloud_login_ack_runtime(
            ws,
            stop_event=stop_event,
            set_authorized=lambda value: calls.append(("authorized", value)),
            set_auth_status_from_ack=lambda data: calls.append(("ack", data["auth_status"])),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            safe_preview=lambda value: value,
            timeout=1.0,
            monotonic=lambda: 0.0,
        ))

        self.assertTrue(result)
        self.assertFalse(ws.closed)
        self.assertIn(("authorized", False), calls)
        self.assertIn(("ack", "waiting"), calls)
        self.assertTrue(any(item[0] == "log" and item[1][0] == "waiting" for item in calls))

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
            runtime_imei=lambda: "123456789012345",
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
            schedule_pending_sms_events=lambda: calls.append(("schedule_sms",)),
        ))

        self.assertIn(("register", ws), calls)
        self.assertIn(("message", ws, "hello"), calls)
        self.assertEqual(calls.count(("schedule_sms",)), 2)
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
            runtime_imei=lambda: "123456789012345",
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

    def test_cloud_ws_main_auth_failure_waits_for_web_password_retry(self):
        stop_event = FakeStopEvent()
        ws = FakeWebSocket([
            '{"type":"device_auth_result","status":"auth_failed","message":"bad"}',
            '{"type":"device_auth_retry"}',
            '{"type":"device_login_ack","auth_status":"ok","message":"ok"}',
        ])
        calls = []

        async def wait_for(awaitable, timeout):
            return await awaitable

        login_attempts = [0]

        async def wait_login_ack(seen_ws):
            login_attempts[0] += 1
            result = await wait_cloud_login_ack_runtime(
                seen_ws,
                stop_event=stop_event,
                set_authorized=lambda value: calls.append(("authorized", value)),
                set_auth_status_from_ack=lambda data: calls.append(("ack", data.get("status"))),
                log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
                safe_preview=lambda value: value,
                timeout=1.0,
                monotonic=lambda: 0.0,
                wait_for=wait_for,
            )
            if login_attempts[0] == 2:
                stop_event.set()
            return result

        async def send_register(seen_ws):
            calls.append(("register", seen_ws))

        async def handle_message(*args):
            calls.append(("message", args))

        async def sleep(_delay):
            return None

        awaitable = cloud_ws_main_runtime(
            "ws://server",
            1,
            stop_event=stop_event,
            runtime_imei=lambda: "123456789012345",
            request_cloud_device_imei=lambda: None,
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
            jitter=lambda start, end: 0.0,
        )
        asyncio.run(awaitable)

        self.assertFalse(ws.closed)
        self.assertFalse(any(item[0] == "message" for item in calls))
        state_indices = [index for index, item in enumerate(calls) if item[0] == "state"]
        self.assertGreaterEqual(len(state_indices), 2)
        self.assertEqual(calls[state_indices[0]][2], {"connected": True, "authorized": False})
        self.assertEqual(calls[state_indices[1]][2], {"connected": False, "authorized": False})
        self.assertIn(("reset",), calls)
        self.assertEqual([item[0] for item in calls if item[0] == "register"], ["register", "register"])
        self.assertTrue(any(
            item[0] == "status"
            and item[1][0] == "🌐 授权失败"
            for item in calls
        ))

    def test_cloud_auth_retry_keeps_waiting_after_another_explicit_rejection(self):
        stop_event = FakeStopEvent()
        ws = FakeWebSocket([
            '{"type":"device_auth_retry"}',
            '{"type":"device_auth_retry"}',
        ])
        calls = []
        login_results = iter(["auth_failed", True])

        async def wait_login_ack(_ws):
            return next(login_results)

        async def send_register(_ws):
            calls.append(("register",))

        result = asyncio.run(wait_for_cloud_auth_retry_runtime(
            ws,
            stop_event=stop_event,
            send_register=send_register,
            wait_login_ack=wait_login_ack,
            set_connection_state=lambda *args, **kwargs: calls.append(("state", args, kwargs)),
            set_cloud_status=lambda *args: calls.append(("status", args)),
            reset_serial_log_state=lambda: calls.append(("reset",)),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            wait_for=lambda awaitable, timeout: awaitable,
        ))

        self.assertTrue(result)
        self.assertEqual(calls.count(("register",)), 2)
        self.assertIn(("reset",), calls)
        self.assertTrue(any(
            item[0] == "state"
            and item[2] == {"connected": False, "authorized": False}
            for item in calls
        ))

    def test_cloud_ws_main_reconnects_when_auth_hold_transport_drops(self):
        stop_event = FakeStopEvent()
        calls = []
        sockets = []

        class DroppedHoldWebSocket(FakeWebSocket):
            async def recv(self):
                raise ConnectionError("hold transport dropped")

        def connect(*_args, **_kwargs):
            ws = DroppedHoldWebSocket() if not sockets else FakeWebSocket()
            sockets.append(ws)
            return FakeConnectContext(ws)

        async def wait_login_ack(_ws):
            if len(sockets) == 1:
                return "auth_failed"
            stop_event.set()
            return True

        async def sleep(delay):
            calls.append(("sleep", delay))

        asyncio.run(cloud_ws_main_runtime(
            "ws://server",
            1,
            stop_event=stop_event,
            runtime_imei=lambda: "123456789012345",
            request_cloud_device_imei=lambda: None,
            set_cloud_status=lambda *args: calls.append(("status", args)),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            connect=connect,
            set_connection_state=lambda *args, **kwargs: calls.append(("state", args, kwargs)),
            reset_serial_log_state=lambda: calls.append(("reset",)),
            send_register=lambda _ws: _async_none(),
            wait_login_ack=wait_login_ack,
            handle_message=lambda *_args: _async_none(),
            cloud_control_enabled=True,
            monotonic=lambda: 0.0,
            sleep=sleep,
            jitter=lambda _start, _end: 0.0,
        ))

        self.assertEqual(len(sockets), 2)
        self.assertIn(("sleep", 0.1), calls)
        self.assertTrue(any(
            item[0] == "log" and "hold transport dropped" in item[1][0]
            for item in calls
        ))

    def test_cloud_ws_main_clears_initial_auth_state_before_close(self):
        stop_event = FakeStopEvent()
        states = []
        close_states = []

        class InspectCloseWebSocket(FakeWebSocket):
            async def close(self):
                close_states.append(states[-1])
                self.closed = True

        ws = InspectCloseWebSocket()

        async def sleep(delay):
            if delay == 0.1:
                stop_event.set()

        asyncio.run(cloud_ws_main_runtime(
            "ws://server",
            1,
            stop_event=stop_event,
            runtime_imei=lambda: "123456789012345",
            request_cloud_device_imei=lambda: None,
            set_cloud_status=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            connect=lambda *_args, **_kwargs: FakeConnectContext(ws),
            set_connection_state=lambda value, **kwargs: states.append((value, kwargs)),
            reset_serial_log_state=lambda: None,
            send_register=lambda _ws: _async_none(),
            wait_login_ack=lambda _ws: _async_false(),
            handle_message=lambda *_args: _async_none(),
            cloud_control_enabled=True,
            monotonic=lambda: 0.0,
            sleep=sleep,
            jitter=lambda _start, _end: 0.0,
        ))

        self.assertTrue(ws.closed)
        self.assertEqual(close_states, [(None, {"connected": False, "authorized": False})])

    def test_cloud_ws_main_reauth_failure_disconnects_without_reconnect(self):
        stop_event = FakeStopEvent()
        states = []
        close_states = []
        calls = []

        class InspectCloseWebSocket(FakeWebSocket):
            async def close(self):
                close_states.append(states[-1])
                self.closed = True

        ws = InspectCloseWebSocket(["reauth-failed"])

        async def sleep(delay):
            calls.append(("sleep", delay))
            if delay == 0.1:
                stop_event.set()

        async def handle_message(_ws, message):
            calls.append(("message", message))
            return "auth_failed"

        asyncio.run(cloud_ws_main_runtime(
            "ws://server",
            1,
            stop_event=stop_event,
            runtime_imei=lambda: "123456789012345",
            request_cloud_device_imei=lambda: None,
            set_cloud_status=lambda *args: calls.append(("status", args)),
            log=lambda *_args, **_kwargs: None,
            connect=lambda *_args, **_kwargs: FakeConnectContext(ws),
            set_connection_state=lambda value, **kwargs: states.append((value, kwargs)),
            reset_serial_log_state=lambda: calls.append(("reset",)),
            send_register=lambda _ws: _async_none(),
            wait_login_ack=lambda _ws: _async_true(),
            handle_message=handle_message,
            cloud_control_enabled=True,
            monotonic=lambda: 0.0,
            sleep=sleep,
            wait_for=lambda awaitable, timeout: awaitable,
            jitter=lambda _start, _end: 0.0,
        ))

        self.assertIn(("message", "reauth-failed"), calls)
        self.assertTrue(ws.closed)
        self.assertEqual(close_states, [(None, {"connected": False, "authorized": False})])
        self.assertFalse(any(item[0] == "sleep" for item in calls))
        self.assertTrue(any(
            item[0] == "status"
            and item[1][0] == "🌐 授权失败"
            for item in calls
        ))

    def test_cloud_ws_main_auth_failure_stops_when_hold_channel_is_stopped(self):
        stop_event = FakeStopEvent()
        calls = []

        ws = FakeWebSocket()

        async def sleep(_delay):
            return None

        async def wait_login_ack(_ws):
            stop_event.set()
            return "auth_failed"

        asyncio.run(cloud_ws_main_runtime(
            "ws://server",
            1,
            stop_event=stop_event,
            runtime_imei=lambda: "123456789012345",
            request_cloud_device_imei=lambda: None,
            set_cloud_status=lambda *args: calls.append(("status", args)),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            connect=lambda *_args, **_kwargs: FakeConnectContext(ws),
            set_connection_state=lambda *_args, **_kwargs: None,
            reset_serial_log_state=lambda: None,
            send_register=lambda _ws: _async_none(),
            wait_login_ack=wait_login_ack,
            handle_message=lambda *_args: _async_none(),
            cloud_control_enabled=True,
            monotonic=lambda: 0.0,
            sleep=sleep,
        ))

        self.assertTrue(stop_event.is_set())
        self.assertTrue(any(
            item[0] == "status" and item[1][0] == "🌐 授权失败"
            for item in calls
        ))

    def test_cloud_ws_main_auth_failures_increase_backoff_before_reset(self):
        stop_event = FakeStopEvent()
        websockets = []
        sleeps = []
        short_sleep_count = [0]

        async def sleep(delay):
            sleeps.append(delay)
            if delay == 0.1:
                short_sleep_count[0] += 1
                if short_sleep_count[0] >= 50:
                    stop_event.set()
            await asyncio.sleep(0)

        def connect(*args, **kwargs):
            ws = FakeWebSocket()
            websockets.append(ws)
            return FakeConnectContext(ws)

        async def wait_login_ack(_ws):
            return False

        async def send_register(_ws):
            return None

        async def run():
            await asyncio.wait_for(
                cloud_ws_main_runtime(
                    "ws://server",
                    2,
                    stop_event=stop_event,
                    runtime_imei=lambda: "123456789012345",
                    request_cloud_device_imei=lambda: None,
                    set_cloud_status=lambda *_args: None,
                    log=lambda *_args, **_kwargs: None,
                    connect=connect,
                    set_connection_state=lambda *_args, **_kwargs: None,
                    reset_serial_log_state=lambda: None,
                    send_register=send_register,
                    wait_login_ack=wait_login_ack,
                    handle_message=lambda *_args: None,
                    cloud_control_enabled=False,
                    monotonic=lambda: 0.0,
                    sleep=sleep,
                    jitter=lambda _start, _end: 0.25,
                ),
                timeout=1.0,
            )

        asyncio.run(run())

        self.assertEqual(sleeps.count(0.1), 50)
        self.assertEqual(sleeps.count(0.25), 1)
        self.assertEqual(len(websockets), 2)
        self.assertTrue(all(ws.closed for ws in websockets))

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
