import asyncio
import threading
import unittest

from sms_core.cloud_message_namespace_runtime import run_registered_cloud_serial_command_namespace_runtime
from sms_core.threading_runtime import WorkerThreadRegistry


class UnprintableCommandError(RuntimeError):
    def __str__(self):
        raise RuntimeError("exception text unavailable")

    def __repr__(self):
        raise RuntimeError("exception repr unavailable")


def command_namespace():
    return {
        "serial_obj": object(), "serial_lock": threading.RLock(),
        "serial_connection_generation": 1, "cloud_ws_conn": object(),
        "cloud_connected": True, "cloud_device_authorized": True,
        "CLOUD_CONTROL_ENABLED": True, "CLOUD_DEVICE_SECRET": "fixture-key",
        "CLOUD_WS_URL": "wss://example.invalid/device",
        "_cloud_runtime_imei": lambda: "fixture-device",
        "SERIAL_COMMAND_THREAD_REGISTRY": WorkerThreadRegistry(),
        "log_file_only": lambda _message: None,
    }


class CloudCommandCompletionTests(unittest.IsolatedAsyncioTestCase):
    async def run_command(self, namespace, **kwargs):
        return await asyncio.wait_for(run_registered_cloud_serial_command_namespace_runtime(
            namespace, asyncio.get_running_loop(), "AT", **kwargs), 2)

    async def wait_workers(self, namespace):
        for worker in namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot():
            await asyncio.to_thread(worker.join, 2)
            self.assertFalse(worker.is_alive())
        self.assertFalse(namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot())

    async def test_worker_error_with_unprintable_diagnostics_completes(self):
        namespace = command_namespace()

        def send(*_args, **_kwargs):
            raise UnprintableCommandError()

        namespace["_cloud_send_serial_command"] = send
        self.assertEqual(await self.run_command(namespace),
                         (False, "发送失败：UnprintableCommandError"))
        await self.wait_workers(namespace)

    async def test_worker_error_omits_untrusted_exception_text(self):
        namespace = command_namespace()

        def send(*_args, **_kwargs):
            raise OSError("fixture exception payload")

        namespace["_cloud_send_serial_command"] = send
        self.assertEqual(await self.run_command(namespace), (False, "发送失败：OSError"))
        await self.wait_workers(namespace)

    async def test_start_error_with_unprintable_diagnostics_returns_failure(self):
        namespace = command_namespace()

        def start(*_args, **_kwargs):
            raise UnprintableCommandError()

        self.assertEqual(await self.run_command(namespace, start_worker=start),
                         (False, "发送失败：UnprintableCommandError"))
        self.assertFalse(namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot())

    async def test_health_and_log_errors_preserve_original_failure(self):
        namespace = command_namespace()

        class Health:
            def record(self, *_args, **_kwargs):
                return {"request_reconnect": True, "modem_unresponsive": True,
                        "consecutive_at_timeouts": 3}

        def fail_log(*_args, **_kwargs):
            raise UnprintableCommandError()

        namespace.update({"CLOUD_MODEM_HEALTH": Health(), "_cloud_log": fail_log,
                          "log_file_only": fail_log,
                          "_cloud_send_serial_command": lambda *_a, **_k: (False, "AT response timeout")})
        self.assertEqual(await self.run_command(namespace), (False, "AT response timeout"))
        await self.wait_workers(namespace)

    async def test_normal_success_keeps_health_metadata(self):
        namespace = command_namespace()
        recorded = []

        class Health:
            def record(self, ok, info, **kwargs):
                recorded.append((ok, info, kwargs))
                return {"request_reconnect": False, "modem_unresponsive": False,
                        "consecutive_at_timeouts": 0}

        namespace.update({"CLOUD_MODEM_HEALTH": Health(),
                          "_cloud_send_serial_command": lambda *_a, **_k: (True, "OK")})
        self.assertEqual(await self.run_command(namespace),
                         (True, "OK", {"modem_unresponsive": False, "consecutive_at_timeouts": 0}))
        self.assertEqual(recorded, [(True, "OK", {"connection": namespace["serial_obj"], "generation": 1})])
        await self.wait_workers(namespace)

    async def test_connection_change_in_worker_preserves_new_health(self):
        namespace = command_namespace()
        replacement = object()

        class Health:
            def record(self, *_args, **_kwargs):
                raise AssertionError("old command must not update replacement connection")

        def send(*_args, **_kwargs):
            namespace["serial_obj"] = replacement
            namespace["serial_connection_generation"] = 2
            return False, "original command failed"

        namespace.update({"CLOUD_MODEM_HEALTH": Health(), "_cloud_send_serial_command": send})
        self.assertEqual(await self.run_command(namespace), (False, "original command failed"))
        self.assertIs(namespace["serial_obj"], replacement)
        await self.wait_workers(namespace)

    async def test_cancelled_future_ignores_late_worker_completion(self):
        namespace = command_namespace()
        started, release = threading.Event(), threading.Event()
        loop_errors = []
        loop = asyncio.get_running_loop()
        previous_handler = loop.get_exception_handler()
        loop.set_exception_handler(lambda _loop, context: loop_errors.append(context))

        def send(*_args, **_kwargs):
            started.set()
            release.wait(2)
            return True, "OK"

        namespace["_cloud_send_serial_command"] = send
        task = asyncio.create_task(run_registered_cloud_serial_command_namespace_runtime(namespace, loop, "AT"))
        try:
            self.assertTrue(await asyncio.to_thread(started.wait, 2))
            task.cancel()
            with self.assertRaises(asyncio.CancelledError):
                await task
            release.set()
            await self.wait_workers(namespace)
            await asyncio.sleep(0)
            self.assertFalse(loop_errors)
        finally:
            release.set()
            loop.set_exception_handler(previous_handler)


class CloudCommandClosedLoopTests(unittest.TestCase):
    def test_worker_completion_after_loop_close_returns_without_error(self):
        namespace = command_namespace()
        pending_workers = []
        loop = asyncio.new_event_loop()

        def start(_name, target, **_kwargs):
            pending_workers.append(target)

        namespace["_cloud_send_serial_command"] = lambda *_a, **_k: (False, "original failure")
        task = loop.create_task(run_registered_cloud_serial_command_namespace_runtime(
            namespace, loop, "AT", start_worker=start))
        try:
            loop.run_until_complete(asyncio.sleep(0))
            self.assertEqual(len(pending_workers), 1)
            task.cancel()
            with self.assertRaises(asyncio.CancelledError):
                loop.run_until_complete(task)
        finally:
            loop.close()
        pending_workers[0]()
