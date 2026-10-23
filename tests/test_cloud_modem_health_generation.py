import asyncio
import threading
import unittest

from sms_core.cloud_command_context import capture_cloud_serial_command_context
from sms_core.cloud_modem_health import CloudModemHealthState
from sms_core.cloud_message_namespace_runtime import (
    apply_cloud_modem_health_namespace_runtime, run_registered_cloud_serial_command_namespace_runtime,
)
from sms_ui.app_infrastructure_namespace_runtime import safe_close_serial_namespace_runtime
from tests.test_cloud_permission_submission import fixture


TIMEOUT = (False, "等待 Modem 指令响应超时")


def health_fixture(threshold=3):
    namespace, modem = fixture()
    calls = []
    modem.close = lambda: (calls.append(("close", modem)), setattr(modem, "is_open", False))
    namespace.update(CLOUD_MODEM_HEALTH=CloudModemHealthState(threshold),
        SERIAL_COMMAND_THREAD_REGISTRY=object(), serial_read_lock=threading.Lock(),
        serial_wakeup_event=threading.Event(), unlock_port_mutex=lambda: calls.append(("unlock",)),
        _cloud_send_serial_command=lambda *_args, **_kwargs: TIMEOUT)
    namespace["safe_close_serial"] = lambda **kwargs: safe_close_serial_namespace_runtime(namespace, **kwargs)
    return namespace, modem, calls


class CloudHealthGenerationTests(unittest.IsolatedAsyncioTestCase):
    async def test_delayed_timeout_does_not_close_or_count_against_replacement(self):
        for same_serial in (False, True):
            with self.subTest(same_serial=same_serial):
                ns, modem, calls = health_fixture(1)
                _, replacement, _ = health_fixture()

                def start_worker(_name, worker, **_kwargs):
                    worker()
                    ns["serial_obj"] = modem if same_serial else replacement
                    ns["serial_connection_generation"] += 1

                await run_registered_cloud_serial_command_namespace_runtime(
                    ns, asyncio.get_running_loop(), "ATI", start_worker=start_worker)
                self.assertEqual(calls, [])
                self.assertTrue(ns["serial_obj"].is_open)
                self.assertFalse(ns["serial_wakeup_event"].is_set())
                self.assertEqual(ns["CLOUD_MODEM_HEALTH"].snapshot()["consecutive_at_timeouts"], 0)

    async def test_three_current_timeouts_still_close_and_wake_original(self):
        ns, modem, calls = health_fixture()
        for index in range(3):
            result = await run_registered_cloud_serial_command_namespace_runtime(
                ns, asyncio.get_running_loop(), "ATI",
                start_worker=lambda _name, worker, **_kwargs: worker())
            self.assertEqual(result[2]["consecutive_at_timeouts"], index + 1)
        self.assertEqual(calls, [("close", modem), ("unlock",)])
        self.assertFalse(modem.is_open)
        self.assertIsNone(ns["serial_obj"])
        self.assertTrue(ns["serial_wakeup_event"].is_set())


class ScopedHealthTests(unittest.TestCase):
    def test_connection_change_starts_a_new_timeout_streak(self):
        ns, modem, calls = health_fixture()
        context = capture_cloud_serial_command_context(ns, command="ATI")
        for _ in range(2):
            apply_cloud_modem_health_namespace_runtime(ns, TIMEOUT, command_context=context)
        ns["serial_connection_generation"] += 1
        new_context = capture_cloud_serial_command_context(ns, command="ATI")
        result = apply_cloud_modem_health_namespace_runtime(ns, TIMEOUT, command_context=new_context)
        self.assertEqual(result[2]["consecutive_at_timeouts"], 1)
        self.assertEqual(calls, [])
        self.assertTrue(modem.is_open)

    def test_reconnect_rechecks_generation_after_read_lock_wait(self):
        ns, modem, calls = health_fixture(1)
        context = capture_cloud_serial_command_context(ns, command="ATI")
        reached, release = threading.Event(), threading.Event()
        original_lock = ns["serial_read_lock"]

        class ReadGate:
            def __enter__(self):
                reached.set()
                if not release.wait(3):
                    raise RuntimeError("test read gate timed out")
                return original_lock.__enter__()

            def __exit__(self, *args):
                return original_lock.__exit__(*args)

        ns["serial_read_lock"] = ReadGate()
        results, errors = [], []

        def apply():
            try:
                results.append(apply_cloud_modem_health_namespace_runtime(ns, TIMEOUT, command_context=context))
            except Exception as exc:
                errors.append(exc)

        worker = threading.Thread(target=apply)
        worker.start()
        try:
            self.assertTrue(reached.wait(3))
            with ns["serial_lock"]:
                ns["serial_connection_generation"] += 1
        finally:
            release.set()
            worker.join(3)
        self.assertFalse(worker.is_alive())
        self.assertEqual(errors, [])
        self.assertEqual(calls, [])
        self.assertTrue(modem.is_open)
        self.assertFalse(ns["serial_wakeup_event"].is_set())

    def test_guarded_close_never_cancels_new_transaction_or_notifies_offline(self):
        ns, modem, calls = health_fixture()
        _, replacement, _ = health_fixture()
        ns["serial_obj"] = replacement
        original_generation = ns["serial_connection_generation"]
        ns["_notify_cloud_channel_status"] = lambda *_args: calls.append(("notify",))
        result = safe_close_serial_namespace_runtime(
            ns, expected_serial=modem, expected_generation=original_generation)
        self.assertIs(result, False)
        self.assertEqual(calls, [])
        self.assertIs(ns["serial_obj"], replacement)
        self.assertEqual(ns["serial_connection_generation"], original_generation)

    def test_guarded_close_rechecks_after_final_connection_lock_wait(self):
        ns, modem, calls = health_fixture()
        _, replacement, _ = health_fixture()
        original_generation = ns["serial_connection_generation"]
        reached, release = threading.Event(), threading.Event()
        original_lock = ns["serial_lock"]

        class ConnectionGate:
            def __enter__(self):
                reached.set()
                if not release.wait(3):
                    raise RuntimeError("test connection gate timed out")
                return original_lock.__enter__()

            def __exit__(self, *args):
                return original_lock.__exit__(*args)

        ns["serial_lock"] = ConnectionGate()
        waiter = ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"].begin("ATI", replacement)
        results = []
        worker = threading.Thread(target=lambda: results.append(safe_close_serial_namespace_runtime(
            ns, expected_serial=modem, expected_generation=original_generation)))
        worker.start()
        try:
            self.assertTrue(reached.wait(3))
            ns["serial_obj"] = replacement
            ns["serial_connection_generation"] += 1
        finally:
            release.set()
            worker.join(3)
        self.assertFalse(worker.is_alive())
        self.assertEqual(results, [False])
        self.assertFalse(waiter.done())
        self.assertEqual(calls, [])
        self.assertTrue(replacement.is_open)
        ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"].finish(waiter)

    def test_sms_isolation_is_not_automatically_reopened_by_timeout(self):
        from sms_core import serial_sender
        ns, modem, calls = health_fixture(1)
        context = capture_cloud_serial_command_context(ns, command="ATI")
        with serial_sender.DEFAULT_SERIAL_WRITE_LOCK:
            serial_sender._isolate_sms_serial(modem)
        try:
            apply_cloud_modem_health_namespace_runtime(ns, TIMEOUT, command_context=context)
            self.assertEqual(calls, [])
            self.assertTrue(modem.is_open)
        finally:
            with serial_sender.DEFAULT_SERIAL_WRITE_LOCK:
                del modem._sms_recovery_error
