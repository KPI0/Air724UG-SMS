import asyncio
from functools import partial
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from sms_core import serial_sender
from sms_core.call_command_context import SerialCallStateSnapshot
from sms_ui.call_popup_namespace_runtime import (
    finish_incoming_call_session_namespace_runtime, show_call_popup_namespace_runtime,
)
from sms_ui.call_popup_runtime import show_call_popup_app_runtime, show_call_popup_runtime
from tests import test_cloud_namespace_bindings as cloud_fixtures


class CallPopupModemConfirmationTests(unittest.TestCase):
    def make_popup(self, response="OK", stage="normal", before_result=None, response_timeout=0.02):
        fixture = cloud_fixtures.CloudCallHistoryDeliveryTests("runTest")
        fixture.setUp()
        fixture.start_call()
        namespace = fixture.namespace
        coordinator = serial_sender.AtCommandResponseCoordinator()
        self.wrote = threading.Event()
        self.ui, self.workers, self.results, self.restores = [], [], [], []
        showing = [True]

        class Modem:
            is_open = True

            def __init__(self):
                self.writes = []

            def write(self, payload):
                self.writes.append(payload)
                self_test.wrote.set()
                if stage == "write-error":
                    raise OSError("synthetic write failure")
                if stage == "partial":
                    return 3
                if stage == "cancel":
                    coordinator.cancel_active("synthetic cancellation")
                elif response is not None:
                    coordinator.observe_line(response, connection=self,
                        call_state_snapshot=SerialCallStateSnapshot(namespace))
                return len(payload)

            def flush(self):
                if stage == "flush-error":
                    raise OSError("synthetic flush failure")

        modem = Modem()
        self_test = self

        def on_ui(callback, _post):
            if showing[0]:
                return callback()
            self.ui.append(callback)

        def send(*args, **kwargs):
            callback = kwargs["on_result"]

            def observe(result):
                self.results.append(result)
                if before_result:
                    before_result(namespace)
                callback(result)

            kwargs["on_result"] = observe
            if "response_coordinator" in kwargs:
                kwargs["response_timeout"] = response_timeout
            worker = serial_sender.send_command_with_result_async(*args, **kwargs)
            self.workers.append(worker)
            return worker

        def open_popup(_parent, caller, _center, answer, hangup, *_):
            return SimpleNamespace(answer=answer, hangup=hangup,
                winfo_exists=lambda: True, _call_popup_caller_num=caller)

        namespace.update({"root": None, "serial_obj": modem, "serial_lock": threading.RLock(),
            "serial_connection_generation": 1, "current_call_popup": None,
            "SERIAL_COMMAND_RESPONSE_COORDINATOR": coordinator,
            "center_window": lambda *_: None, "port_ui": lambda *_: None,
            "set_status": lambda *_: None, "run_on_ui_thread": on_ui,
            "ui_post": self.ui.append, "close_call_popup": lambda: None})
        popup = show_call_popup_namespace_runtime(namespace, "synthetic-caller", "call-a",
            show_app_runtime=partial(show_call_popup_app_runtime,
                show_runtime=partial(show_call_popup_runtime, open_popup=open_popup,
                                     send_command_async=send)))
        showing[0] = False
        self.fixture, self.namespace, self.modem, self.coordinator = fixture, namespace, modem, coordinator
        return popup

    def click(self, popup, command):
        if command == "ATA":
            popup.answer(lambda: None, lambda: self.restores.append(command))
        else:
            popup.hangup(lambda: self.restores.append(command))
        for worker in self.workers:
            worker.join(2)
            self.assertFalse(worker.is_alive())
        self.assertEqual(len(self.results), 1)

    def history_only(self):
        asyncio.run(self.namespace["_cloud_send_payload"](self.fixture.ws, self.fixture.payload()))
        return self.fixture.writes[-1][1]["history_only"]

    def drain_ui(self):
        while self.ui:
            self.ui.pop(0)()

    def test_modem_error_keeps_same_ringing_call_live_and_allows_retry(self):
        for command in ("ATA", "ATH"):
            for response in ("ERROR", "+CME ERROR: 3"):
                with self.subTest(command=command, response=response):
                    popup = self.make_popup(response)
                    self.click(popup, command)
                    self.assertFalse(self.results[0].ok)
                    self.assertFalse(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)
                    self.assertFalse(self.history_only())
                    self.assertFalse(serial_sender.serial_sms_recovery_error(self.modem))
                    self.drain_ui()
                    self.assertEqual(self.restores, [command])

    def test_only_own_modem_ok_confirms_before_ui_delivery(self):
        for command in ("ATA", "ATH"):
            with self.subTest(command=command):
                popup = self.make_popup()
                self.click(popup, command)
                self.assertTrue(self.results[0].ok)
                self.assertTrue(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)
                self.assertTrue(self.ui)
                self.assertTrue(self.history_only())

    def test_unknown_attempt_isolates_original_serial_and_keeps_history_notification_live(self):
        for stage in ("write-error", "partial", "flush-error", "normal", "cancel"):
            for command in ("ATA", "ATH"):
                with self.subTest(stage=stage, command=command):
                    popup = self.make_popup(None, stage)
                    self.click(popup, command)
                    self.assertFalse(self.results[0].ok)
                    self.assertIn("重启设备", self.results[0].error)
                    self.assertIn("通话", serial_sender.serial_sms_recovery_error(self.modem))
                    self.assertFalse(self.history_only())
                    before = list(self.modem.writes)
                    self.assertFalse(serial_sender.write_serial_command_result(self.modem, "ATI").ok)
                    self.assertEqual(self.modem.writes, before)
                    self.assertNotIn(b"\x1b", self.modem.writes)

    def test_trusted_terminal_during_write_survives_flush_error(self):
        for response in ("OK", "ERROR"):
            with self.subTest(response=response):
                popup = self.make_popup(response, "flush-error")
                self.click(popup, "ATA")
                self.assertEqual(self.results[0].ok, response == "OK")
                self.assertEqual(self.history_only(), response == "OK")
                self.assertFalse(serial_sender.serial_sms_recovery_error(self.modem))

    def test_success_callback_cannot_confirm_replacement_call_or_connection(self):
        for change in ("call", "serial", "generation"):
            with self.subTest(change=change):
                def replace_current(namespace):
                    if change == "call":
                        finish_incoming_call_session_namespace_runtime(namespace)
                        self.fixture.start_call()
                    elif change == "serial":
                        namespace["serial_obj"] = object()
                    else:
                        namespace["serial_connection_generation"] += 1
                popup = self.make_popup(before_result=replace_current)
                self.click(popup, "ATA")
                self.assertTrue(self.results[0].ok)
                self.assertFalse(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)
                self.assertFalse(self.ui)

    def test_late_terminal_cannot_reopen_isolated_serial_for_new_command(self):
        popup = self.make_popup(None)
        self.click(popup, "ATH")
        before = list(self.modem.writes)
        self.coordinator.observe_line("OK", connection=self.modem)
        result = serial_sender.write_serial_command_confirmed_locked(
            self.namespace["serial_lock"], lambda: self.modem, "ATI",
            response_coordinator=self.coordinator, response_timeout=0.02)
        self.assertFalse(result.ok)
        self.assertEqual(self.modem.writes, before)
        self.assertFalse(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)

    def test_timeout_isolation_finishes_before_next_transaction_can_write(self):
        popup = self.make_popup(None, response_timeout=0.15)
        popup.answer(lambda: None, lambda: None)
        self.assertTrue(self.wrote.wait(1))
        next_results, next_done = [], threading.Event()

        def next_command():
            next_results.append(serial_sender.write_serial_command_result(self.modem, "ATI"))
            next_done.set()

        worker = threading.Thread(target=next_command)
        worker.start()
        try:
            self.assertFalse(next_done.wait(0.02))
            for original in self.workers:
                original.join(2)
                self.assertFalse(original.is_alive())
            worker.join(2)
            self.assertFalse(worker.is_alive())
            self.assertTrue(serial_sender.serial_sms_recovery_error(self.modem))
            self.assertFalse(next_results[0].ok)
            self.assertEqual(self.modem.writes, [b"ATA\r\n"])
        finally:
            self.coordinator.cancel_active("test cleanup")
            worker.join(2)

    def test_serial_replacement_waits_until_success_confirmation_commits(self):
        popup = self.make_popup()
        tracker = self.namespace["INCOMING_CALL_SESSION"]
        original_confirm = tracker.confirm_handled
        confirming, release, replacing, replaced = (threading.Event() for _ in range(4))

        def confirm(token):
            confirming.set()
            self.assertTrue(release.wait(1))
            return original_confirm(token)

        def replace_serial():
            replacing.set()
            with self.namespace["serial_lock"]:
                self.namespace["serial_obj"] = object()
                self.namespace["serial_connection_generation"] += 1
                replaced.set()

        replacement = threading.Thread(target=replace_serial)
        with patch.object(tracker, "confirm_handled", side_effect=confirm):
            popup.answer(lambda: None, lambda: None)
            try:
                self.assertTrue(confirming.wait(1))
                replacement.start()
                self.assertTrue(replacing.wait(1))
                self.assertFalse(replaced.wait(0.02))
            finally:
                release.set()
                for worker in self.workers:
                    worker.join(2)
                    self.assertFalse(worker.is_alive())
                if replacement.ident is not None:
                    replacement.join(2)
        self.assertTrue(replaced.is_set())
        self.assertTrue(tracker.snapshot().confirmed_handled)
