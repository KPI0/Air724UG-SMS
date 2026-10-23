"""Permission changes must cancel commands before their actual serial writes."""
import threading
import unittest
from unittest.mock import patch

from sms_core import serial_sender
from sms_core.call_command_context import call_state_lock
from sms_core.cloud_command_context import (
    capture_cloud_serial_command_context, set_cloud_command_permissions,
)
from sms_core.cloud_message_namespace_runtime import send_cloud_serial_command_namespace_runtime
from tests.test_cloud_command_connection_context import Modem
from tests import test_sms_prompt_recovery as prompt_tests
from tests import test_settings_namespace_runtime as settings_tests
from tests import test_config_sync_namespace_runtime as sync_tests
from sms_ui.settings_namespace_runtime import open_security_settings_namespace_runtime
from sms_ui.config_sync_namespace_runtime import reload_shared_ui_config_namespace_runtime


class Gate:
    def __init__(self):
        self.reached = threading.Event()
        self.release = threading.Event()

    def __enter__(self):
        self.reached.set()
        if not self.release.wait(3):
            raise RuntimeError("test gate timed out")
        return self

    def __exit__(self, *_args):
        return False


def fixture():
    at = serial_sender.AtCommandResponseCoordinator()
    sms = serial_sender.SmsPduSendCoordinator()
    modem = Modem(at, sms)
    ns = {
        "serial_lock": threading.RLock(), "serial_obj": modem,
        "serial_connection_generation": 1, "_cloud_runtime_imei": lambda: "000000000000001",
        "CLOUD_DEVICE_SECRET": "synthetic-only", "CLOUD_WS_URL": "wss://example.invalid/ws/device",
        "CLOUD_CONTROL_ENABLED": True, "cloud_ws_conn": object(),
        "cloud_connected": True, "cloud_device_authorized": True,
        "CLOUD_SENSITIVE_COMMAND_PERMISSIONS": {"device_power": True, "sms": True, "phone_number": True},
        "SERIAL_COMMAND_RESPONSE_COORDINATOR": at, "SMS_SEND_COORDINATOR": sms,
        "_cloud_log": lambda *_args, **_kwargs: None,
    }
    return ns, modem


class CloudPermissionSubmissionTests(unittest.TestCase):
    def exercise(self, command, metadata, lock_name, category, reenable=False):
        ns, modem = fixture()
        gate = Gate()
        results = []
        errors = []

        def send():
            try:
                results.append(send_cloud_serial_command_namespace_runtime(ns, command, metadata))
            except Exception as exc:
                errors.append(exc)

        with patch.object(serial_sender, lock_name, gate):
            worker = threading.Thread(target=send)
            worker.start()
            try:
                self.assertTrue(gate.reached.wait(3))
                if reenable:
                    set_cloud_command_permissions(ns, {category: False})
                    set_cloud_command_permissions(ns, {category: True})
                else:
                    with call_state_lock(ns):
                        ns["CLOUD_SENSITIVE_COMMAND_PERMISSIONS"] = {category: False}
            finally:
                gate.release.set()
                worker.join(3)
        self.assertFalse(worker.is_alive())
        self.assertEqual(errors, [])
        self.assertEqual(modem.writes, [])
        self.assertFalse(results[0][0])

    def test_revocation_at_transaction_and_final_write_locks_rejects_all_writers(self):
        cases = (
            ("AT+REBOOT", {}, "device_power"),
            ("CLOUD:SEND_SMS", {"sms_phone": "10086", "sms_message": "synthetic"}, "sms"),
            ("CLOUD:SET_OWN_NUMBER", {"own_number": "+19990000001"}, "phone_number"),
        )
        for command, metadata, category in cases:
            for lock in ("DEFAULT_SERIAL_TRANSACTION_LOCK", "DEFAULT_SERIAL_WRITE_LOCK"):
                with self.subTest(command=command, lock=lock):
                    self.exercise(command, metadata, lock, category)

    def test_reenabled_permission_does_not_revive_commands_at_either_lock(self):
        for command, metadata, category in (
            ("AT+REBOOT", {}, "device_power"),
            ("CLOUD:SEND_SMS", {"sms_phone": "10086", "sms_message": "synthetic"}, "sms"),
            ("CLOUD:SET_OWN_NUMBER", {"own_number": "+19990000001"}, "phone_number"),
        ):
            for lock in ("DEFAULT_SERIAL_TRANSACTION_LOCK", "DEFAULT_SERIAL_WRITE_LOCK"):
                with self.subTest(command=command, lock=lock):
                    self.exercise(command, metadata, lock, category, reenable=True)

    def test_unrelated_permission_change_preserves_ordinary_at(self):
        ns, modem = fixture()
        context = capture_cloud_serial_command_context(ns, command="ATI")
        set_cloud_command_permissions(ns, False)
        result = send_cloud_serial_command_namespace_runtime(ns, "ATI", command_context=context)
        self.assertTrue(result[0])
        self.assertEqual(modem.writes, [b"ATI\r\n"])

    def test_number_transaction_rechecks_permission_between_each_write(self):
        ns, modem = fixture()
        original_write = modem.write

        def revoke_after_selection(payload):
            count = original_write(payload)
            set_cloud_command_permissions(ns, {"phone_number": False})
            return count

        modem.write = revoke_after_selection
        result = send_cloud_serial_command_namespace_runtime(
            ns, "CLOUD:SET_OWN_NUMBER", {"own_number": "+19990000001"},
        )
        self.assertFalse(result[0])
        self.assertEqual(len(modem.writes), 1)
        self.assertTrue(modem.writes[0].startswith(b'AT+CPBS="ON"'))

    def test_sms_permission_revocation_at_prompt_escapes_original_transaction(self):
        helper = prompt_tests.SmsPromptRecoveryTests()
        for reenable in (False, True):
            with self.subTest(reenable=reenable):
                ns, modem, at, sms = helper.fixture()

                def revoke():
                    set_cloud_command_permissions(ns, {"sms": False})
                    if reenable:
                        set_cloud_command_permissions(ns, {"sms": True})

                modem.on_prompt = revoke
                result = send_cloud_serial_command_namespace_runtime(
                    ns, "CLOUD:SEND_SMS", {"sms_phone": "10086", "sms_message": "synthetic"},
                )
                self.assertFalse(result[0])
                self.assertEqual(modem.writes[-1], b"\x1b")
                self.assertFalse(modem.pdu_written.is_set())
                self.assertEqual(modem.phase, "command")
                self.assertTrue(helper.next_command(ns, at).ok)
                self.assertIsNone(sms._active)

    def test_revocation_after_ctrl_z_never_sends_escape(self):
        helper = prompt_tests.SmsPromptRecoveryTests()
        ns, modem, at, sms = helper.fixture()
        original_write = modem.write

        def revoke_after_pdu(payload):
            count = original_write(payload)
            if payload.endswith(b"\x1a"):
                set_cloud_command_permissions(ns, {"sms": False})
            return count

        modem.write = revoke_after_pdu
        result = send_cloud_serial_command_namespace_runtime(
            ns, "CLOUD:SEND_SMS", {"sms_phone": "10086", "sms_message": "synthetic"},
        )
        self.assertTrue(result[0])
        self.assertTrue(modem.pdu_written.is_set())
        self.assertNotIn(b"\x1b", modem.writes)
        self.assertTrue(helper.next_command(ns, at).ok)
        self.assertIsNone(sms._active)

    def test_security_dialog_save_revokes_captured_request_permanently(self):
        ns, _modem = fixture()
        ns.update(settings_tests.SettingsNamespaceRuntimeTests().base_namespace())
        set_cloud_command_permissions(ns, {"sms": True})
        context = capture_cloud_serial_command_context(ns, command="CLOUD:SEND_SMS")
        forwarded = {}
        open_security_settings_namespace_runtime(
            ns, open_setting_runtime=lambda *_args, **kwargs: forwarded.update(kwargs),
        )
        forwarded["set_permissions"]({"sms": False})
        forwarded["set_permissions"]({"sms": True})
        self.assertIn("敏感权限", context.rejection_reason(ns))

    def test_external_config_reload_revokes_captured_request_permanently(self):
        helper = sync_tests.ConfigSyncNamespaceRuntimeTests()
        ns, _modem = fixture()
        ns.update(helper.make_namespace())
        set_cloud_command_permissions(ns, {"sms": True})
        context = capture_cloud_serial_command_context(ns, command="CLOUD:SEND_SMS")
        for value in ("0", "1"):
            snapshot = helper.disk_snapshot()
            snapshot["cloud_control"] = {
                "allow_sensitive_sms": value, "enabled": "1",
                "url": ns["CLOUD_WS_URL"], "device_secret": ns["CLOUD_DEVICE_SECRET"],
                "reconnect_interval": "5",
            }
            reload_shared_ui_config_namespace_runtime(ns, load_snapshot=lambda _path: snapshot)
        self.assertIn("敏感权限", context.rejection_reason(ns))

    def test_permission_setter_waits_for_actual_write_to_finish(self):
        ns, modem = fixture()
        write_started = threading.Event()
        release_write = threading.Event()
        setter_started = threading.Event()
        revoked = threading.Event()
        errors = []
        original_write = modem.write

        def blocking_write(payload):
            write_started.set()
            if not release_write.wait(3):
                raise RuntimeError("test write timed out")
            return original_write(payload)

        def send():
            try:
                send_cloud_serial_command_namespace_runtime(ns, "AT+REBOOT")
            except Exception as exc:
                errors.append(exc)

        def revoke():
            setter_started.set()
            set_cloud_command_permissions(ns, False)
            revoked.set()

        modem.write = blocking_write
        sender = threading.Thread(target=send)
        setter = threading.Thread(target=revoke)
        sender.start()
        try:
            self.assertTrue(write_started.wait(3))
            setter.start()
            self.assertTrue(setter_started.wait(3))
            self.assertFalse(revoked.wait(0.1))
        finally:
            release_write.set()
            sender.join(3)
            if setter.ident is not None:
                setter.join(3)
        self.assertFalse(sender.is_alive())
        self.assertFalse(setter.is_alive())
        self.assertTrue(revoked.is_set())
        self.assertEqual(errors, [])
        self.assertEqual(modem.writes, [b"AT+REBOOT\r\n"])
