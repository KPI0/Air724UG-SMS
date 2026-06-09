import unittest
from unittest.mock import patch

import sms_ui.app_infrastructure_namespace_bindings as bindings


class AppInfrastructureNamespaceBindingsTests(unittest.TestCase):
    def make_namespace(self):
        return {
            "_auto_connect_notice": "notice",
            "system_ui": lambda *args: None,
        }

    def test_install_registers_expected_names(self):
        namespace = self.make_namespace()

        result = bindings.install_app_infrastructure_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        for name in (
            "safe_save_config",
            "safe_close_serial",
            "auto_connect_ui",
            "lock_port_mutex",
            "unlock_port_mutex",
            "check_single_instance",
            "show_sms_popup",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_bindings_forward_namespace_and_arguments(self):
        namespace = self.make_namespace()
        bindings.install_app_infrastructure_namespace_bindings(namespace)

        with patch.object(bindings, "safe_save_config_namespace_runtime", return_value="saved") as save, \
                patch.object(bindings, "safe_close_serial_namespace_runtime", return_value="closed") as close, \
                patch.object(bindings, "lock_port_mutex_namespace_runtime", return_value="mutex") as lock, \
                patch.object(bindings, "unlock_port_mutex_namespace_runtime", return_value="unlocked") as unlock, \
                patch.object(bindings, "check_single_instance_namespace_runtime", return_value="single") as single, \
                patch.object(bindings, "show_sms_popup_namespace_runtime", return_value="shown") as popup, \
                patch.object(bindings, "emit_repeat_notice", return_value="emitted") as repeat:
            self.assertEqual(namespace["safe_save_config"](), "saved")
            self.assertEqual(namespace["safe_close_serial"](), "closed")
            self.assertEqual(namespace["lock_port_mutex"]("COM5"), "mutex")
            self.assertEqual(namespace["unlock_port_mutex"](), "unlocked")
            self.assertEqual(namespace["check_single_instance"](), "single")
            self.assertEqual(namespace["show_sms_popup"]("msg"), "shown")
            self.assertEqual(namespace["auto_connect_ui"]("hint"), "emitted")

        save.assert_called_once_with(namespace)
        close.assert_called_once_with(namespace)
        lock.assert_called_once_with(namespace, "COM5")
        unlock.assert_called_once_with(namespace)
        single.assert_called_once_with(namespace)
        popup.assert_called_once_with(namespace, "msg")
        repeat.assert_called_once_with("notice", "hint", namespace["system_ui"])


if __name__ == "__main__":
    unittest.main()
