import unittest
from unittest.mock import patch

import sms_ui.serial_debug_namespace_bindings as bindings


class SerialDebugNamespaceBindingsTests(unittest.TestCase):
    def test_install_registers_expected_names(self):
        namespace = {}

        result = bindings.install_serial_debug_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        for name in (
            "get_serial_debug_state",
            "set_serial_debug_state",
            "open_serial_debug_window",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_bindings_forward_namespace_and_arguments(self):
        namespace = {}
        bindings.install_serial_debug_namespace_bindings(namespace)

        with patch.object(bindings, "get_serial_debug_state_namespace_runtime", return_value="state") as get_state, \
                patch.object(bindings, "set_serial_debug_state_namespace_runtime", return_value="set") as set_state, \
                patch.object(bindings, "open_serial_debug_window_namespace_runtime", return_value="opened") as open_window:
            self.assertEqual(namespace["get_serial_debug_state"]("debug_enabled"), "state")
            self.assertEqual(namespace["set_serial_debug_state"]("drop_count", 5), "set")
            self.assertEqual(namespace["open_serial_debug_window"](), "opened")

        get_state.assert_called_once_with(namespace, "debug_enabled")
        set_state.assert_called_once_with(namespace, "drop_count", 5)
        open_window.assert_called_once_with(namespace)


if __name__ == "__main__":
    unittest.main()
