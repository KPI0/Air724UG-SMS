import unittest
from unittest.mock import patch

import sms_app.serial_namespace_bindings as bindings


class SerialNamespaceBindingsTests(unittest.TestCase):
    def make_namespace(self):
        calls = []

        class FakeSerialModule:
            SerialException = RuntimeError

        return {
            "calls": calls,
            "_rebind_hint_notice": "rebind_notice",
            "_serial_error_notice": "serial_notice",
            "system_ui": lambda *args, **kwargs: calls.append(("system_ui", args, kwargs)),
            "serial_lock": "lock",
            "serial_obj": "serial_obj",
            "serial": FakeSerialModule,
            "write_serial_command_result": lambda *args: calls.append(("write", args)),
            "_parse_cloud_sms_callback_head": "parser",
            "LOG_PREFIX": "system",
            "APP_INSTANCE_NUMBER": 3,
        }

    def test_install_registers_expected_names_and_defaults(self):
        namespace = self.make_namespace()

        result = bindings.install_serial_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        self.assertIs(namespace["choose_manual_rebind_candidate"], bindings.choose_manual_rebind_candidate)
        self.assertIs(namespace["manual_rebind_hint"], bindings.manual_rebind_hint)
        for name in (
            "scan_com_ports_all",
            "find_luat_best_port",
            "_push_serial_debug",
            "try_rebind_manual_port",
            "rebind_hint_ui",
            "serial_error_ui",
            "set_call_popup",
            "close_call_popup",
            "close_missed_call_popup",
            "show_call_popup",
            "set_missed_call_popup",
            "show_missed_call_popup",
            "start_incoming_call_session",
            "mark_incoming_call_handled",
            "finish_incoming_call_session",
            "reset_incoming_call_session",
            "resolve_serial_target_port",
            "open_and_initialize_serial",
            "schedule_delayed_connected_log",
            "read_serial_line_safely",
            "try_manual_rebind_after_error",
            "send_call_hangup_command",
            "get_serial_call_state",
            "set_serial_call_state",
            "set_serial_log_prefix",
            "read_serial",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_install_preserves_existing_default_dependencies(self):
        namespace = self.make_namespace()
        choose = lambda: "choose"
        hint = lambda: "hint"
        namespace["choose_manual_rebind_candidate"] = choose
        namespace["manual_rebind_hint"] = hint

        bindings.install_serial_namespace_bindings(namespace)

        self.assertIs(namespace["choose_manual_rebind_candidate"], choose)
        self.assertIs(namespace["manual_rebind_hint"], hint)

    def test_scan_rebind_popup_and_startup_bindings_forward_namespace(self):
        namespace = self.make_namespace()
        bindings.install_serial_namespace_bindings(namespace)

        with patch.object(bindings, "scan_com_ports_all_namespace_runtime", return_value=["COM1"]) as scan, \
                patch.object(bindings, "find_luat_best_port_namespace_runtime", return_value=("COM1", "LUAT")) as find, \
                patch.object(bindings, "push_serial_debug_namespace_runtime", return_value="debugged") as debug, \
                patch.object(bindings, "try_rebind_manual_port_namespace_runtime", return_value=True) as rebind, \
                patch.object(bindings, "set_call_popup_namespace_runtime", return_value="set") as set_popup, \
                patch.object(bindings, "close_call_popup_namespace_runtime", return_value="closed") as close_popup, \
                patch.object(bindings, "close_missed_call_popup_namespace_runtime", return_value="missed-closed") as close_missed_popup, \
                patch.object(bindings, "show_call_popup_namespace_runtime", return_value="shown") as show_popup, \
                patch.object(bindings, "resolve_serial_target_port_namespace_runtime", return_value="COM1") as resolve, \
                patch.object(bindings, "open_and_initialize_serial_namespace_runtime", return_value="serial") as open_serial, \
                patch.object(bindings, "schedule_delayed_connected_log_namespace_runtime", return_value="scheduled") as connected_log:
            self.assertEqual(namespace["scan_com_ports_all"](), ["COM1"])
            self.assertEqual(namespace["find_luat_best_port"](), ("COM1", "LUAT"))
            self.assertEqual(namespace["_push_serial_debug"]("raw"), "debugged")
            self.assertTrue(namespace["try_rebind_manual_port"]("changed"))
            self.assertEqual(namespace["set_call_popup"]("win"), "set")
            self.assertEqual(namespace["close_call_popup"](), "closed")
            self.assertEqual(namespace["close_missed_call_popup"](), "missed-closed")
            self.assertEqual(namespace["show_call_popup"]("10086"), "shown")
            self.assertEqual(namespace["resolve_serial_target_port"](), "COM1")
            self.assertEqual(namespace["open_and_initialize_serial"]("COM1"), "serial")
            self.assertEqual(namespace["schedule_delayed_connected_log"]("COM1", 115200, delay=4), "scheduled")

        scan.assert_called_once_with(namespace)
        find.assert_called_once_with(namespace)
        debug.assert_called_once_with(namespace, "raw")
        rebind.assert_called_once_with(namespace, "changed")
        set_popup.assert_called_once_with(namespace, "win")
        close_popup.assert_called_once_with(namespace)
        close_missed_popup.assert_called_once_with(namespace)
        show_popup.assert_called_once_with(namespace, "10086")
        resolve.assert_called_once_with(namespace)
        open_serial.assert_called_once_with(namespace, "COM1")
        connected_log.assert_called_once_with(namespace, "COM1", 115200, delay=4)

    def test_notice_bindings_forward_repeat_state(self):
        namespace = self.make_namespace()
        bindings.install_serial_namespace_bindings(namespace)

        with patch.object(bindings, "emit_repeat_notice", return_value="emitted") as emit:
            self.assertEqual(namespace["rebind_hint_ui"]("hint"), "emitted")
            self.assertEqual(namespace["serial_error_ui"]("error", repeat_key="port"), "emitted")

        self.assertEqual(emit.call_args_list[0].args, ("rebind_notice", "hint", namespace["system_ui"]))
        self.assertEqual(emit.call_args_list[1].args, ("serial_notice", "error", namespace["system_ui"]))
        self.assertEqual(emit.call_args_list[1].kwargs, {"repeat_key": "port"})

    def test_io_call_state_and_reader_bindings_forward_runtime_context(self):
        namespace = self.make_namespace()
        bindings.install_serial_namespace_bindings(namespace)

        with patch.object(bindings, "read_serial_line_safely_runtime", return_value=b"line") as read_line, \
                patch.object(bindings, "send_call_hangup_runtime", return_value="sent") as hangup, \
                patch.object(bindings, "get_serial_call_state_namespace_runtime", return_value=(1.0, "10086")) as get_state, \
                patch.object(bindings, "set_serial_call_state_namespace_runtime", return_value="state_set") as set_state, \
                patch.object(bindings, "try_manual_rebind_after_error_namespace_runtime", return_value=True) as after_error, \
                patch.object(bindings, "run_serial_reader_namespace_runtime", return_value="ran") as reader:
            self.assertEqual(namespace["read_serial_line_safely"](), b"line")
            self.assertEqual(namespace["send_call_hangup_command"](), "sent")
            self.assertEqual(namespace["get_serial_call_state"](), (1.0, "10086"))
            self.assertEqual(namespace["set_serial_call_state"](2.0, "10010"), "state_set")
            self.assertTrue(namespace["try_manual_rebind_after_error"](RuntimeError("gone")))
            namespace["set_serial_log_prefix"]("COM1")
            self.assertEqual(namespace["read_serial"](), "ran")

        self.assertEqual(read_line.call_args.args[0], "lock")
        self.assertEqual(read_line.call_args.args[1](), "serial_obj")
        self.assertIs(read_line.call_args.args[2], RuntimeError)
        self.assertEqual(hangup.call_args.args[0], "lock")
        self.assertEqual(hangup.call_args.args[1](), "serial_obj")
        self.assertIs(hangup.call_args.args[2], namespace["write_serial_command_result"])
        get_state.assert_called_once_with(namespace)
        set_state.assert_called_once_with(namespace, 2.0, "10010")
        self.assertIn("Manual", after_error.call_args.kwargs["hint_message"])
        self.assertEqual(namespace["LOG_PREFIX"], "COM1")
        self.assertEqual(reader.call_args.args, (namespace,))
        self.assertEqual(reader.call_args.kwargs["parse_callback_head"], "parser")
        self.assertIs(reader.call_args.kwargs["apply_disconnect_effects"], bindings.apply_serial_disconnect_effects)

    def test_system_log_prefix_uses_instance_number_after_disconnect(self):
        namespace = self.make_namespace()
        bindings.install_serial_namespace_bindings(namespace)

        namespace["set_serial_log_prefix"]("system")

        self.assertEqual(namespace["LOG_PREFIX"], "system_3")


if __name__ == "__main__":
    unittest.main()
