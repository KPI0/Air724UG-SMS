import unittest

from sms_core.serial_reconnect_app_runtime import (
    try_rebind_manual_port_runtime,
)


class SerialReconnectAppRuntimeTests(unittest.TestCase):
    def test_try_rebind_manual_port_runtime_forwards_app_dependencies(self):
        calls = []

        def runtime(**kwargs):
            calls.append(kwargs)
            kwargs["set_port"]("COM7")
            kwargs["reset_rebind_hint"]()
            return True

        result = try_rebind_manual_port_runtime(
            "changed",
            mode="Manual",
            current_port="COM5",
            baud=115200,
            find_luat_best_port=lambda: ("COM7", "LUAT"),
            list_ports=lambda: [],
            choose_candidate=lambda *_args, **_kwargs: None,
            config="config",
            save_config=lambda: calls.append(("save",)),
            set_port=lambda port: calls.append(("port", port)),
            system_ui=lambda *_args: None,
            set_status=lambda *_args: None,
            wake_serial=lambda: None,
            reset_rebind_hint=lambda: calls.append(("reset",)),
            hint_formatter=lambda *_args: "",
            runtime=runtime,
        )

        self.assertTrue(result)
        kwargs = calls[0]
        self.assertEqual(kwargs["mode"], "Manual")
        self.assertEqual(kwargs["current_port"], "COM5")
        self.assertEqual(kwargs["baud"], 115200)
        self.assertEqual(kwargs["reason"], "changed")
        self.assertEqual(kwargs["config"], "config")
        self.assertIn(("port", "COM7"), calls)
        self.assertIn(("reset",), calls)


if __name__ == "__main__":
    unittest.main()
