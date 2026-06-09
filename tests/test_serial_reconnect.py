import configparser
import unittest

from sms_core.serial_reconnect import (
    apply_serial_disconnect_effects,
    is_serial_open_denied,
    is_serial_port_gone_error,
    manual_rebind_runtime,
    serial_disconnect_status,
    serial_open_denied_repeat_key,
)
from sms_core.serial_ports import ManualRebindCandidate


class SerialReconnectTests(unittest.TestCase):
    def test_serial_open_denied_detection_and_repeat_key(self):
        self.assertTrue(is_serial_open_denied("Access is denied."))
        self.assertTrue(is_serial_open_denied("拒绝访问"))
        self.assertFalse(is_serial_open_denied("file not found"))
        self.assertEqual(serial_open_denied_repeat_key(" com5 "), "serial-open-denied:COM5")
        self.assertEqual(serial_open_denied_repeat_key(None), "serial-open-denied:")

    def test_serial_disconnect_status(self):
        self.assertEqual(serial_disconnect_status("COM5"), "🔴 断开/失败：COM5（自动重连中…）")

    def test_apply_serial_disconnect_effects_resets_ui_and_closes_port(self):
        calls = []

        apply_serial_disconnect_effects(
            PermissionError("Access is denied."),
            target_port="COM7",
            current_port="COM5",
            serial_error_ui=lambda text, repeat_key="": calls.append(("error", text, repeat_key)),
            set_status=lambda text, color: calls.append(("status", text, color)),
            set_temperature=lambda value: calls.append(("temp", value)),
            set_signal=lambda value: calls.append(("signal", value)),
            close_serial=lambda: calls.append(("close",)),
        )

        self.assertEqual(calls[0][0], "error")
        self.assertIn("Access is denied.", calls[0][1])
        self.assertEqual(calls[0][2], "serial-open-denied:COM7")
        self.assertIn(("status", "🔴 断开/失败：COM5（自动重连中…）", "red"), calls)
        self.assertIn(("temp", "--"), calls)
        self.assertIn(("signal", "--"), calls)
        self.assertIn(("close",), calls)

    def test_serial_port_gone_detection(self):
        self.assertTrue(is_serial_port_gone_error("could not open port COM5"))
        self.assertTrue(is_serial_port_gone_error("系统找不到指定的文件"))
        self.assertFalse(is_serial_port_gone_error("temporary parse error"))

    def test_manual_rebind_runtime_updates_config_and_wakes_serial(self):
        config = configparser.ConfigParser()
        calls = []

        ok = manual_rebind_runtime(
            mode="Manual",
            current_port="COM5",
            baud=115200,
            reason="changed",
            find_luat_best_port=lambda: ("COM7", "LUAT MODEM"),
            list_ports=lambda: ["unused"],
            choose_candidate=lambda dev, desc, ports, current_port="": ManualRebindCandidate(dev, desc),
            config=config,
            save_config=lambda: calls.append(("save",)),
            set_port=lambda port: calls.append(("port", port)),
            system_ui=lambda text, tag: calls.append(("ui", text, tag)),
            set_status=lambda text, color: calls.append(("status", text, color)),
            wake_serial=lambda: calls.append(("wake",)),
            reset_rebind_hint=lambda: calls.append(("reset",)),
            hint_formatter=lambda old, new, desc, reason: f"{old}->{new}:{desc}:{reason}",
        )

        self.assertTrue(ok)
        self.assertEqual(config.get("serial", "mode"), "Manual")
        self.assertEqual(config.get("serial", "port"), "COM7")
        self.assertEqual(config.get("serial", "baud"), "115200")
        self.assertIn(("port", "COM7"), calls)
        self.assertIn(("save",), calls)
        self.assertIn(("ui", "COM5->COM7:LUAT MODEM:changed", "normal"), calls)
        self.assertIn(("wake",), calls)
        self.assertIn(("reset",), calls)

    def test_manual_rebind_runtime_ignores_non_manual_mode(self):
        calls = []

        ok = manual_rebind_runtime(
            mode="Auto",
            current_port="COM5",
            baud=115200,
            reason="changed",
            find_luat_best_port=lambda: self.fail("should not scan"),
            list_ports=lambda: [],
            choose_candidate=lambda *_args, **_kwargs: ManualRebindCandidate(),
            config=configparser.ConfigParser(),
            save_config=lambda: calls.append(("save",)),
            set_port=lambda port: calls.append(("port", port)),
            system_ui=lambda text, tag: calls.append(("ui", text, tag)),
            set_status=lambda text, color: calls.append(("status", text, color)),
            wake_serial=lambda: calls.append(("wake",)),
            reset_rebind_hint=lambda: calls.append(("reset",)),
            hint_formatter=lambda *_: "",
        )

        self.assertFalse(ok)
        self.assertEqual(calls, [])

    def test_manual_rebind_runtime_falls_back_to_empty_ports_on_listing_error(self):
        calls = []

        ok = manual_rebind_runtime(
            mode="Manual",
            current_port="COM5",
            baud=9600,
            reason="changed",
            find_luat_best_port=lambda: (None, None),
            list_ports=lambda: (_ for _ in ()).throw(RuntimeError("list failed")),
            choose_candidate=lambda dev, desc, ports, current_port="": calls.append(("candidate", ports)) or ManualRebindCandidate("COM9", "fallback"),
            config=configparser.ConfigParser(),
            save_config=lambda: None,
            set_port=lambda port: calls.append(("port", port)),
            system_ui=lambda *_: None,
            set_status=lambda *_: None,
            wake_serial=lambda: None,
            reset_rebind_hint=lambda: None,
            hint_formatter=lambda *_: "",
        )

        self.assertTrue(ok)
        self.assertIn(("candidate", []), calls)


if __name__ == "__main__":
    unittest.main()
