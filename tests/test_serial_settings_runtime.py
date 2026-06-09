import configparser
import unittest

from sms_ui.serial_settings_runtime import (
    apply_serial_setting_runtime,
    open_serial_setting_runtime,
    serial_settings_status,
)


class SerialSettingsRuntimeTests(unittest.TestCase):
    def test_serial_settings_status_formats_auto_port(self):
        self.assertEqual(
            serial_settings_status("Auto", "", 115200),
            "⚙️ 串口设置已更新：mode=Auto port=(Auto) baud=115200",
        )

    def test_apply_serial_setting_runtime_updates_config_and_reconnects(self):
        config = configparser.ConfigParser()
        calls = []

        apply_serial_setting_runtime(
            "Manual",
            "COM7",
            9600,
            config=config,
            save_config=lambda: calls.append(("save",)),
            set_serial_state=lambda mode, port, baud: calls.append(("state", mode, port, baud)),
            set_status=lambda text, color: calls.append(("status", text, color)),
            safe_close_serial=lambda: calls.append(("close",)),
            wake_serial=lambda: calls.append(("wake",)),
            system_ui=lambda message: calls.append(("ui", message)),
        )

        self.assertEqual(config.get("serial", "mode"), "Manual")
        self.assertEqual(config.get("serial", "port"), "COM7")
        self.assertEqual(config.get("serial", "baud"), "9600")
        self.assertEqual(calls[:4], [
            ("state", "Manual", "COM7", 9600),
            ("save",),
            ("status", "🟡 应用中，重连…", "orange"),
            ("close",),
        ])
        self.assertIn(("wake",), calls)
        self.assertIn(("ui", "⚙️ 串口设置已更新：mode=Manual port=COM7 baud=9600"), calls)

    def test_apply_serial_setting_runtime_ignores_wake_errors(self):
        calls = []

        apply_serial_setting_runtime(
            "Auto",
            "",
            115200,
            config=configparser.ConfigParser(),
            save_config=lambda: None,
            set_serial_state=lambda *_: None,
            set_status=lambda *_: None,
            safe_close_serial=lambda: calls.append("closed"),
            wake_serial=lambda: (_ for _ in ()).throw(RuntimeError("boom")),
            system_ui=lambda *_: calls.append("ui"),
        )

        self.assertEqual(calls, ["closed", "ui"])

    def test_open_serial_setting_runtime_wires_dialog_apply(self):
        config = configparser.ConfigParser()
        opened = {}
        calls = []

        def open_dialog(parent, mode, port, baud, scan_ports, apply, center_window):
            opened.update(
                parent=parent,
                mode=mode,
                port=port,
                baud=baud,
                ports=scan_ports(),
                center_window=center_window,
            )
            apply("Auto", "", 115200)

        open_serial_setting_runtime(
            "root",
            current_mode="Manual",
            current_port="COM3",
            current_baud=9600,
            scan_ports=lambda: ["COM3"],
            center_window="center",
            config=config,
            save_config=lambda: calls.append("save"),
            set_serial_state=lambda mode, port, baud: calls.append((mode, port, baud)),
            set_status=lambda *_: calls.append("status"),
            safe_close_serial=lambda: calls.append("close"),
            wake_serial=lambda: calls.append("wake"),
            system_ui=lambda *_: calls.append("ui"),
            open_dialog=open_dialog,
        )

        self.assertEqual(opened, {
            "parent": "root",
            "mode": "Manual",
            "port": "COM3",
            "baud": 9600,
            "ports": ["COM3"],
            "center_window": "center",
        })
        self.assertEqual(config.get("serial", "mode"), "Auto")
        self.assertIn(("Auto", "", 115200), calls)


if __name__ == "__main__":
    unittest.main()
