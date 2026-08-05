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
            ("save",),
            ("state", "Manual", "COM7", 9600),
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

    def test_apply_serial_setting_runtime_rejects_non_positive_baud(self):
        for baud in (0, -1):
            with self.subTest(baud=baud):
                calls = []
                config = configparser.ConfigParser()
                config["serial"] = {"mode": "Manual", "port": "COM3", "baud": "9600"}
                before = dict(config.items("serial", raw=True))

                result = apply_serial_setting_runtime(
                    "Auto",
                    "",
                    baud,
                    config=config,
                    save_config=lambda: calls.append("save"),
                    set_serial_state=lambda *_: calls.append("state"),
                    set_status=lambda *_: calls.append("status"),
                    safe_close_serial=lambda: calls.append("close"),
                    wake_serial=lambda: calls.append("wake"),
                    system_ui=lambda message: calls.append(message),
                )

                self.assertFalse(result)
                self.assertEqual(dict(config.items("serial", raw=True)), before)
                self.assertEqual(calls, ["❌ 串口设置无效：波特率必须是大于 0 的整数"])

    def test_apply_serial_setting_runtime_reports_false_save_result(self):
        calls = []
        config = configparser.ConfigParser()
        config["serial"] = {"mode": "Manual", "port": "COM3", "baud": "9600", "extra": "keep"}
        before = dict(config.items("serial", raw=True))

        result = apply_serial_setting_runtime(
            "Auto",
            "",
            115200,
            config=config,
            save_config=lambda: calls.append("save") or False,
            set_serial_state=lambda *_: calls.append("state"),
            set_status=lambda *_: calls.append("status"),
            safe_close_serial=lambda: calls.append("close"),
            wake_serial=lambda: calls.append("wake"),
            system_ui=lambda message: calls.append(message),
        )

        self.assertFalse(result)
        self.assertEqual(dict(config.items("serial", raw=True)), before)
        self.assertIn("串口设置保存失败", calls[-1])
        self.assertNotIn("state", calls)
        self.assertNotIn("status", calls)
        self.assertNotIn("close", calls)
        self.assertNotIn("wake", calls)

    def test_apply_serial_setting_runtime_rolls_back_when_save_raises(self):
        calls = []
        config = configparser.ConfigParser()
        config["serial"] = {"mode": "Manual", "port": "COM3", "baud": "9600"}
        before = dict(config.items("serial", raw=True))

        result = apply_serial_setting_runtime(
            "Auto",
            "",
            115200,
            config=config,
            save_config=lambda: (_ for _ in ()).throw(RuntimeError("disk")),
            set_serial_state=lambda *_: calls.append("state"),
            set_status=lambda *_: calls.append("status"),
            safe_close_serial=lambda: calls.append("close"),
            wake_serial=lambda: calls.append("wake"),
            system_ui=lambda message: calls.append(message),
        )

        self.assertFalse(result)
        self.assertEqual(dict(config.items("serial", raw=True)), before)
        self.assertEqual(calls, ["❌ 串口设置保存失败，已保留原设置"])

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

    def test_open_serial_setting_runtime_propagates_save_failure(self):
        results = []

        def open_dialog(_parent, _mode, _port, _baud, _scan, apply, _center):
            results.append(apply("Auto", "", 115200))

        open_serial_setting_runtime(
            "root",
            current_mode="Manual",
            current_port="COM3",
            current_baud=9600,
            scan_ports=lambda: ["COM3"],
            center_window="center",
            config=configparser.ConfigParser(),
            save_config=lambda: False,
            set_serial_state=lambda *_: self.fail("state changed after failed save"),
            set_status=lambda *_: self.fail("status changed after failed save"),
            safe_close_serial=lambda: self.fail("serial closed after failed save"),
            wake_serial=lambda: self.fail("serial woke after failed save"),
            system_ui=lambda *_: None,
            open_dialog=open_dialog,
        )

        self.assertEqual(results, [False])


if __name__ == "__main__":
    unittest.main()
