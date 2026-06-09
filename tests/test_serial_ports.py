import unittest

from sms_core.serial_ports import (
    SerialPortInfo,
    choose_manual_rebind_candidate,
    choose_luat_modem_port,
    is_luat_modem_candidate,
    manual_rebind_hint,
    unlocked_ports,
)


class FakePort:
    def __init__(self, device, description="", hwid=""):
        self.device = device
        self.description = description
        self.hwid = hwid


class SerialPortSelectionTests(unittest.TestCase):
    def test_is_luat_modem_candidate_filters_non_business_ports(self):
        self.assertTrue(is_luat_modem_candidate(SerialPortInfo("COM1", "LUAT USB MODEM", "")))
        self.assertFalse(is_luat_modem_candidate(SerialPortInfo("COM2", "LUAT USB DIAG", "")))
        self.assertFalse(is_luat_modem_candidate(SerialPortInfo("COM3", "LUAT USB Device 1 AT", "")))
        self.assertFalse(is_luat_modem_candidate(SerialPortInfo("COM4", "USB MODEM", "")))

    def test_choose_luat_modem_port_prefers_remembered_port(self):
        ports = [
            FakePort("COM1", "LUAT USB MODEM USB DEVICE 0"),
            FakePort("COM9", "LUAT USB Device 9", "LUAT"),
        ]

        self.assertEqual(
            choose_luat_modem_port(ports, remembered_port="COM9"),
            ("COM9", "LUAT USB Device 9"),
        )

    def test_choose_luat_modem_port_skips_locked_ports(self):
        ports = [
            FakePort("COM1", "LUAT USB MODEM USB DEVICE 0"),
            FakePort("COM2", "LUAT USB MODEM"),
        ]

        self.assertEqual(
            choose_luat_modem_port(ports, is_locked=lambda device: device == "COM1"),
            ("COM2", "LUAT USB MODEM"),
        )

    def test_unlocked_ports_filters_locked_devices(self):
        ports = [FakePort("COM1"), FakePort("COM2")]

        self.assertEqual(
            [port.device for port in unlocked_ports(ports, lambda device: device == "COM1")],
            ["COM2"],
        )

    def test_choose_manual_rebind_candidate_prefers_luat_result(self):
        candidate = choose_manual_rebind_candidate(
            "COM9",
            "LUAT USB MODEM",
            [FakePort("COM1", "fallback")],
            current_port="COM5",
        )

        self.assertTrue(candidate.found)
        self.assertEqual(candidate.device, "COM9")
        self.assertEqual(candidate.description, "LUAT USB MODEM")

    def test_choose_manual_rebind_candidate_uses_single_port_fallback(self):
        candidate = choose_manual_rebind_candidate(
            None,
            None,
            [FakePort("COM7", "single")],
            current_port="COM5",
        )

        self.assertTrue(candidate.found)
        self.assertEqual(candidate.device, "COM7")
        self.assertEqual(candidate.description, "single")

    def test_choose_manual_rebind_candidate_rejects_missing_or_same_port(self):
        self.assertFalse(choose_manual_rebind_candidate(None, None, [], current_port="COM5").found)
        self.assertFalse(
            choose_manual_rebind_candidate("COM5", "same", [FakePort("COM7")], current_port="COM5").found
        )

    def test_manual_rebind_hint_formats_optional_context(self):
        self.assertEqual(
            manual_rebind_hint("COM5", "COM7", "LUAT USB MODEM", "端口变化"),
            "🔁 手动模式端口失效，已自动重绑：COM5 -> COM7（LUAT USB MODEM）；原因：端口变化",
        )


if __name__ == "__main__":
    unittest.main()
