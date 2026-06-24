import unittest

from sms_core.serial_line_effects import apply_serial_line_effects, push_serial_debug_insights


class SerialLineEffectsTests(unittest.TestCase):
    def test_apply_serial_line_effects_dispatches_raw_line_and_metrics(self):
        calls = []
        line = "[I]-[ril.proatc] +RFTEMPERATURE: 28.84"

        apply_serial_line_effects(
            line,
            lambda value: calls.append(("debug", value)),
            lambda value: calls.append(("cloud", value)),
            lambda value: calls.append(("imei", value)),
            lambda value: calls.append(("temp", value)),
            lambda value: calls.append(("signal", value)),
        )

        self.assertEqual(calls[:3], [
            ("debug", line),
            ("cloud", line),
            ("imei", line),
        ])
        self.assertIn(("temp", "28.84"), calls)
        self.assertNotIn(("signal", "28.84"), calls)

    def test_apply_serial_line_effects_updates_signal_when_present(self):
        calls = []
        line = "[I]-[ril.proatc] +CESQ: 99,99,255,255,26,49"

        apply_serial_line_effects(
            line,
            lambda value: calls.append(("debug", value)),
            lambda value: calls.append(("cloud", value)),
            lambda value: calls.append(("imei", value)),
            lambda value: calls.append(("temp", value)),
            lambda value: calls.append(("signal", value)),
        )

        self.assertIn(("signal", "49"), calls)

    def test_apply_serial_line_effects_updates_local_number_from_cnum(self):
        calls = []
        line = '[I]-[ril.proatc] +CNUM: "","+8613812345678",145'

        apply_serial_line_effects(
            line,
            lambda value: calls.append(("debug", value)),
            lambda value: calls.append(("cloud", value)),
            lambda value: calls.append(("imei", value)),
            lambda value: calls.append(("temp", value)),
            lambda value: calls.append(("signal", value)),
            lambda value: calls.append(("local", value)),
        )

        self.assertIn(("local", "+8613812345678"), calls)

    def test_push_serial_debug_insights_pushes_parser_messages(self):
        calls = []

        push_serial_debug_insights(
            '[I]-[ril.proatc] +COPS: 0,2,"46011",7',
            lambda value: calls.append(value),
        )

        self.assertTrue(calls)
        self.assertTrue(any("46011" in item for item in calls))


if __name__ == "__main__":
    unittest.main()
