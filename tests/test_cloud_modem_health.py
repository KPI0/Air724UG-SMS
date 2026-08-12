import unittest

from sms_core.cloud_modem_health import CloudModemHealthState


class CloudModemHealthStateTests(unittest.TestCase):
    def test_three_consecutive_response_timeouts_request_one_reconnect(self):
        health = CloudModemHealthState(reconnect_threshold=3)

        first = health.record(False, "等待 Modem 指令响应超时")
        second = health.record(False, "等待 Modem 指令响应超时")
        third = health.record(False, "等待 Modem 指令响应超时")
        fourth = health.record(False, "等待 Modem 指令响应超时")

        self.assertFalse(first["modem_unresponsive"])
        self.assertFalse(second["request_reconnect"])
        self.assertTrue(third["modem_unresponsive"])
        self.assertTrue(third["request_reconnect"])
        self.assertFalse(fourth["request_reconnect"])

    def test_success_clears_unresponsive_state_and_allows_future_reconnect(self):
        health = CloudModemHealthState(reconnect_threshold=2)
        health.record(False, "等待 Modem 指令响应超时")
        self.assertTrue(
            health.record(False, "等待 Modem 指令响应超时")["request_reconnect"]
        )

        recovered = health.record(True, "")

        self.assertFalse(recovered["modem_unresponsive"])
        self.assertEqual(recovered["consecutive_at_timeouts"], 0)
        health.record(False, "等待 Modem 指令响应超时")
        self.assertTrue(
            health.record(False, "等待 Modem 指令响应超时")["request_reconnect"]
        )

    def test_modem_error_does_not_count_as_response_timeout(self):
        health = CloudModemHealthState(reconnect_threshold=2)
        health.record(False, "等待 Modem 指令响应超时")

        state = health.record(False, "+CME ERROR: 3")

        self.assertEqual(state["consecutive_at_timeouts"], 0)
        self.assertFalse(state["modem_unresponsive"])


if __name__ == "__main__":
    unittest.main()
