import unittest
from types import SimpleNamespace

from sms_ui.device_reset_runtime import RESET_COMMAND, send_reset_command_runtime


class DeviceResetRuntimeTests(unittest.TestCase):
    def test_send_reset_command_runtime_skips_when_cancelled(self):
        calls = []

        result = send_reset_command_runtime(
            confirm_reset=lambda: False,
            send_command_async=lambda *args, **kwargs: calls.append(("send", args, kwargs)),
            serial_lock=object(),
            get_serial=lambda: object(),
            ui_post=lambda callback: callback(),
            system_ui=lambda *args: calls.append(("ui", args)),
            show_warning=lambda *args: calls.append(("warning", args)),
        )

        self.assertEqual(result, "cancelled")
        self.assertEqual(calls, [])

    def test_send_reset_command_runtime_submits_reset_command(self):
        calls = []
        serial_lock = object()
        serial_obj = object()

        result = send_reset_command_runtime(
            confirm_reset=lambda: True,
            send_command_async=lambda *args, **kwargs: calls.append(("send", args, kwargs)),
            serial_lock=serial_lock,
            get_serial=lambda: serial_obj,
            ui_post=lambda callback: callback(),
            system_ui=lambda *args: calls.append(("ui", args)),
            show_warning=lambda *args: calls.append(("warning", args)),
        )

        self.assertEqual(result, "submitted")
        self.assertEqual(calls[0][0], "send")
        self.assertEqual(calls[0][1], (serial_lock, calls[0][1][1], RESET_COMMAND))
        self.assertIs(calls[0][1][1](), serial_obj)
        self.assertIn("on_result", calls[0][2])

    def test_send_reset_command_runtime_logs_success(self):
        calls = []
        result_callback = None

        def send_command_async(*_args, **kwargs):
            nonlocal result_callback
            result_callback = kwargs["on_result"]

        send_reset_command_runtime(
            confirm_reset=lambda: True,
            send_command_async=send_command_async,
            serial_lock=object(),
            get_serial=lambda: object(),
            ui_post=lambda callback: callback(),
            system_ui=lambda *args: calls.append(("ui", args)),
            show_warning=lambda *args: calls.append(("warning", args)),
        )

        result_callback(SimpleNamespace(ok=True, error=None))

        self.assertEqual(calls, [("ui", (f"🔄 已发送重启指令：{RESET_COMMAND}", "normal"))])

    def test_send_reset_command_runtime_warns_on_failure(self):
        calls = []
        result_callback = None

        def send_command_async(*_args, **kwargs):
            nonlocal result_callback
            result_callback = kwargs["on_result"]

        send_reset_command_runtime(
            confirm_reset=lambda: True,
            send_command_async=send_command_async,
            serial_lock=object(),
            get_serial=lambda: object(),
            ui_post=lambda callback: calls.append(("post", callback)) or callback(),
            system_ui=lambda *args: calls.append(("ui", args)),
            show_warning=lambda *args: calls.append(("warning", args)),
        )

        result_callback(SimpleNamespace(ok=False, error="closed"))

        self.assertEqual(calls[0][0], "post")
        self.assertEqual(calls[1], ("warning", ("提示", "串口当前未连接或发送失败：closed")))
        self.assertEqual(calls[2], ("ui", ("❌ 发送重启指令失败：closed", "normal")))


if __name__ == "__main__":
    unittest.main()
