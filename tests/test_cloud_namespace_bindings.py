import asyncio
import unittest
from unittest.mock import patch

import sms_app.cloud_namespace_bindings as bindings


class CloudNamespaceBindingsTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        return {
            "calls": calls,
            "cloud_ws_loop": "loop",
            "cloud_ws_conn": "ws",
            "cloud_connected": True,
            "CLOUD_AUTO_UPLOAD": True,
            "_cloud_identity_payload": lambda: {"imei": "861"},
            "_cloud_log": lambda *args, **kwargs: calls.append(("log", args, kwargs)),
        }

    def test_install_cloud_namespace_bindings_registers_expected_names(self):
        namespace = self.make_namespace()

        result = bindings.install_cloud_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        for name in (
            "_cloud_runtime_imei",
            "_cloud_send_register",
            "_cloud_send_unregister",
            "_cloud_schedule_unregister",
            "_cloud_unregister_then_close",
            "save_cloud_control_setting",
            "_cloud_send_payload",
            "_cloud_send_call_event",
            "_cloud_reply",
            "_handle_cloud_message",
            "start_cloud_control",
            "stop_cloud_control",
            "restart_cloud_control",
            "refresh_cloud_control_settings_from_config",
            "open_cloud_control_window",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_sync_bindings_forward_namespace_and_arguments(self):
        namespace = self.make_namespace()
        bindings.install_cloud_namespace_bindings(namespace)

        with patch.object(bindings, "cloud_runtime_imei_namespace_runtime", return_value="861") as runtime_imei, \
                patch.object(bindings, "save_cloud_control_setting_namespace_runtime", return_value="saved") as save_setting, \
                patch.object(bindings, "cloud_log_namespace_runtime", return_value="logged") as cloud_log, \
                patch.object(bindings, "send_cloud_serial_log_namespace_runtime", return_value="serial_log") as serial_log, \
                patch.object(bindings, "refresh_cloud_control_settings_namespace_runtime", return_value=True) as refresh_settings, \
                patch.object(bindings, "start_cloud_control_namespace_runtime", return_value=True) as start_control:
            self.assertEqual(namespace["_cloud_runtime_imei"](), "861")
            self.assertEqual(namespace["save_cloud_control_setting"](enabled=True, url="ws://host"), "saved")
            self.assertEqual(namespace["_cloud_log"]("hello", show_main=True), "logged")
            self.assertEqual(namespace["_cloud_send_serial_log"]("line"), "serial_log")
            self.assertTrue(namespace["refresh_cloud_control_settings_from_config"]())
            self.assertTrue(namespace["start_cloud_control"](show_errors=True))

        runtime_imei.assert_called_once_with(namespace)
        self.assertEqual(save_setting.call_args.args, (namespace,))
        self.assertTrue(save_setting.call_args.kwargs["enabled"])
        self.assertEqual(save_setting.call_args.kwargs["url"], "ws://host")
        cloud_log.assert_called_once_with(namespace, "hello", show_main=True)
        serial_log.assert_called_once_with(namespace, "line")
        refresh_settings.assert_called_once_with(namespace)
        start_control.assert_called_once_with(namespace, show_errors=True)

    def test_schedule_and_close_bindings_forward_connection_context(self):
        namespace = self.make_namespace()
        bindings.install_cloud_namespace_bindings(namespace)

        with patch.object(bindings, "schedule_cloud_unregister_runtime", return_value="scheduled") as schedule_runtime:
            self.assertEqual(namespace["_cloud_schedule_unregister"]("hidden"), "scheduled")

        kwargs = schedule_runtime.call_args.kwargs
        self.assertEqual(kwargs["reason"], "hidden")
        self.assertEqual(kwargs["get_loop"](), "loop")
        self.assertEqual(kwargs["get_ws"](), "ws")
        self.assertTrue(kwargs["is_connected"]())
        self.assertIs(kwargs["send_unregister"], namespace["_cloud_send_unregister"])

        async def close_runtime(ws, **kwargs):
            return ("closed", ws, kwargs)

        with patch.object(bindings, "unregister_then_close_cloud_connection_runtime", side_effect=close_runtime):
            result = asyncio.run(namespace["_cloud_unregister_then_close"]("ws2", reason="disconnect"))

        self.assertEqual(result[0:2], ("closed", "ws2"))
        self.assertTrue(result[2]["auto_upload"])
        self.assertIs(result[2]["send_unregister"], namespace["_cloud_send_unregister"])

    def test_async_bindings_forward_to_underlying_runtimes(self):
        namespace = self.make_namespace()
        bindings.install_cloud_namespace_bindings(namespace)

        async def wait_runtime(ns, ws, timeout):
            return ("ack", ns is namespace, ws, timeout)

        async def payload_runtime(ws, payload):
            return ("payload", ws, payload)

        async def reply_runtime(ws, payload, **kwargs):
            return ("reply", ws, payload, kwargs["identity_payload"]())

        with patch.object(bindings, "wait_cloud_login_ack_namespace_runtime", side_effect=wait_runtime), \
                patch.object(bindings, "cloud_identity_payload_namespace_runtime", return_value={"imei": "861"}), \
                patch.object(bindings, "send_cloud_payload_runtime", side_effect=payload_runtime), \
                patch.object(bindings, "reply_cloud_payload_runtime", side_effect=reply_runtime):
            ack = asyncio.run(namespace["_cloud_wait_login_ack"]("ws", timeout=3.0))
            payload = asyncio.run(namespace["_cloud_send_payload"]("ws", {"type": "x"}))
            reply = asyncio.run(namespace["_cloud_reply"]("ws", {"ok": True}))

        self.assertEqual(ack, ("ack", True, "ws", 3.0))
        self.assertEqual(payload, ("payload", "ws", {"type": "x"}))
        self.assertEqual(reply, ("reply", "ws", {"ok": True}, {"imei": "861"}))


if __name__ == "__main__":
    unittest.main()
