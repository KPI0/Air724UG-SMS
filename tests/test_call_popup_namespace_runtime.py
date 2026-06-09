import unittest

from sms_ui.call_popup_namespace_runtime import (
    close_call_popup_namespace_runtime,
    get_serial_call_state_namespace_runtime,
    set_call_popup_namespace_runtime,
    set_serial_call_state_namespace_runtime,
    show_call_popup_namespace_runtime,
)


class CallPopupNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        return {
            "root": "root",
            "current_call_popup": "popup",
            "ring_timeout_target": 3.0,
            "current_dial_num": "10086",
            "center_window": "center",
            "serial_lock": "lock",
            "serial_obj": "serial",
            "port_ui": lambda *args: ("port", args),
            "set_status": lambda *args: ("status", args),
            "ui_post": "post",
            "close_call_popup": lambda: "closed",
            "run_on_ui_thread": lambda callback, ui_post: callback(),
        }

    def test_set_call_popup_namespace_runtime_updates_window(self):
        namespace = self.base_namespace()

        result = set_call_popup_namespace_runtime(namespace, "next")

        self.assertIsNone(result)
        self.assertEqual(namespace["current_call_popup"], "next")

    def test_close_call_popup_namespace_runtime_forwards_state_callbacks(self):
        namespace = self.base_namespace()
        calls = []

        result = close_call_popup_namespace_runtime(
            namespace,
            close_app_runtime=lambda **kwargs: calls.append(kwargs) or "closed",
        )

        self.assertEqual(result, "closed")
        forwarded = calls[0]
        self.assertEqual(forwarded["get_popup"](), "popup")
        forwarded["set_popup"](None)
        self.assertIsNone(namespace["current_call_popup"])
        self.assertEqual(forwarded["ui_post"], "post")

    def test_show_call_popup_namespace_runtime_forwards_dependencies_and_sets_ring_timeout(self):
        namespace = self.base_namespace()
        calls = []

        result = show_call_popup_namespace_runtime(
            namespace,
            "13800138000",
            show_app_runtime=lambda **kwargs: calls.append(kwargs) or "shown",
        )

        self.assertEqual(result, "shown")
        forwarded = calls[0]
        self.assertEqual(forwarded["parent"], "root")
        self.assertEqual(forwarded["caller_num"], "13800138000")
        self.assertEqual(forwarded["get_popup"](), "popup")
        forwarded["set_popup"]("next_popup")
        self.assertEqual(namespace["current_call_popup"], "next_popup")
        self.assertEqual(forwarded["get_serial"](), "serial")
        forwarded["set_ring_timeout"](-1.0)
        self.assertEqual(namespace["ring_timeout_target"], -1.0)

    def test_serial_call_state_namespace_runtime_reads_and_writes_state(self):
        namespace = self.base_namespace()

        self.assertEqual(get_serial_call_state_namespace_runtime(namespace), (3.0, "10086"))

        result = set_serial_call_state_namespace_runtime(namespace, 0.0, "")

        self.assertIsNone(result)
        self.assertEqual(namespace["ring_timeout_target"], 0.0)
        self.assertEqual(namespace["current_dial_num"], "")


if __name__ == "__main__":
    unittest.main()
