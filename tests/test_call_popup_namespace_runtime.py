import unittest

from types import SimpleNamespace

from sms_ui.call_popup_namespace_runtime import (
    close_call_popup_namespace_runtime,
    close_missed_call_popup_namespace_runtime,
    close_phone_popups_namespace_runtime,
    finish_incoming_call_session_namespace_runtime,
    get_serial_call_state_namespace_runtime,
    mark_incoming_call_handled_namespace_runtime,
    reset_incoming_call_session_namespace_runtime,
    set_call_popup_namespace_runtime,
    set_missed_call_popup_namespace_runtime,
    set_serial_call_state_namespace_runtime,
    show_call_popup_namespace_runtime,
    show_missed_call_popup_namespace_runtime,
    start_incoming_call_session_namespace_runtime,
)


class CallPopupNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        tracker = SimpleNamespace(
            start=lambda caller_num: ("start", caller_num),
            mark_handled=lambda: "handled",
            finish=lambda: "finished",
            reset=lambda: "reset",
        )
        return {
            "root": "root",
            "current_call_popup": "popup",
            "current_missed_call_popup": "missed_popup",
            "INCOMING_CALL_SESSION": tracker,
            "CALL_POPUP_ENABLED": True,
            "ring_timeout_target": 3.0,
            "current_dial_num": "10086",
            "center_window": "center",
            "serial_lock": "lock",
            "serial_obj": "serial",
            "port_ui": lambda *args: ("port", args),
            "set_status": lambda *args: ("status", args),
            "ui_post": "post",
            "close_call_popup": lambda: "closed",
            "close_missed_call_popup": lambda: "missed_closed",
            "show_window": "show_window",
            "run_on_ui_thread": lambda callback, ui_post: callback(),
            "log_file_only": lambda message: ("log", message),
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
        self.assertEqual(forwarded["log_error"]("close log"), ("log", "close log"))

    def test_close_missed_call_popup_namespace_runtime_forwards_state_callbacks(self):
        namespace = self.base_namespace()
        calls = []

        result = close_missed_call_popup_namespace_runtime(
            namespace,
            close_app_runtime=lambda **kwargs: calls.append(kwargs) or "closed",
        )

        self.assertEqual(result, "closed")
        forwarded = calls[0]
        self.assertEqual(forwarded["get_popup"](), "missed_popup")
        forwarded["set_popup"](None)
        self.assertIsNone(namespace["current_missed_call_popup"])

    def test_close_phone_popups_namespace_runtime_closes_both_windows(self):
        namespace = self.base_namespace()
        calls = []
        namespace["close_call_popup"] = lambda: calls.append("incoming")
        namespace["close_missed_call_popup"] = lambda: calls.append("missed")

        close_phone_popups_namespace_runtime(namespace)

        self.assertEqual(calls, ["incoming", "missed"])

    def test_close_phone_popups_continues_when_one_window_close_fails(self):
        namespace = self.base_namespace()
        calls = []

        def fail_close():
            raise RuntimeError("failed")

        namespace["close_call_popup"] = fail_close
        namespace["close_missed_call_popup"] = lambda: calls.append("missed")
        namespace["log_file_only"] = calls.append

        close_phone_popups_namespace_runtime(namespace)

        self.assertEqual(calls[-1], "missed")
        self.assertTrue(any("close_call_popup" in str(item) for item in calls))

    def test_show_call_popup_namespace_runtime_forwards_dependencies_and_sets_ring_timeout(self):
        namespace = self.base_namespace()
        calls = []

        result = show_call_popup_namespace_runtime(
            namespace,
            "13123123123",
            show_app_runtime=lambda **kwargs: calls.append(kwargs) or "shown",
        )

        self.assertEqual(result, "shown")
        forwarded = calls[0]
        self.assertEqual(forwarded["parent"], "root")
        self.assertEqual(forwarded["caller_num"], "13123123123")
        self.assertEqual(forwarded["get_popup"](), "popup")
        forwarded["set_popup"]("next_popup")
        self.assertEqual(namespace["current_call_popup"], "next_popup")
        self.assertEqual(forwarded["get_serial"](), "serial")
        forwarded["set_ring_timeout"](-1.0)
        self.assertEqual(namespace["ring_timeout_target"], -1.0)
        self.assertEqual(forwarded["log_error"]("show log"), ("log", "show log"))
        self.assertEqual(forwarded["mark_call_handled"](), "handled")
        self.assertTrue(forwarded["is_enabled"]())
        namespace["CALL_POPUP_ENABLED"] = False
        self.assertFalse(forwarded["is_enabled"]())

    def test_missed_popup_namespace_runtime_forwards_singleton_state(self):
        namespace = self.base_namespace()
        calls = []
        missed_call = object()

        result = show_missed_call_popup_namespace_runtime(
            namespace,
            missed_call,
            show_app_runtime=lambda **kwargs: calls.append(kwargs) or "shown",
        )

        self.assertEqual(result, "shown")
        forwarded = calls[0]
        self.assertIs(forwarded["missed_call"], missed_call)
        self.assertEqual(forwarded["get_popup"](), "missed_popup")
        forwarded["set_popup"]("next")
        self.assertEqual(namespace["current_missed_call_popup"], "next")
        self.assertEqual(forwarded["show_window"], "show_window")
        self.assertTrue(forwarded["is_enabled"]())

    def test_call_session_namespace_runtime_forwards_tracker_actions(self):
        namespace = self.base_namespace()

        self.assertEqual(
            start_incoming_call_session_namespace_runtime(namespace, "10086"),
            ("start", "10086"),
        )
        self.assertEqual(mark_incoming_call_handled_namespace_runtime(namespace), "handled")
        self.assertEqual(finish_incoming_call_session_namespace_runtime(namespace), "finished")
        self.assertEqual(reset_incoming_call_session_namespace_runtime(namespace), "reset")
        self.assertIsNone(set_missed_call_popup_namespace_runtime(namespace, None))
        self.assertIsNone(namespace["current_missed_call_popup"])

    def test_phone_popup_setting_disables_incoming_and_missed_windows(self):
        namespace = self.base_namespace()
        namespace["CALL_POPUP_ENABLED"] = False

        self.assertEqual(
            show_call_popup_namespace_runtime(
                namespace,
                "10086",
                show_app_runtime=lambda **_kwargs: self.fail("incoming popup opened"),
            ),
            "disabled",
        )
        self.assertEqual(
            show_missed_call_popup_namespace_runtime(
                namespace,
                object(),
                show_app_runtime=lambda **_kwargs: self.fail("missed popup opened"),
            ),
            "disabled",
        )

    def test_serial_call_state_namespace_runtime_reads_and_writes_state(self):
        namespace = self.base_namespace()

        self.assertEqual(get_serial_call_state_namespace_runtime(namespace), (3.0, "10086"))

        result = set_serial_call_state_namespace_runtime(namespace, 0.0, "")

        self.assertIsNone(result)
        self.assertEqual(namespace["ring_timeout_target"], 0.0)
        self.assertEqual(namespace["current_dial_num"], "")


if __name__ == "__main__":
    unittest.main()
