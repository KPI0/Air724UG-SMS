import threading
import unittest
from unittest.mock import patch

import sms_ui.app_lifecycle_namespace_bindings as bindings
from sms_ui.thread_runtime import run_on_ui_thread


class FakeVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


class AppLifecycleNamespaceBindingsTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        return {
            "calls": calls,
            "run_on_ui_thread": lambda callback, ui_post: calls.append(("run", ui_post)) or callback(),
            "ui_post": "ui_post",
            "autostart_var": FakeVar(True),
        }

    def test_install_registers_expected_names(self):
        namespace = self.make_namespace()

        result = bindings.install_app_lifecycle_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        for name in (
            "set_autostart",
            "cleanup_and_exit",
            "toggle_voice_broadcast",
            "toggle_multi_instance",
            "toggle_autostart",
            "toggle_popup",
            "toggle_call_popup",
            "restart_software",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_bindings_forward_namespace_and_arguments(self):
        namespace = self.make_namespace()
        bindings.install_app_lifecycle_namespace_bindings(namespace)

        with patch.object(bindings, "set_autostart_namespace_runtime", return_value="autostart") as set_autostart, \
                patch.object(bindings, "cleanup_and_exit_namespace_runtime", return_value="cleaned") as cleanup, \
                patch.object(bindings, "toggle_voice_broadcast_namespace_runtime", return_value="voice") as voice, \
                patch.object(bindings, "toggle_multi_instance_namespace_runtime", return_value="multi") as multi, \
                patch.object(bindings, "toggle_popup_namespace_runtime", return_value="popup") as popup, \
                patch.object(bindings, "toggle_call_popup_namespace_runtime", return_value="call-popup") as call_popup, \
                patch.object(bindings, "restart_software_namespace_runtime", return_value="restart") as restart:
            self.assertEqual(namespace["set_autostart"](False), "autostart")
            self.assertEqual(namespace["cleanup_and_exit"](), "cleaned")
            self.assertEqual(namespace["toggle_voice_broadcast"](), "voice")
            self.assertEqual(namespace["toggle_multi_instance"](), "multi")
            self.assertEqual(namespace["toggle_autostart"](), "autostart")
            self.assertEqual(namespace["toggle_popup"](), "popup")
            self.assertEqual(namespace["toggle_call_popup"](), "call-popup")
            self.assertEqual(namespace["restart_software"](), "restart")

        self.assertEqual(set_autostart.call_args_list[0].args, (namespace, False))
        self.assertEqual(set_autostart.call_args_list[1].args, (namespace, True))
        cleanup.assert_called_once_with(namespace)
        self.assertIn(("run", "ui_post"), namespace["calls"])
        voice.assert_called_once_with(namespace)
        multi.assert_called_once_with(namespace)
        popup.assert_called_once_with(namespace)
        call_popup.assert_called_once_with(namespace)
        restart.assert_called_once_with(namespace)

    def test_cleanup_from_background_thread_waits_for_ui_queue(self):
        namespace = self.make_namespace()
        posted = []
        namespace["run_on_ui_thread"] = run_on_ui_thread
        namespace["ui_post"] = posted.append
        bindings.install_app_lifecycle_namespace_bindings(namespace)

        with patch.object(bindings, "cleanup_and_exit_namespace_runtime", return_value="cleaned") as cleanup:
            worker = threading.Thread(target=namespace["cleanup_and_exit"])
            worker.start()
            worker.join(timeout=1)

            self.assertFalse(worker.is_alive())
            cleanup.assert_not_called()
            self.assertEqual(len(posted), 1)

            self.assertEqual(posted[0](), "cleaned")

        cleanup.assert_called_once_with(namespace)


if __name__ == "__main__":
    unittest.main()
