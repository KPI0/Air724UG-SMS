import unittest
from unittest.mock import patch

import sms_ui.third_push_namespace_bindings as bindings


class ThirdPushNamespaceBindingsTests(unittest.TestCase):
    def test_install_registers_expected_names(self):
        namespace = {}

        result = bindings.install_third_push_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        for name in (
            "ensure_third_push_config",
            "apply_third_push_settings",
            "refresh_third_push_settings_from_config",
            "save_third_push_setting",
            "_third_push_worker",
            "show_third_push_test_result",
            "enqueue_third_push",
            "open_third_push_window",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_bindings_forward_namespace_and_arguments(self):
        namespace = {}
        bindings.install_third_push_namespace_bindings(namespace)

        with patch.object(bindings, "ensure_third_push_config_namespace_runtime", return_value="ensured") as ensure, \
                patch.object(bindings, "apply_third_push_settings_namespace_runtime", return_value="applied") as apply, \
                patch.object(bindings, "refresh_third_push_settings_namespace_runtime", return_value="refreshed") as refresh, \
                patch.object(bindings, "save_third_push_setting_namespace_runtime", return_value="saved") as save, \
                patch.object(bindings, "third_push_worker_namespace_runtime", return_value="worked") as worker, \
                patch.object(bindings, "show_third_push_test_result_namespace_runtime", return_value="shown") as show, \
                patch.object(bindings, "enqueue_third_push_namespace_runtime", return_value="queued") as enqueue, \
                patch.object(bindings, "open_third_push_window_namespace_runtime", return_value="opened") as open_window:
            self.assertEqual(namespace["ensure_third_push_config"](save=True), "ensured")
            self.assertEqual(namespace["apply_third_push_settings"]("settings"), "applied")
            self.assertEqual(namespace["refresh_third_push_settings_from_config"](), "refreshed")
            self.assertEqual(namespace["save_third_push_setting"](enabled=True, notify_type="bark"), "saved")
            self.assertEqual(namespace["_third_push_worker"](), "worked")
            self.assertEqual(namespace["show_third_push_test_result"](["ok"], ["fail"]), "shown")
            self.assertEqual(
                namespace["enqueue_third_push"](
                    "raw",
                    show_success=True,
                    show_result=True,
                    channels=["bark"],
                    settings={"bark_url": "url"},
                    template="{msg}",
                    event_type="call",
                ),
                "queued",
            )
            self.assertEqual(namespace["open_third_push_window"](), "opened")

        ensure.assert_called_once_with(namespace, save=True)
        apply.assert_called_once_with(namespace, "settings")
        refresh.assert_called_once_with(namespace)
        self.assertEqual(save.call_args.args, (namespace,))
        self.assertTrue(save.call_args.kwargs["enabled"])
        self.assertEqual(save.call_args.kwargs["notify_type"], "bark")
        worker.assert_called_once_with(namespace)
        show.assert_called_once_with(namespace, ["ok"], ["fail"])
        self.assertEqual(enqueue.call_args.args, (namespace, "raw"))
        self.assertEqual(enqueue.call_args.kwargs["channels"], ["bark"])
        self.assertEqual(enqueue.call_args.kwargs["event_type"], "call")
        open_window.assert_called_once_with(namespace)


if __name__ == "__main__":
    unittest.main()
