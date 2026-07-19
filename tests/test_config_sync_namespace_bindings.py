import unittest
from unittest.mock import patch

import sms_ui.config_sync_namespace_bindings as bindings


class ConfigSyncNamespaceBindingsTests(unittest.TestCase):
    def test_install_and_forward_bindings(self):
        namespace = {}
        self.assertIs(bindings.install_config_sync_namespace_bindings(namespace), namespace)

        with patch.object(
            bindings,
            "register_config_sync_refresher_namespace_runtime",
            return_value="unregister",
        ) as register_runtime, patch.object(
            bindings,
            "reload_shared_ui_config_namespace_runtime",
            return_value=("短信弹窗",),
        ) as reload_runtime, patch.object(
            bindings,
            "start_config_file_watch_namespace_runtime",
            return_value="watch-id",
        ) as start_runtime:
            callback = lambda: None
            self.assertEqual(
                namespace["register_config_sync_refresher"]("keywords", callback),
                "unregister",
            )
            self.assertEqual(namespace["reload_shared_ui_config"](), ("短信弹窗",))
            self.assertEqual(namespace["start_config_file_watch"](2500), "watch-id")

        register_runtime.assert_called_once_with(namespace, group="keywords", callback=callback)
        reload_runtime.assert_called_once_with(namespace)
        start_runtime.assert_called_once_with(namespace, interval_ms=2500)


if __name__ == "__main__":
    unittest.main()
