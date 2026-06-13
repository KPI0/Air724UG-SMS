import configparser
import unittest

from sms_ui.settings_runtime import (
    multi_instance_status,
    save_multi_instance_config,
    toggle_multi_instance_runtime,
)


class MultiInstanceSettingsRuntimeTests(unittest.TestCase):
    def test_save_multi_instance_config_writes_ui_flag(self):
        config = configparser.ConfigParser()
        saved = []

        save_multi_instance_config(config, True, lambda: saved.append("saved"))

        self.assertEqual(config.get("ui", "allow_multi_instance"), "1")
        self.assertEqual(saved, ["saved"])

    def test_save_multi_instance_config_logs_save_errors(self):
        config = configparser.ConfigParser()
        logs = []

        save_multi_instance_config(
            config,
            True,
            lambda: (_ for _ in ()).throw(RuntimeError("save failed")),
            log_error=logs.append,
        )

        self.assertEqual(len(logs), 1)
        self.assertIn("save failed", logs[0])

    def test_multi_instance_status_formats_enabled_state(self):
        self.assertIn("\u5df2\u5f00\u542f", multi_instance_status(True))
        self.assertIn("\u5df2\u5173\u95ed", multi_instance_status(False))

    def test_toggle_multi_instance_runtime_updates_state_config_and_ui(self):
        config = configparser.ConfigParser()
        saved = []
        state = []
        messages = []

        toggle_multi_instance_runtime(
            True,
            config,
            lambda: saved.append("saved"),
            state.append,
            lambda *args: messages.append(args),
        )

        self.assertEqual(config.get("ui", "allow_multi_instance"), "1")
        self.assertEqual(saved, ["saved"])
        self.assertEqual(state, [True])
        self.assertEqual(messages[0][1], "normal")


if __name__ == "__main__":
    unittest.main()
