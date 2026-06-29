import configparser
import unittest

from sms_ui.desktop_shortcut_runtime import (
    DEFAULT_DESKTOP_SHORTCUT_NAME,
    desktop_shortcut_default_name,
    open_desktop_shortcut_dialog_runtime,
    save_desktop_shortcut_name_runtime,
)


class DesktopShortcutRuntimeTests(unittest.TestCase):
    def test_desktop_shortcut_default_name_uses_fallback(self):
        self.assertEqual(
            desktop_shortcut_default_name(configparser.ConfigParser()),
            DEFAULT_DESKTOP_SHORTCUT_NAME,
        )

    def test_save_desktop_shortcut_name_runtime_creates_ui_section(self):
        config = configparser.ConfigParser()
        calls = []

        result = save_desktop_shortcut_name_runtime(config, "My SMS", lambda: calls.append("save"))

        self.assertEqual(config.get("ui", "desktop_shortcut_name"), "My SMS")
        self.assertEqual(calls, ["save"])
        self.assertTrue(result)

    def test_save_desktop_shortcut_name_runtime_rejects_save_failure(self):
        config = configparser.ConfigParser()

        with self.assertRaises(RuntimeError):
            save_desktop_shortcut_name_runtime(config, "My SMS", lambda: False)

    def test_open_desktop_shortcut_dialog_runtime_applies_create_now(self):
        config = configparser.ConfigParser()
        calls = []

        def open_dialog(parent, default_name, on_apply, on_save, center_window):
            calls.append(("open", parent, default_name, center_window))
            on_apply("Desk Name")

        open_desktop_shortcut_dialog_runtime(
            "root",
            config=config,
            save_config=lambda: calls.append(("save",)),
            create_shortcut=lambda name: calls.append(("create", name)),
            system_ui=lambda *args: calls.append(("ui", args)),
            center_window="center",
            open_dialog=open_dialog,
        )

        self.assertEqual(calls[0], ("open", "root", DEFAULT_DESKTOP_SHORTCUT_NAME, "center"))
        self.assertEqual(calls[1], ("create", "Desk Name"))
        self.assertEqual(calls[2], ("save",))
        self.assertIn("Desk Name.lnk", calls[3][1][0])
        self.assertEqual(config.get("ui", "desktop_shortcut_name"), "Desk Name")

    def test_open_desktop_shortcut_dialog_runtime_saves_only(self):
        config = configparser.ConfigParser()
        config["ui"] = {"desktop_shortcut_name": "Old"}
        calls = []

        def open_dialog(_parent, default_name, _on_apply, on_save, _center_window):
            calls.append(("default", default_name))
            on_save("New")

        open_desktop_shortcut_dialog_runtime(
            "root",
            config=config,
            save_config=lambda: calls.append(("save",)),
            create_shortcut=lambda name: calls.append(("create", name)),
            system_ui=lambda *args: calls.append(("ui", args)),
            center_window="center",
            open_dialog=open_dialog,
        )

        self.assertEqual(calls[0], ("default", "Old"))
        self.assertEqual(calls[1], ("save",))
        self.assertIn("New", calls[2][1][0])
        self.assertEqual(config.get("ui", "desktop_shortcut_name"), "New")

    def test_open_desktop_shortcut_dialog_runtime_propagates_save_failure(self):
        config = configparser.ConfigParser()

        def open_dialog(_parent, _default_name, _on_apply, on_save, _center_window):
            on_save("New")

        with self.assertRaises(RuntimeError):
            open_desktop_shortcut_dialog_runtime(
                "root",
                config=config,
                save_config=lambda: False,
                create_shortcut=lambda name: None,
                system_ui=lambda *args: None,
                center_window="center",
                open_dialog=open_dialog,
            )


if __name__ == "__main__":
    unittest.main()
