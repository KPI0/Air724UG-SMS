import configparser
import gc
import os
import tempfile
import tkinter as tk
import unittest
import weakref
from unittest.mock import patch

from sms_core.config_runtime import remember_config_snapshot, safe_save_config_runtime
from sms_ui.config_sync_namespace_bindings import install_config_sync_namespace_bindings
from sms_ui.config_sync_namespace_runtime import (
    refresh_autostart_menu_namespace_runtime,
    start_config_file_watch_namespace_runtime,
)
from sms_ui.app_infrastructure_namespace_runtime import safe_save_config_namespace_runtime
from sms_ui.maintenance_namespace_runtime import (
    open_log_cleanup_dialog_namespace_runtime,
    open_update_proxy_dialog_namespace_runtime,
)
from sms_ui.settings_namespace_runtime import (
    open_desktop_shortcut_dialog_namespace_runtime,
    open_sms_font_dialog_namespace_runtime,
    open_voice_text_dialog_namespace_runtime,
)
from sms_ui import sms_font_dialog, utility_dialogs
from tests import test_config_sync_namespace_runtime as fixtures
from tests.test_config_sync_runtime import FakeRootTimer


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


class SettingsWindowSyncTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        self.root.withdraw()
        self.root.update()
        self.directory = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.directory.name, "config.ini")
        config = configparser.ConfigParser(interpolation=None)
        config["ui"] = {
            "voice_text": "old voice", "sms_font_size": "30",
            "sms_font_color": "#ff0000", "log_retention_days": "30",
            "desktop_shortcut_name": "Old name",
        }
        config["update"] = {
            "api_proxy_base": "https://old-api.example.invalid/",
            "proxy_base": "https://old-download.example.invalid/",
        }
        config["cloud_control"] = {"auto_upload": "0"}
        with open(self.path, "w", encoding="utf-8") as file_obj:
            config.write(file_obj)
        self.first, self.second = [self.make_namespace() for _ in range(2)]

    def tearDown(self):
        self.root.update()
        self.root.destroy()
        self.directory.cleanup()

    def make_namespace(self):
        namespace = fixtures.ConfigSyncNamespaceRuntimeTests().make_namespace()
        namespace["config"].clear()
        namespace["config"].read(self.path, encoding="utf-8")
        remember_config_snapshot(namespace["config"])
        namespace.update({
            "root": self.root, "CONFIG_FILE": self.path,
            "AUTO_CLEANUP_INTERVAL_HOURS": 6,
            "GITHUB_OWNER": "owner", "GITHUB_REPO": "repo",
            "ui_post": lambda callback: callback(),
            "center_window": lambda *_: None,
            "_ui_open_sms_font_dialog": sms_font_dialog.open_sms_font_dialog,
            "_ui_open_voice_text_dialog": utility_dialogs.open_voice_text_dialog,
            "_ui_open_desktop_shortcut_dialog": utility_dialogs.open_desktop_shortcut_dialog,
            "create_desktop_shortcut": lambda name: namespace["calls"].append(("shortcut", name)),
            "CLOUD_AUTO_UPLOAD": False,
        })
        namespace["safe_save_config"] = lambda: safe_save_config_runtime(
            config=namespace["config"], config_file=self.path,
            config_lock=namespace["CONFIG_LOCK"],
        )
        install_config_sync_namespace_bindings(namespace)
        namespace["reload_shared_ui_config"]()
        namespace["calls"].clear()
        return namespace

    def save_external(self, section, **values):
        for key, value in values.items():
            self.first["config"].set(section, key, str(value))
        self.assertTrue(self.first["safe_save_config"]())
        return self.second["reload_shared_ui_config"]()

    def open_window(self, opener):
        opener(self.second)
        self.root.update()
        return self.root.grab_current()

    def variable(self, widget):
        return tk.StringVar(self.root, name=widget.cget("textvariable"))

    def button(self, win, label):
        return next(widget for widget in descendants(win)
                    if isinstance(widget, tk.Button) and widget.cget("text") == label)

    def test_voice_and_cleanup_editors_refresh_clean_values_and_keep_drafts(self):
        for kind in ("voice", "cleanup"):
            with self.subTest(kind=kind):
                is_voice = kind == "voice"
                opener = open_voice_text_dialog_namespace_runtime if is_voice else open_log_cleanup_dialog_namespace_runtime
                win = self.open_window(opener)
                if is_voice:
                    widget = next(w for w in descendants(win) if isinstance(w, tk.Text))
                    read = lambda: widget.get("1.0", "end-1c")

                    def write(value):
                        widget.delete("1.0", "end")
                        widget.insert("1.0", value)

                    key, external, draft = "voice_text", "external voice", "unfinished voice"
                else:
                    widget = next(w for w in descendants(win) if isinstance(w, tk.Entry))
                    variable = self.variable(widget)
                    read, write = variable.get, variable.set
                    key, external, draft = "log_retention_days", "7", "12"
                self.save_external("ui", **{key: external})
                self.assertEqual(read(), external)
                write(draft)
                self.save_external("ui", **{key: "next voice" if is_voice else "2"})
                self.assertEqual(read(), draft)
                group = "voice_text" if is_voice else "log_cleanup"
                self.assertIn(group, self.second["_CONFIG_SYNC_REFRESHERS"])
                win.destroy()
                self.assertNotIn(group, self.second["_CONFIG_SYNC_REFRESHERS"])

    def test_font_keeps_edited_size_and_refreshes_color_with_preview(self):
        win = self.open_window(open_sms_font_dialog_namespace_runtime)
        widgets = list(descendants(win))
        size = self.variable(next(w for w in widgets if isinstance(w, tk.Spinbox)))
        color = self.variable(next(w for w in widgets if isinstance(w, tk.Entry)))
        canvas = next(w for w in widgets if isinstance(w, tk.Canvas))
        size.set("41")
        self.save_external("ui", sms_font_size=24, sms_font_color="#123456")
        self.assertEqual((size.get(), color.get()), ("41", "#123456"))
        self.assertEqual(canvas.itemcget(canvas.find_all()[0], "fill"), "#123456")
        color.set("invalid draft color")
        self.save_external("ui", sms_font_color="#abcdef")
        self.assertEqual(color.get(), "invalid draft color")
        win.destroy()
        self.assertNotIn("sms_font", self.second["_CONFIG_SYNC_REFRESHERS"])
        # Closed font editors must no longer have preview traces.
        size.set("50")
        color.set("blue")

    def test_proxy_refreshes_untouched_field_and_rebases_after_successful_save(self):
        win = self.open_window(open_update_proxy_dialog_namespace_runtime)
        entries = [w for w in descendants(win) if isinstance(w, tk.Entry)]
        api, download = [self.variable(w) for w in entries]
        api.set("draft-api.example.invalid")
        changed = self.save_external(
            "update", api_proxy_base="https://external-api.example.invalid/",
            proxy_base="https://external-download.example.invalid/",
        )
        self.assertIn("更新代理", changed)
        self.assertEqual(api.get(), "draft-api.example.invalid")
        self.assertEqual(download.get(), "https://external-download.example.invalid/")
        with patch("sms_ui.update_proxy_dialog.messagebox.showinfo"):
            self.button(win, "保存").invoke()
        self.assertEqual(api.get(), "https://draft-api.example.invalid/")
        self.save_external("update", api_proxy_base="https://next-api.example.invalid/")
        self.assertEqual(api.get(), "https://next-api.example.invalid/")
        self.assertEqual(self.second["reload_shared_ui_config"](), ())
        win.destroy()
        self.assertNotIn("update_proxy", self.second["_CONFIG_SYNC_REFRESHERS"])

    def test_failed_proxy_save_preserves_draft_and_clean_field_tracking(self):
        self.second["safe_save_config"] = lambda: False
        win = self.open_window(open_update_proxy_dialog_namespace_runtime)
        api, download = [self.variable(w) for w in descendants(win) if isinstance(w, tk.Entry)]
        api.set("draft-api.example.invalid")
        with patch("sms_ui.update_proxy_dialog.messagebox.showerror") as show_error:
            self.button(win, "保存").invoke()
        show_error.assert_called_once()
        self.save_external("update", api_proxy_base="https://new-api.example.invalid/",
                           proxy_base="https://new-download.example.invalid/")
        self.assertEqual(api.get(), "draft-api.example.invalid")
        self.assertEqual(download.get(), "https://new-download.example.invalid/")
        win.destroy()

    def test_shortcut_name_sync_does_not_create_shortcuts_and_rebases_saved_form(self):
        win = self.open_window(open_desktop_shortcut_dialog_namespace_runtime)
        name = self.variable(next(w for w in descendants(win) if isinstance(w, tk.Entry)))
        changed = self.save_external("ui", desktop_shortcut_name="External name")
        self.assertIn("快捷方式名称", changed)
        self.assertEqual(name.get(), "External name")
        self.assertFalse(any(call[0] == "shortcut" for call in self.second["calls"]))
        name.set("Draft name")
        self.save_external("ui", desktop_shortcut_name="Next name")
        self.assertEqual(name.get(), "Draft name")
        with patch("sms_ui.utility_dialogs.messagebox.showinfo"):
            self.button(win, "保存").invoke()
        self.save_external("ui", desktop_shortcut_name="Newest name")
        self.assertEqual(name.get(), "Newest name")
        with patch("sms_ui.utility_dialogs.messagebox.showinfo"):
            name.set("Applied name")
            self.button(win, "应用").invoke()
        self.assertIn(("shortcut", "Applied name"), self.second["calls"])
        self.save_external("ui", desktop_shortcut_name="Final name")
        self.assertEqual(name.get(), "Final name")
        win.destroy()
        self.assertNotIn("desktop_shortcut", self.second["_CONFIG_SYNC_REFRESHERS"])

    def test_reload_after_local_save_merges_external_name_still_refreshes_open_editor(self):
        win = self.open_window(open_desktop_shortcut_dialog_namespace_runtime)
        name = self.variable(next(w for w in descendants(win) if isinstance(w, tk.Entry)))
        self.first["config"].set("ui", "desktop_shortcut_name", "External name")
        self.assertTrue(self.first["safe_save_config"]())
        self.second["config"].set("ui", "popup_enabled", "0")
        self.assertTrue(self.second["safe_save_config"]())
        self.assertEqual(self.second["config"].get("ui", "desktop_shortcut_name"), "External name")
        self.second["reload_shared_ui_config"]()
        self.assertEqual(name.get(), "External name")
        win.destroy()

    def test_publication_setting_syncs_without_connecting_disabled_cloud_control(self):
        self.second["restart_cloud_control"] = lambda **_: self.fail("cloud control is disabled")
        changed = self.save_external("cloud_control", auto_upload="1")
        self.assertTrue(self.second["CLOUD_AUTO_UPLOAD"])
        self.assertIn("主动公开设备", changed)

    def test_voice_save_before_watch_refresh_uses_merged_text_for_runtime(self):
        win = self.open_window(open_voice_text_dialog_namespace_runtime)
        self.first["config"].set("ui", "voice_text", "external voice")
        self.assertTrue(self.first["safe_save_config"]())
        self.button(win, "保存").invoke()
        self.assertEqual(self.second["config"].get("ui", "voice_text"), "external voice")
        self.assertEqual(self.second["VOICE_TEXT"], "external voice")
        self.assertFalse(win.winfo_exists())

    def test_font_save_before_watch_refresh_keeps_external_color_in_runtime(self):
        win = self.open_window(open_sms_font_dialog_namespace_runtime)
        size = self.variable(next(w for w in descendants(win) if isinstance(w, tk.Spinbox)))
        size.set("41")
        self.first["config"].set("ui", "sms_font_color", "#123456")
        self.assertTrue(self.first["safe_save_config"]())
        self.button(win, "保存").invoke()
        self.assertEqual(self.second["config"].get("ui", "sms_font_color"), "#123456")
        self.assertEqual((self.second["SMS_FONT_SIZE"], self.second["SMS_FONT_COLOR"]), (41, "#123456"))
        self.assertFalse(win.winfo_exists())

    def test_cleanup_apply_before_watch_refresh_uses_merged_days(self):
        win = self.open_window(open_log_cleanup_dialog_namespace_runtime)
        self.first["config"].set("ui", "log_retention_days", "7")
        self.assertTrue(self.first["safe_save_config"]())
        with patch("sms_ui.utility_dialogs.messagebox.askyesno", return_value=True), patch(
            "sms_ui.utility_dialogs.messagebox.showinfo",
        ):
            self.button(win, "确认").invoke()
        self.assertEqual(self.second["config"].get("ui", "log_retention_days"), "7")
        self.assertEqual(self.second["LOG_RETENTION_DAYS"], 7)
        self.assertFalse(win.winfo_exists())

    def test_shortcut_apply_before_watch_refresh_uses_saved_name(self):
        win = self.open_window(open_desktop_shortcut_dialog_namespace_runtime)
        self.first["config"].set("ui", "desktop_shortcut_name", "External name")
        self.assertTrue(self.first["safe_save_config"]())
        with patch("sms_ui.utility_dialogs.messagebox.showinfo"):
            self.button(win, "应用").invoke()
        self.assertEqual(self.second["config"].get("ui", "desktop_shortcut_name"), "External name")
        self.assertIn(("shortcut", "External name"), self.second["calls"])
        win.destroy()

    def test_production_ui_save_keeps_merged_font_values_after_worker_completes(self):
        self.second["safe_save_config"] = lambda: safe_save_config_namespace_runtime(self.second)
        win = self.open_window(open_sms_font_dialog_namespace_runtime)
        size = self.variable(next(w for w in descendants(win) if isinstance(w, tk.Spinbox)))
        size.set("41")
        self.first["config"].set("ui", "sms_font_color", "blue")
        self.assertTrue(self.first["safe_save_config"]())
        self.button(win, "保存").invoke()
        self.assertEqual((self.second["SMS_FONT_SIZE"], self.second["SMS_FONT_COLOR"]), (41, "blue"))
        self.assertFalse(self.second["_CONFIG_SAVE_ACTIVE"])
        self.assertNotIn("_CONFIG_SAVE_THREAD", self.second)

    def test_invalid_external_font_and_retention_values_keep_valid_runtime(self):
        win = self.open_window(open_sms_font_dialog_namespace_runtime)
        self.first["config"].set("ui", "sms_font_size", "invalid")
        self.first["config"].set("ui", "sms_font_color", "not-a-color")
        self.assertTrue(self.first["safe_save_config"]())
        self.button(win, "保存").invoke()
        self.assertEqual((self.second["SMS_FONT_SIZE"], self.second["SMS_FONT_COLOR"]), (30, "#ff0000"))
        win = self.open_window(open_log_cleanup_dialog_namespace_runtime)
        self.first["config"].set("ui", "log_retention_days", "-3")
        self.assertTrue(self.first["safe_save_config"]())
        with patch("sms_ui.utility_dialogs.messagebox.askyesno", return_value=True), patch(
            "sms_ui.utility_dialogs.messagebox.showinfo",
        ):
            self.button(win, "确认").invoke()
        self.assertEqual(self.second["LOG_RETENTION_DAYS"], 30)

    def test_syncing_dialogs_release_windows_subscriptions_and_tk_variables(self):
        references = []
        original_variables = set(self.root.tk.call("info", "vars", "PY_VAR*"))
        for opener in (
            open_sms_font_dialog_namespace_runtime,
            open_voice_text_dialog_namespace_runtime,
            open_log_cleanup_dialog_namespace_runtime,
            open_update_proxy_dialog_namespace_runtime,
            open_desktop_shortcut_dialog_namespace_runtime,
        ):
            for _ in range(3):
                win = self.open_window(opener)
                references.append(weakref.ref(win))
                win.destroy()
                del win
                self.root.update()
        gc.collect()
        self.assertTrue(all(reference() is None for reference in references))
        self.assertEqual(self.second["_CONFIG_SYNC_REFRESHERS"], {})
        self.assertEqual(set(self.root.tk.call("info", "vars", "PY_VAR*")), original_variables)

    def test_failed_editor_refresh_is_redacted_and_does_not_block_other_windows(self):
        win = self.open_window(open_desktop_shortcut_dialog_namespace_runtime)
        name = self.variable(next(w for w in descendants(win) if isinstance(w, tk.Entry)))

        def broken_refresh():
            raise ValueError("synthetic-sensitive-value")

        self.second["register_config_sync_refresher"]("update_proxy", broken_refresh)
        self.save_external("ui", desktop_shortcut_name="External name")
        self.assertEqual(name.get(), "External name")
        logs = [call[1] for call in self.second["calls"] if call[0] == "log"]
        self.assertIn("Refresh synced update_proxy settings window failed (ValueError)", logs)
        self.assertFalse(any("synthetic-sensitive-value" in message for message in logs))
        win.destroy()


class AutostartMenuSyncTests(unittest.TestCase):
    def test_existing_timer_refreshes_autostart_with_unchanged_or_missing_config(self):
        for signature in ((1, 1, 1), None):
            with self.subTest(signature=signature):
                namespace = fixtures.ConfigSyncNamespaceRuntimeTests().make_namespace()
                timer = FakeRootTimer()
                namespace["root"] = timer
                namespace["autostart_var"] = fixtures.FakeVar(False)
                current = {"enabled": False}
                namespace["is_autostart_enabled"] = lambda: current["enabled"]

                def schedule(**kwargs):
                    from sms_ui.config_sync_runtime import schedule_config_file_watch_runtime
                    return schedule_config_file_watch_runtime(
                        **kwargs, signature_func=lambda _: signature,
                    )

                first = start_config_file_watch_namespace_runtime(namespace, schedule_runtime=schedule)
                current["enabled"] = True
                timer.run(first)
                self.assertTrue(namespace["autostart_var"].get())
                timer.run(namespace["CONFIG_FILE_WATCH_STATE"].after_id)
                self.assertEqual(len(namespace["calls"]), 1)
                current["enabled"] = False
                timer.run(namespace["CONFIG_FILE_WATCH_STATE"].after_id)
                self.assertFalse(namespace["autostart_var"].get())
                self.assertEqual(len(namespace["calls"]), 2)
                namespace["_CONFIG_SAVE_ACTIVE"] = True
                current["enabled"] = True
                timer.run(namespace["CONFIG_FILE_WATCH_STATE"].after_id)
                self.assertFalse(namespace["autostart_var"].get())

    def test_status_check_failure_retains_checkbox_and_hides_exception_details(self):
        namespace = fixtures.ConfigSyncNamespaceRuntimeTests().make_namespace()
        namespace["autostart_var"] = fixtures.FakeVar(True)
        namespace["is_autostart_enabled"] = lambda: (_ for _ in ()).throw(OSError("sensitive path"))
        self.assertFalse(refresh_autostart_menu_namespace_runtime(namespace))
        self.assertTrue(namespace["autostart_var"].get())
        self.assertEqual(namespace["calls"], [("log", "Refresh autostart menu failed (OSError)")])


if __name__ == "__main__":
    unittest.main()
