import unittest
from unittest.mock import patch

from sms_core.third_push import THIRD_PUSH_CHANNELS, THIRD_PUSH_SETTINGS_KEYS
from sms_ui.third_push_window import (
    confirm_close_with_unsaved_changes,
    open_third_push_window_runtime,
)
from sms_ui.third_push_window_form import ThirdPushFormController


class FakeVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


def make_form(state, calls=None):
    calls = calls if calls is not None else []
    form = ThirdPushFormController.__new__(ThirdPushFormController)
    form.win = "window"
    form.state_provider = lambda: state
    form.on_option_changed = (
        lambda option, value, values, win: calls.append((option, value, values, win)) or True
    )
    form.on_dirty_changed = lambda dirty: calls.append(("dirty", dirty))
    form.enabled_var = FakeVar(1 if state["enabled"] else 0)
    form.sms_push_var = FakeVar(1 if state["sms_enabled"] else 0)
    form.call_push_var = FakeVar(1 if state["call_enabled"] else 0)
    form.channel_vars = {
        channel: FakeVar(channel in state["channels"])
        for channel, _label in THIRD_PUSH_CHANNELS
    }
    form.entry_vars = {
        key: FakeVar(state["settings"].get(key, ""))
        for key in THIRD_PUSH_SETTINGS_KEYS
    }
    form.custom_body_text = None
    form._syncing = False
    form._dirty = False
    form._saved_snapshot = form._snapshot_from_state(state)
    form._set_custom_body_text = lambda _value: None
    return form


class ThirdPushWindowFormTests(unittest.TestCase):
    def test_parameter_change_marks_form_dirty_and_mark_saved_clears_it(self):
        calls = []
        state = {
            "enabled": False,
            "sms_enabled": True,
            "call_enabled": True,
            "channels": ["wecom"],
            "settings": {"wecom_webhook": "https://old.example"},
        }
        form = make_form(state, calls)

        form.entry_vars["wecom_webhook"].set("https://new.example")
        form._notify_dirty()

        self.assertTrue(form.is_dirty())
        self.assertIn(("dirty", True), calls)

        form.mark_all_saved()

        self.assertFalse(form.is_dirty())
        self.assertIn(("dirty", False), calls)

    def test_sms_toggle_saves_only_option_and_keeps_parameter_draft_dirty(self):
        calls = []
        state = {
            "enabled": True,
            "sms_enabled": True,
            "call_enabled": True,
            "channels": ["wecom"],
            "settings": {"wecom_webhook": "https://old.example"},
        }
        form = make_form(state, calls)
        form.entry_vars["wecom_webhook"].set("https://draft.example")
        form.sms_push_var.set(0)

        form._option_toggled("sms_enabled", form.sms_push_var)

        self.assertIn(("sms_enabled", False, None, "window"), calls)
        self.assertTrue(form.is_dirty())
        self.assertFalse(form._saved_snapshot["sms_enabled"])

    def test_enabling_validates_and_saves_complete_form(self):
        calls = []
        state = {
            "enabled": False,
            "sms_enabled": True,
            "call_enabled": True,
            "channels": ["wecom"],
            "settings": {"wecom_webhook": "https://example.test"},
        }
        form = make_form(state, calls)
        values = (True, True, True, ["wecom"], {"wecom_webhook": "https://example.test"})
        form.enabled_var.set(1)
        form.collect = lambda validate=True: values

        form._option_toggled("enabled", form.enabled_var)

        self.assertIn(("enabled", True, values, "window"), calls)
        self.assertFalse(form.is_dirty())

    def test_failed_immediate_save_restores_latest_runtime_state(self):
        calls = []
        state = {
            "enabled": True,
            "sms_enabled": True,
            "call_enabled": False,
            "channels": ["wecom"],
            "settings": {"wecom_webhook": "https://example.test"},
        }
        form = make_form(state, calls)
        form.on_option_changed = lambda *_args: False
        form.call_push_var.set(1)

        form._option_toggled("call_enabled", form.call_push_var)

        self.assertFalse(bool(form.call_push_var.get()))
        self.assertFalse(form.is_dirty())

    def test_reopen_sync_preserves_dirty_draft_but_refreshes_clean_form(self):
        state = {
            "enabled": True,
            "sms_enabled": True,
            "call_enabled": True,
            "channels": ["wecom"],
            "settings": {"wecom_webhook": "https://old.example"},
        }
        form = make_form(state)
        form.entry_vars["wecom_webhook"].set("https://draft.example")
        form._notify_dirty()
        state["settings"] = {"wecom_webhook": "https://external.example"}

        self.assertFalse(form.sync_from_state_if_clean())
        self.assertEqual(form.entry_vars["wecom_webhook"].get(), "https://draft.example")
        self.assertTrue(form.is_dirty())

        form.mark_all_saved()
        state["settings"] = {"wecom_webhook": "https://latest.example"}

        self.assertTrue(form.sync_from_state_if_clean())
        self.assertEqual(form.entry_vars["wecom_webhook"].get(), "https://latest.example")
        self.assertFalse(form.is_dirty())

    def test_close_confirmation_only_runs_for_dirty_form(self):
        calls = []

        class FakeForm:
            def __init__(self, dirty):
                self.dirty = dirty

            def store_custom_body_text(self):
                calls.append("store")

            def is_dirty(self):
                return self.dirty

        self.assertTrue(confirm_close_with_unsaved_changes(
            FakeForm(False),
            "window",
            confirm=lambda *_args, **_kwargs: self.fail("clean form should not confirm"),
        ))
        self.assertFalse(confirm_close_with_unsaved_changes(
            FakeForm(True),
            "window",
            confirm=lambda *args, **kwargs: calls.append((args, kwargs)) or False,
        ))
        self.assertEqual(calls.count("store"), 2)

    def test_runtime_immediate_options_save_and_regular_save_is_non_modal(self):
        calls = []

        def fake_dialog(
            _parent,
            _state_provider,
            on_save,
            _on_test,
            _on_close,
            _center_window,
            **kwargs,
        ):
            option_changed = kwargs["on_option_changed"]
            values = (
                True,
                True,
                False,
                ["wecom"],
                {"wecom_webhook": "https://example.test"},
            )
            calls.append(("enable_result", option_changed("enabled", True, values, "window")))
            calls.append(("sms_result", option_changed("sms_enabled", False, None, "window")))
            calls.append(("save_result", on_save(values, "window")))
            return "window"

        with (
            patch("sms_ui.third_push_window.open_third_push_window_dialog", side_effect=fake_dialog),
            patch("sms_ui.third_push_window.messagebox.showinfo") as show_info,
        ):
            result = open_third_push_window_runtime(
                "root",
                None,
                lambda: {
                    "enabled": False,
                    "sms_enabled": True,
                    "call_enabled": False,
                    "channels": ["wecom"],
                    "settings": {"wecom_webhook": "https://example.test"},
                },
                lambda: None,
                lambda **kwargs: calls.append(("save", kwargs)) or object(),
                lambda *_args, **_kwargs: True,
                lambda *args: calls.append(("system", args)),
                lambda *_args: False,
                lambda win: calls.append(("window", win)),
                "center",
            )

        self.assertEqual(result, "window")
        self.assertIn(("save", {
            "enabled": True,
            "sms_enabled": True,
            "call_enabled": False,
            "notify_type": ["wecom"],
            "settings": {"wecom_webhook": "https://example.test"},
        }), calls)
        self.assertIn(("save", {"sms_enabled": False}), calls)
        self.assertIn(("enable_result", True), calls)
        self.assertIn(("sms_result", True), calls)
        self.assertIn(("save_result", True), calls)
        show_info.assert_not_called()


if __name__ == "__main__":
    unittest.main()
