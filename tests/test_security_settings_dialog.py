import unittest
from unittest.mock import patch

from sms_ui import security_settings_dialog
from sms_ui.security_settings_dialog import open_security_settings_dialog


class FakeVar:
    def __init__(self, value=None):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeWidget:
    created = []

    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs
        self.destroyed = False
        self.grabbed = False
        self.transient_parent = None
        self.pack_options = None
        self.grid_options = None
        self.grid_columns = []
        self.geometry_values = []
        self.protocols = {}
        self.deiconify_count = 0
        self.lift_count = 0
        self.focus_count = 0
        FakeWidget.created.append(self)

    def pack(self, **kwargs):
        self.pack_options = kwargs
        return self

    def grid(self, **kwargs):
        self.grid_options = kwargs
        return self

    def grid_columnconfigure(self, column, **kwargs):
        self.grid_columns.append((column, kwargs))
        return None

    def withdraw(self):
        return None

    def title(self, _value):
        return None

    def geometry(self, value):
        self.geometry_values.append(value)
        return None

    def resizable(self, *_args):
        return None

    def transient(self, parent):
        self.transient_parent = parent
        return None

    def grab_set(self):
        self.grabbed = True
        return None

    def update_idletasks(self):
        return None

    def deiconify(self):
        self.deiconify_count += 1
        return None

    def lift(self):
        self.lift_count += 1
        return None

    def focus_force(self):
        self.focus_count += 1
        return None

    def winfo_exists(self):
        return not self.destroyed

    def protocol(self, name, callback):
        self.protocols[name] = callback
        return None

    def bind(self, *_args):
        return None

    def destroy(self):
        self.destroyed = True


class FakeTk:
    BOTH = "both"
    X = "x"
    LEFT = "left"
    RIGHT = "right"
    Toplevel = FakeWidget
    Frame = FakeWidget
    LabelFrame = FakeWidget
    Label = FakeWidget
    Checkbutton = FakeWidget
    Button = FakeWidget
    BooleanVar = FakeVar
    StringVar = FakeVar


class FakeTtk:
    Frame = FakeWidget
    LabelFrame = FakeWidget
    Label = FakeWidget
    Checkbutton = FakeWidget
    Button = FakeWidget


class SecuritySettingsDialogTests(unittest.TestCase):
    def setUp(self):
        FakeWidget.created = []
        security_settings_dialog._active_security_settings_refs = None

    def test_checkbox_toggle_applies_immediately_without_save_button(self):
        changes = []

        with patch("sms_ui.security_settings_dialog.tk", FakeTk), patch(
            "sms_ui.security_settings_dialog.ttk", FakeTtk
        ), patch(
            "sms_ui.security_settings_dialog.messagebox.askyesno",
            return_value=True,
        ):
            refs = open_security_settings_dialog(
                "root",
                {},
                lambda permissions: changes.append(dict(permissions)) or True,
                lambda *_args: None,
            )
            refs["permission_vars"]["sms"].set(True)
            self.assertTrue(refs["toggle"]("sms"))

        self.assertEqual(len(changes), 1)
        self.assertTrue(changes[0]["sms"])
        self.assertFalse(changes[0]["call"])
        button_labels = [
            widget.kwargs.get("text")
            for widget in FakeWidget.created
            if widget.kwargs.get("text")
        ]
        self.assertIn("关闭", button_labels)
        self.assertIn("开启全部", button_labels)
        self.assertIn("关闭全部", button_labels)
        self.assertNotIn("保存", button_labels)

        close_button = next(
            widget for widget in FakeWidget.created
            if widget.kwargs.get("text") == "关闭"
        )
        self.assertEqual(close_button.grid_options["column"], 0)
        footer = close_button.args[0]
        self.assertIn(
            (0, {"weight": 1}),
            footer.grid_columns,
        )
        self.assertNotIn(
            "勾选或取消后立即保存并生效",
            button_labels,
        )
        status_label = next(
            widget for widget in FakeWidget.created
            if widget.kwargs.get("textvariable") is refs["status_var"]
        )
        self.assertEqual(status_label.pack_options["side"], FakeTk.RIGHT)
        self.assertFalse(FakeWidget.created[0].grabbed)
        self.assertIsNone(FakeWidget.created[0].transient_parent)
        self.assertEqual(FakeWidget.created[0].geometry_values[-1], "620x285")

    def test_enable_all_and_disable_all_apply_immediately(self):
        changes = []

        with patch("sms_ui.security_settings_dialog.tk", FakeTk), patch(
            "sms_ui.security_settings_dialog.ttk", FakeTtk
        ), patch(
            "sms_ui.security_settings_dialog.messagebox.askyesno",
            return_value=True,
        ):
            refs = open_security_settings_dialog(
                "root",
                {},
                lambda permissions: changes.append(dict(permissions)) or True,
                lambda *_args: None,
            )
            self.assertTrue(refs["set_all"](True))
            self.assertTrue(refs["set_all"](False))

        self.assertTrue(all(changes[0].values()))
        self.assertTrue(all(not value for value in changes[1].values()))

    def test_repeated_open_focuses_existing_window_and_close_allows_reopen(self):
        with patch("sms_ui.security_settings_dialog.tk", FakeTk), patch(
            "sms_ui.security_settings_dialog.ttk", FakeTtk
        ):
            first = open_security_settings_dialog(
                "root",
                {},
                lambda _permissions: True,
                lambda *_args: None,
            )
            created_count = len(FakeWidget.created)
            second = open_security_settings_dialog(
                "root",
                {},
                lambda _permissions: True,
                lambda *_args: None,
            )

            self.assertIs(second, first)
            self.assertEqual(len(FakeWidget.created), created_count)
            self.assertEqual(first["window"].lift_count, 2)
            self.assertEqual(first["window"].focus_count, 2)

            first["close"]()
            third = open_security_settings_dialog(
                "root",
                {},
                lambda _permissions: True,
                lambda *_args: None,
            )

        self.assertIsNot(third, first)
        self.assertIsNot(third["window"], first["window"])

    def test_cancel_or_save_failure_restores_previous_checkbox_state(self):
        changes = []

        with patch("sms_ui.security_settings_dialog.tk", FakeTk), patch(
            "sms_ui.security_settings_dialog.ttk", FakeTtk
        ), patch(
            "sms_ui.security_settings_dialog.messagebox.askyesno",
            return_value=False,
        ):
            refs = open_security_settings_dialog(
                "root",
                {},
                lambda permissions: changes.append(dict(permissions)) or True,
                lambda *_args: None,
            )
            refs["permission_vars"]["call"].set(True)
            self.assertFalse(refs["toggle"]("call"))

        self.assertFalse(refs["permission_vars"]["call"].get())
        self.assertEqual(changes, [])
        refs["close"]()

        with patch("sms_ui.security_settings_dialog.tk", FakeTk), patch(
            "sms_ui.security_settings_dialog.ttk", FakeTtk
        ), patch(
            "sms_ui.security_settings_dialog.messagebox.askyesno",
            return_value=True,
        ), patch("sms_ui.security_settings_dialog.messagebox.showerror") as show_error:
            refs = open_security_settings_dialog(
                "root",
                {},
                lambda _permissions: False,
                lambda *_args: None,
            )
            refs["permission_vars"]["pin"].set(True)
            self.assertFalse(refs["toggle"]("pin"))

        self.assertFalse(refs["permission_vars"]["pin"].get())
        show_error.assert_called_once()


if __name__ == "__main__":
    unittest.main()
