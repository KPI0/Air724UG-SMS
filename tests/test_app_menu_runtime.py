import unittest

from sms_ui.app_menu_runtime import build_main_menu_runtime


class FakeVar:
    def __init__(self, value=False):
        self.value = value

    def get(self):
        return self.value


class FakeMenu:
    def __init__(self, parent=None, tearoff=0):
        self.parent = parent
        self.tearoff = tearoff
        self.items = []

    def add_command(self, label, command):
        self.items.append(("command", label, command))

    def add_checkbutton(self, label, variable, command):
        self.items.append(("checkbutton", label, variable, command))

    def add_cascade(self, label, menu):
        self.items.append(("cascade", label, menu))

    def add_separator(self):
        self.items.append(("separator",))

    def index(self, value):
        if value == "end":
            return len(self.items) - 1
        return None


class FakeTkModule:
    Menu = FakeMenu

    @staticmethod
    def BooleanVar(value=False):
        return FakeVar(value)


class FakeRoot:
    def __init__(self):
        self.menu = None

    def config(self, **kwargs):
        self.menu = kwargs.get("menu")


class AppMenuRuntimeTests(unittest.TestCase):
    def test_build_main_menu_runtime_wires_menu_commands_and_state(self):
        root = FakeRoot()
        popup_var = FakeVar(True)
        calls = []
        command_names = [
            "clear_window",
            "open_log_dir",
            "restart_software",
            "send_reset_cmd",
            "cleanup_and_exit",
            "open_serial_setting",
            "open_keywords_setting",
            "open_call_filter_setting",
            "toggle_voice_broadcast",
            "toggle_autostart",
            "toggle_multi_instance",
            "toggle_popup",
            "open_log_cleanup_dialog",
            "open_update_proxy_dialog",
            "open_desktop_shortcut_dialog",
            "open_voice_text_dialog",
            "open_sms_font_dialog",
            "open_cloud_control_window",
            "open_third_push_window",
            "open_serial_debug_window",
            "show_about",
            "check_update_and_prompt",
        ]
        commands = {name: (lambda n=name: calls.append(n)) for name in command_names}

        state = build_main_menu_runtime(
            root,
            FakeTkModule,
            is_autostart_enabled=lambda: True,
            allow_multi_instance=False,
            popup_var=popup_var,
            commands=commands,
        )

        self.assertIs(root.menu, state["menu_bar"])
        self.assertEqual(state["voice_menu_index"], 4)
        self.assertTrue(state["autostart_var"].get())
        self.assertFalse(state["multi_instance_var"].get())

        file_menu = root.menu.items[0][2]
        file_menu.items[0][2]()
        root.menu.items[1][2]()
        settings_menu = root.menu.items[5][2]
        settings_menu.items[0][3]()
        settings_menu.items[2][3]()
        help_menu = root.menu.items[6][2]
        help_menu.items[1][2]()

        self.assertEqual(calls, [
            "clear_window",
            "open_serial_setting",
            "toggle_autostart",
            "toggle_popup",
            "check_update_and_prompt",
        ])


if __name__ == "__main__":
    unittest.main()
