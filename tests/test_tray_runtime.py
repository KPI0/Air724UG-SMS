import unittest

from sms_ui.tray_runtime import create_tray_icon_runtime, load_tray_image, stop_tray_icon_runtime


class FakeImageModule:
    def __init__(self, open_result=None, open_error=None, new_result="fallback", new_error=None):
        self.open_result = open_result
        self.open_error = open_error
        self.new_result = new_result
        self.new_error = new_error
        self.open_calls = []
        self.new_calls = []

    def open(self, path):
        self.open_calls.append(path)
        if self.open_error:
            raise self.open_error
        return self.open_result

    def new(self, mode, size, color=None):
        self.new_calls.append((mode, size, color))
        if self.new_error:
            raise self.new_error
        return self.new_result


class FakePystray:
    class MenuItem:
        def __init__(self, label, action, default=False):
            self.label = label
            self.action = action
            self.default = default

    class Menu:
        def __init__(self, *items):
            self.items = items

    class Icon:
        def __init__(self, name, image, title, menu):
            self.name = name
            self.image = image
            self.title = title
            self.menu = menu
            self.run_calls = 0

        def run(self):
            self.run_calls += 1


class FakeIcon:
    def __init__(self, fail_visible=False, fail_stop=False):
        self._visible = True
        self.fail_visible = fail_visible
        self.fail_stop = fail_stop
        self.stop_calls = 0

    @property
    def visible(self):
        return self._visible

    @visible.setter
    def visible(self, value):
        if self.fail_visible:
            raise RuntimeError("visible failed")
        self._visible = value

    def stop(self):
        self.stop_calls += 1
        if self.fail_stop:
            raise RuntimeError("stop failed")


class TrayRuntimeTests(unittest.TestCase):
    def test_load_tray_image_prefers_icon_path(self):
        image_module = FakeImageModule(open_result="icon")

        self.assertEqual(load_tray_image("icon.ico", image_module=image_module), "icon")
        self.assertEqual(image_module.open_calls, ["icon.ico"])
        self.assertEqual(image_module.new_calls, [])

    def test_load_tray_image_uses_fallback_image(self):
        image_module = FakeImageModule(open_error=RuntimeError("missing"), new_result="fallback")

        self.assertEqual(load_tray_image("missing.ico", image_module=image_module), "fallback")
        self.assertEqual(image_module.new_calls, [("RGB", (16, 16), (200, 30, 30))])

    def test_create_tray_icon_runtime_wires_menu_callbacks(self):
        calls = []
        icons = []

        icon = create_tray_icon_runtime(
            icon_path="icon.ico",
            title="title",
            show_window=lambda: calls.append("show"),
            hide_window=lambda: calls.append("hide"),
            cleanup_and_exit=lambda: calls.append("exit"),
            set_tray_icon=icons.append,
            pystray_module=FakePystray,
            image_loader=lambda path: f"image:{path}",
        )

        self.assertIs(icon, icons[0])
        self.assertEqual(icon.name, "sms_tray")
        self.assertEqual(icon.image, "image:icon.ico")
        self.assertEqual(icon.title, "title")
        self.assertEqual(icon.run_calls, 1)

        for item in icon.menu.items:
            item.action()
        self.assertEqual(calls, ["show", "hide", "exit"])

    def test_create_tray_icon_runtime_skips_when_image_missing(self):
        calls = []

        result = create_tray_icon_runtime(
            icon_path="icon.ico",
            title="title",
            show_window=lambda: None,
            hide_window=lambda: None,
            cleanup_and_exit=lambda: None,
            set_tray_icon=calls.append,
            pystray_module=FakePystray,
            image_loader=lambda path: None,
        )

        self.assertIsNone(result)
        self.assertEqual(calls, [])

    def test_stop_tray_icon_runtime_clears_and_stops_icon(self):
        calls = []
        icon = FakeIcon()

        result = stop_tray_icon_runtime(
            tray_icon=icon,
            clear_tray_icon=lambda: calls.append("clear"),
            wait_after=0.25,
            sleep=lambda seconds: calls.append(("sleep", seconds)),
        )

        self.assertTrue(result)
        self.assertEqual(calls, ["clear", ("sleep", 0.25)])
        self.assertFalse(icon.visible)
        self.assertEqual(icon.stop_calls, 1)

    def test_stop_tray_icon_runtime_tolerates_icon_errors(self):
        calls = []
        icon = FakeIcon(fail_visible=True, fail_stop=True)

        result = stop_tray_icon_runtime(
            tray_icon=icon,
            clear_tray_icon=lambda: calls.append("clear"),
            wait_after=0,
        )

        self.assertTrue(result)
        self.assertEqual(calls, ["clear"])
        self.assertEqual(icon.stop_calls, 1)


if __name__ == "__main__":
    unittest.main()
