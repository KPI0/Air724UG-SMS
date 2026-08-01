import unittest
from datetime import datetime
from types import SimpleNamespace

from sms_ui.missed_call_popup_runtime import (
    show_missed_call_popup_app_runtime,
    show_missed_call_popup_runtime,
)


class FakePopup:
    def __init__(self):
        self.exists = True
        self.destroyed = False
        self.updates = []

    def winfo_exists(self):
        return self.exists

    def destroy(self):
        self.destroyed = True
        self.exists = False

    def missed_call_popup_update(self, caller_num, started_at, count):
        self.updates.append((caller_num, started_at, count))


class MissedCallPopupRuntimeTests(unittest.TestCase):
    def test_singleton_popup_updates_latest_call_and_count(self):
        state = [None]
        popup = FakePopup()
        opened = []
        started_at = datetime(2026, 8, 1, 10, 30, 0)

        def open_popup(parent, caller_num, call_time, center_window, on_close):
            opened.append((parent, caller_num, call_time, on_close))
            return popup

        first = SimpleNamespace(caller_num="10086", started_at=started_at)
        result = show_missed_call_popup_runtime(
            first,
            parent="root",
            current_popup=state[0],
            set_popup=lambda value: state.__setitem__(0, value),
            center_window="center",
            show_window=lambda: None,
            open_popup=open_popup,
        )

        self.assertEqual(result, "shown")
        self.assertIs(state[0], popup)
        self.assertEqual(popup.missed_call_popup_count, 1)

        second_time = datetime(2026, 8, 1, 10, 35, 0)
        second = SimpleNamespace(caller_num="10010", started_at=second_time)
        result = show_missed_call_popup_runtime(
            second,
            parent="root",
            current_popup=state[0],
            set_popup=lambda value: state.__setitem__(0, value),
            center_window="center",
            show_window=lambda: None,
            open_popup=open_popup,
        )

        self.assertEqual(result, "updated")
        self.assertEqual(popup.updates, [("10010", second_time, 2)])
        self.assertEqual(popup.missed_call_popup_count, 2)
        self.assertEqual(len(opened), 1)

    def test_close_resets_singleton_and_restores_main_window(self):
        state = [None]
        popup = FakePopup()
        shown = []
        close_callback = []

        def open_popup(parent, caller_num, call_time, center_window, on_close):
            close_callback.append(on_close)
            return popup

        show_missed_call_popup_runtime(
            SimpleNamespace(caller_num="10086", started_at=None),
            parent="root",
            current_popup=None,
            set_popup=lambda value: state.__setitem__(0, value),
            center_window="center",
            show_window=lambda: shown.append("main"),
            open_popup=open_popup,
        )
        close_callback[0]()

        self.assertTrue(popup.destroyed)
        self.assertIsNone(state[0])
        self.assertEqual(shown, ["main"])

    def test_app_runtime_posts_to_ui_thread(self):
        calls = []
        missed_call = object()

        result = show_missed_call_popup_app_runtime(
            missed_call=missed_call,
            parent="root",
            get_popup=lambda: "popup",
            set_popup=lambda value: None,
            center_window="center",
            show_window="show",
            run_on_ui_thread=lambda callback, ui_post: calls.append(("run", ui_post)) or callback(),
            ui_post="post",
            show_runtime=lambda *args, **kwargs: calls.append((args, kwargs)) or "shown",
        )

        self.assertEqual(result, "shown")
        self.assertEqual(calls[0], ("run", "post"))
        self.assertIs(calls[1][0][0], missed_call)
        self.assertEqual(calls[1][1]["current_popup"], "popup")

    def test_app_runtime_rechecks_setting_on_ui_thread(self):
        queued = []
        enabled = [True]
        shown = []

        result = show_missed_call_popup_app_runtime(
            missed_call=object(),
            parent="root",
            get_popup=lambda: None,
            set_popup=lambda value: None,
            center_window="center",
            show_window="show",
            run_on_ui_thread=lambda callback, _ui_post: queued.append(callback),
            ui_post="post",
            is_enabled=lambda: enabled[0],
            show_runtime=lambda *args, **kwargs: shown.append((args, kwargs)),
        )

        self.assertIsNone(result)
        enabled[0] = False
        self.assertEqual(queued[0](), "disabled")
        self.assertEqual(shown, [])


if __name__ == "__main__":
    unittest.main()
