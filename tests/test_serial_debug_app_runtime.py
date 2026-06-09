import unittest

from sms_ui.serial_debug_app_runtime import open_serial_debug_window_runtime


class SerialDebugAppRuntimeTests(unittest.TestCase):
    def test_open_serial_debug_window_runtime_adapts_state_callbacks(self):
        state = {
            "window_refs": ("old_win", "old_text"),
            "debug_enabled": True,
            "drop_count": 3,
            "current_dial_num": "",
        }
        updates = []
        opened = {}

        def get_state(name):
            return state[name]

        def set_state(name, *values):
            updates.append((name, values))
            if name == "window_refs":
                state[name] = values
            elif name == "clear_window_refs":
                state["window_refs"] = (None, None)
            else:
                state[name] = values[0] if values else None

        def open_dialog(
            parent,
            current_window,
            current_text,
            debug_enabled,
            get_drop_count,
            serial_queue,
            serial_lock,
            get_serial_obj,
            push_serial_debug,
            port_ui,
            set_status,
            format_connected_status,
            get_port,
            set_current_dial_num,
            set_debug_enabled,
            set_drop_count,
            clear_window_refs,
            center_window,
        ):
            opened.update(
                parent=parent,
                current_window=current_window,
                current_text=current_text,
                debug_enabled=debug_enabled,
                drop_count=get_drop_count(),
                serial_queue=serial_queue,
                serial_lock=serial_lock,
                serial_obj=get_serial_obj(),
                port=get_port(),
                centered=center_window,
            )
            push_serial_debug("raw")
            port_ui("port")
            set_status("status")
            self.assertEqual(format_connected_status("COM5"), "connected:COM5")
            set_current_dial_num("10086")
            set_debug_enabled(False)
            set_drop_count(7)
            clear_window_refs()
            return "new_win", "new_text"

        calls = []
        window, text = open_serial_debug_window_runtime(
            "root",
            get_state=get_state,
            set_state=set_state,
            serial_queue="queue",
            serial_lock="lock",
            get_serial_obj=lambda: "serial",
            push_serial_debug=lambda value: calls.append(("debug", value)),
            port_ui=lambda value: calls.append(("port_ui", value)),
            set_status=lambda value: calls.append(("status", value)),
            format_connected_status=lambda port: f"connected:{port}",
            get_port=lambda: "COM5",
            center_window="center",
            open_dialog=open_dialog,
        )

        self.assertEqual((window, text), ("new_win", "new_text"))
        self.assertEqual(opened["current_window"], "old_win")
        self.assertEqual(opened["current_text"], "old_text")
        self.assertTrue(opened["debug_enabled"])
        self.assertEqual(opened["drop_count"], 3)
        self.assertEqual(opened["serial_obj"], "serial")
        self.assertEqual(opened["port"], "COM5")
        self.assertEqual(calls, [("debug", "raw"), ("port_ui", "port"), ("status", "status")])
        self.assertIn(("current_dial_num", ("10086",)), updates)
        self.assertIn(("debug_enabled", (False,)), updates)
        self.assertIn(("drop_count", (7,)), updates)
        self.assertIn(("clear_window_refs", ()), updates)
        self.assertEqual(state["window_refs"], ("new_win", "new_text"))


if __name__ == "__main__":
    unittest.main()
