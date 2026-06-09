import unittest
import queue

from sms_ui.serial_debug_namespace_runtime import (
    get_serial_debug_state_namespace_runtime,
    open_serial_debug_window_namespace_runtime,
    push_serial_debug_namespace_runtime,
    set_serial_debug_state_namespace_runtime,
)


class SerialDebugNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        return {
            "root": "root",
            "serial_debug_win": "old_win",
            "serial_debug_text": "old_text",
            "SERIAL_DEBUG_ENABLED": True,
            "serial_debug_drop_count": 4,
            "current_dial_num": "",
            "serial_debug_queue": "queue",
            "serial_lock": "lock",
            "serial_obj": "serial",
            "PORT": "COM5",
            "_push_serial_debug": lambda value: ("debug", value),
            "port_ui": lambda value: ("port", value),
            "set_status": lambda value: ("status", value),
            "format_connected_status": lambda port: f"connected:{port}",
            "center_window": "center",
        }

    def test_get_serial_debug_state_namespace_runtime_reads_known_values(self):
        namespace = self.base_namespace()

        self.assertEqual(get_serial_debug_state_namespace_runtime(namespace, "window_refs"), ("old_win", "old_text"))
        self.assertTrue(get_serial_debug_state_namespace_runtime(namespace, "debug_enabled"))
        self.assertEqual(get_serial_debug_state_namespace_runtime(namespace, "drop_count"), 4)
        with self.assertRaises(KeyError):
            get_serial_debug_state_namespace_runtime(namespace, "missing")

    def test_set_serial_debug_state_namespace_runtime_updates_known_values(self):
        namespace = self.base_namespace()

        set_serial_debug_state_namespace_runtime(namespace, "window_refs", "new_win", "new_text")
        set_serial_debug_state_namespace_runtime(namespace, "debug_enabled", 0)
        set_serial_debug_state_namespace_runtime(namespace, "drop_count", "8")
        set_serial_debug_state_namespace_runtime(namespace, "current_dial_num", "10086")

        self.assertEqual(namespace["serial_debug_win"], "new_win")
        self.assertEqual(namespace["serial_debug_text"], "new_text")
        self.assertFalse(namespace["SERIAL_DEBUG_ENABLED"])
        self.assertEqual(namespace["serial_debug_drop_count"], 8)
        self.assertEqual(namespace["current_dial_num"], "10086")

        set_serial_debug_state_namespace_runtime(namespace, "clear_window_refs")
        self.assertIsNone(namespace["serial_debug_win"])
        self.assertIsNone(namespace["serial_debug_text"])
        with self.assertRaises(KeyError):
            set_serial_debug_state_namespace_runtime(namespace, "missing")

    def test_open_serial_debug_window_namespace_runtime_forwards_callbacks(self):
        namespace = self.base_namespace()
        calls = []

        result = open_serial_debug_window_namespace_runtime(
            namespace,
            open_window_runtime=lambda parent, **kwargs: calls.append((parent, kwargs)) or ("win", "text"),
        )

        self.assertEqual(result, ("win", "text"))
        parent, forwarded = calls[0]
        self.assertEqual(parent, "root")
        self.assertEqual(forwarded["serial_queue"], "queue")
        self.assertEqual(forwarded["serial_lock"], "lock")
        self.assertEqual(forwarded["get_serial_obj"](), "serial")
        self.assertEqual(forwarded["get_port"](), "COM5")
        self.assertEqual(forwarded["center_window"], "center")
        self.assertEqual(forwarded["get_state"]("window_refs"), ("old_win", "old_text"))
        forwarded["set_state"]("drop_count", 9)
        self.assertEqual(namespace["serial_debug_drop_count"], 9)

    def test_push_serial_debug_namespace_runtime_filters_and_queues_lines(self):
        namespace = self.base_namespace()
        namespace["serial_debug_queue"] = queue.Queue(maxsize=1)

        self.assertIsNone(push_serial_debug_namespace_runtime(namespace, None))
        self.assertIsNone(push_serial_debug_namespace_runtime(namespace, "   "))
        self.assertEqual(push_serial_debug_namespace_runtime(namespace, "AT"), "queued")
        self.assertEqual(namespace["serial_debug_queue"].get_nowait(), "AT")

        namespace["SERIAL_DEBUG_ENABLED"] = False
        self.assertIsNone(push_serial_debug_namespace_runtime(namespace, "ATI"))

    def test_push_serial_debug_namespace_runtime_counts_full_queue_drops(self):
        namespace = self.base_namespace()
        namespace["serial_debug_queue"] = queue.Queue(maxsize=1)
        namespace["serial_debug_queue"].put_nowait("old")

        self.assertEqual(push_serial_debug_namespace_runtime(namespace, "new"), "full")

        self.assertEqual(namespace["serial_debug_drop_count"], 5)


if __name__ == "__main__":
    unittest.main()
