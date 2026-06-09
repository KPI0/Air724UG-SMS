import unittest

from sms_core.serial_namespace_runtime import (
    find_luat_best_port_namespace_runtime,
    open_and_initialize_serial_namespace_runtime,
    resolve_serial_target_port_namespace_runtime,
    scan_com_ports_all_namespace_runtime,
    schedule_delayed_connected_log_namespace_runtime,
    try_manual_rebind_after_error_namespace_runtime,
    try_rebind_manual_port_namespace_runtime,
)


class FakeEvent:
    def __init__(self):
        self.calls = []

    def set(self):
        self.calls.append(("set",))

    def wait(self, **kwargs):
        self.calls.append(("wait", kwargs))

    def clear(self):
        self.calls.append(("clear",))


class FakeNotice:
    def __init__(self):
        self.calls = []

    def reset(self):
        self.calls.append(("reset",))

    def clear(self):
        self.calls.append(("clear",))


class FakeListPorts:
    def __init__(self, ports):
        self._ports = ports

    def comports(self):
        return self._ports


class FakePort:
    def __init__(self, device):
        self.device = device


class FakeSerialModule:
    class Serial:
        def __init__(self, port, baud, timeout=None, write_timeout=None):
            self.port = port
            self.baud = baud
            self.timeout = timeout
            self.write_timeout = write_timeout


class FakeRoot:
    def after(self, *args):
        return ("after", args)


class FakeStatusVar:
    def get(self):
        return "idle"


class SerialNamespaceRuntimeTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        return {
            "calls": calls,
            "MODE": "Manual",
            "PORT": "COM5",
            "BAUD": 115200,
            "RECONNECT_INTERVAL": 3,
            "find_luat_best_port": lambda: ("COM7", "LUAT"),
            "list_ports": FakeListPorts([FakePort("COM7")]),
            "choose_manual_rebind_candidate": lambda *_args, **_kwargs: "candidate",
            "config": "config",
            "safe_save_config": lambda: calls.append(("save_config",)),
            "system_ui": lambda *args: calls.append(("system_ui", args)),
            "set_status": lambda *args: calls.append(("status", args)),
            "serial_wakeup_event": FakeEvent(),
            "_rebind_hint_notice": FakeNotice(),
            "manual_rebind_hint": lambda *args: ("hint", args),
            "is_port_locked_by_other": lambda port: False,
            "auto_connect_ui": lambda message: calls.append(("auto", message)),
            "serial_lock": "serial_lock",
            "serial": FakeSerialModule,
            "serial_obj": None,
            "lock_port_mutex": lambda port: calls.append(("mutex", port)),
            "cloud_imei_query_deadline": 0,
            "serial_error_ui": lambda *args, **kwargs: calls.append(("error", args, kwargs)),
            "_auto_connect_notice": FakeNotice(),
            "_serial_error_notice": FakeNotice(),
            "ui_post": lambda callback: calls.append(("post", callback)),
            "root": FakeRoot(),
            "status_var": FakeStatusVar(),
            "APP_START_MONO": 10.0,
            "START_UI_DELAY": 2.0,
            "rebind_hint_ui": lambda message: calls.append(("hint_ui", message)),
            "try_rebind_manual_port": lambda reason: calls.append(("rebind", reason)) or True,
        }

    def test_try_rebind_manual_port_namespace_runtime_forwards_and_mutates_port(self):
        namespace = self.make_namespace()
        calls = []

        def rebind_runtime(reason, **kwargs):
            calls.append((reason, kwargs))
            kwargs["set_port"]("COM7")
            kwargs["wake_serial"]()
            kwargs["reset_rebind_hint"]()
            return True

        result = try_rebind_manual_port_namespace_runtime(
            namespace,
            "changed",
            rebind_runtime=rebind_runtime,
        )

        self.assertTrue(result)
        self.assertEqual(namespace["PORT"], "COM7")
        self.assertEqual(calls[0][0], "changed")
        self.assertEqual(calls[0][1]["mode"], "Manual")
        self.assertEqual(calls[0][1]["list_ports"]()[0].device, "COM7")
        self.assertEqual(namespace["serial_wakeup_event"].calls, [("set",)])
        self.assertEqual(namespace["_rebind_hint_notice"].calls, [("reset",)])

    def test_resolve_serial_target_port_namespace_runtime_forwards_state(self):
        namespace = self.make_namespace()
        calls = []

        result = resolve_serial_target_port_namespace_runtime(
            namespace,
            resolve_runtime=lambda **kwargs: calls.append(kwargs) or "COM7",
        )

        self.assertEqual(result, "COM7")
        self.assertEqual(calls[0]["mode"], "Manual")
        self.assertEqual(calls[0]["current_port"], "COM5")
        calls[0]["wakeup_wait"](timeout=1)
        calls[0]["wakeup_clear"]()
        self.assertEqual(namespace["serial_wakeup_event"].calls, [
            ("wait", {"timeout": 1}),
            ("clear",),
        ])

    def test_open_and_initialize_serial_namespace_runtime_writes_global_state(self):
        namespace = self.make_namespace()
        calls = []

        def open_runtime(**kwargs):
            serial_obj = kwargs["open_serial"]("COM9", 9600)
            calls.append((serial_obj.port, serial_obj.baud, serial_obj.timeout, serial_obj.write_timeout))
            kwargs["set_serial_obj"](serial_obj)
            kwargs["set_port"]("COM9")
            kwargs["set_cloud_imei_query_deadline"](16.0)
            return serial_obj

        result = open_and_initialize_serial_namespace_runtime(
            namespace,
            "COM9",
            open_runtime=open_runtime,
        )

        self.assertIs(result, namespace["serial_obj"])
        self.assertEqual(namespace["PORT"], "COM9")
        self.assertEqual(namespace["cloud_imei_query_deadline"], 16.0)
        self.assertEqual(calls, [("COM9", 9600, 0.3, 0.5)])

    def test_schedule_delayed_connected_log_namespace_runtime_forwards_callbacks(self):
        namespace = self.make_namespace()
        calls = []

        result = schedule_delayed_connected_log_namespace_runtime(
            namespace,
            "COM5",
            115200,
            delay=1,
            start_runtime=lambda *args, **kwargs: calls.append((args, kwargs)) or "thread",
        )

        self.assertEqual(result, "thread")
        self.assertEqual(calls[0][0], ("COM5", 115200))
        self.assertEqual(calls[0][1]["delay"], 1)
        calls[0][1]["reset_auto_connect_state"]()
        calls[0][1]["clear_serial_error_repeat_state"]()
        self.assertEqual(namespace["_auto_connect_notice"].calls, [("reset",)])
        self.assertEqual(namespace["_serial_error_notice"].calls, [("clear",)])
        self.assertEqual(calls[0][1]["get_status"](), "idle")

    def test_try_manual_rebind_after_error_namespace_runtime_checks_mode_and_error(self):
        namespace = self.make_namespace()

        self.assertFalse(
            try_manual_rebind_after_error_namespace_runtime(
                namespace,
                RuntimeError("parse"),
                hint_message="hint",
                is_gone_error=lambda _error: False,
            )
        )
        self.assertTrue(
            try_manual_rebind_after_error_namespace_runtime(
                namespace,
                RuntimeError("gone"),
                hint_message="hint",
                is_gone_error=lambda _error: True,
            )
        )
        self.assertIn(("hint_ui", "hint"), namespace["calls"])
        self.assertIn(("rebind", "端口号变化或设备插拔"), namespace["calls"])

    def test_scan_and_find_luat_port_namespace_runtime_read_port_state(self):
        namespace = self.make_namespace()
        calls = []

        self.assertEqual(scan_com_ports_all_namespace_runtime(namespace), ["COM7"])
        result = find_luat_best_port_namespace_runtime(
            namespace,
            choose_port=lambda ports, **kwargs: calls.append((ports, kwargs)) or ("COM7", "LUAT"),
        )

        self.assertEqual(result, ("COM7", "LUAT"))
        self.assertEqual(calls[0][0][0].device, "COM7")
        self.assertEqual(calls[0][1]["remembered_port"], "COM5")
        self.assertIs(calls[0][1]["is_locked"], namespace["is_port_locked_by_other"])


if __name__ == "__main__":
    unittest.main()
