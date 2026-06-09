import unittest

from sms_ui.app_restart_runtime import restart_software_app_runtime


class FakeMessageBox:
    def __init__(self, answer=True):
        self.answer = answer
        self.ask_calls = []
        self.error_calls = []

    def askyesno(self, title, message, parent=None):
        self.ask_calls.append((title, message, parent))
        return self.answer

    def showerror(self, title, message, parent=None):
        self.error_calls.append((title, message, parent))


class AppRestartRuntimeTests(unittest.TestCase):
    def test_restart_software_app_runtime_wires_runtime_dependencies(self):
        calls = []
        messagebox = FakeMessageBox(answer=True)

        def restart_runtime(**kwargs):
            calls.append(kwargs)
            self.assertTrue(kwargs["confirm_restart"]())
            kwargs["show_launch_error"](RuntimeError("boom"))
            return "result"

        result = restart_software_app_runtime(
            root="root",
            messagebox=messagebox,
            is_exiting=False,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            autostart_flag="--autostart",
            restart_helper_flag="--restart-helper",
            log_error=lambda message: calls.append(("log", message)),
            system_ui=lambda *args: calls.append(("system", args)),
            stop_tray_icon=lambda **kwargs: calls.append(("tray", kwargs)),
            safe_set_events=lambda *events: calls.append(("events", events)),
            stop_events=("third", "serial"),
            stop_cloud_control=lambda **kwargs: calls.append(("cloud", kwargs)),
            safe_close_serial=lambda: calls.append(("close",)),
            app_mutex="mutex",
            release_mutex=lambda mutex: calls.append(("release", mutex)),
            flush_log_queue=lambda queue: calls.append(("flush", queue)),
            file_log_queue="queue",
            exit_process=lambda code: calls.append(("exit", code)),
            argv=["sms.pyw", "--debug"],
            current_pid=123,
            restart_runtime=restart_runtime,
        )

        self.assertEqual(result, "result")
        kwargs = calls[0]
        self.assertEqual(kwargs["argv"], ["sms.pyw", "--debug"])
        self.assertEqual(kwargs["current_pid"], 123)
        self.assertEqual(kwargs["stop_events"], ("third", "serial"))
        self.assertEqual(kwargs["file_log_queue"], "queue")
        self.assertEqual(messagebox.ask_calls[0][2], "root")
        self.assertEqual(messagebox.error_calls[0][2], "root")


if __name__ == "__main__":
    unittest.main()
