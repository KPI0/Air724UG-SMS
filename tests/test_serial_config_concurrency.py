import threading
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from sms_core.config_runtime import load_config_snapshot, safe_save_config_runtime
from sms_core.serial_namespace_runtime import try_rebind_manual_port_namespace_runtime
from sms_ui.app_infrastructure_namespace_runtime import safe_save_config_namespace_runtime
from sms_ui.serial_settings_runtime import apply_serial_setting_runtime
from tests import test_client_quality_regressions as quality


class SerialConfigConcurrencyTests(unittest.TestCase):
    def setUp(self):
        self.fixture = quality.TkQualityTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.doCleanups)
        self.root = self.fixture.root

    def tearDown(self):
        self.root = None

    def namespace(self):
        ns = self.fixture.namespace()
        ns["config"]["serial"] = {
            "mode": "Manual", "port": "COM901", "baud": "115200", "unknown": "keep",
        }
        self.assertTrue(safe_save_config_namespace_runtime(ns))
        ns.update(
            MODE="Manual", PORT="COM901", BAUD=115200, serial_connection_generation=1,
            find_luat_best_port=lambda: ("COM902", "synthetic"),
            list_ports=SimpleNamespace(comports=lambda: []),
            choose_manual_rebind_candidate=lambda *_a, **_k: SimpleNamespace(
                found=True, device="COM902", description="synthetic"),
            safe_save_config=lambda **kwargs: safe_save_config_namespace_runtime(ns, **kwargs),
            system_ui=lambda *_a: None, set_status=lambda *_a: None,
            serial_wakeup_event=threading.Event(), _rebind_hint_notice=SimpleNamespace(reset=lambda: None),
            manual_rebind_hint=lambda *_a: "synthetic rebind",
        )
        return ns

    def apply(self, ns, mode, port, baud, *, before_apply=lambda: None):
        def set_state(mode, port, baud):
            before_apply()
            ns.update(MODE=mode, PORT=port, BAUD=baud)

        return apply_serial_setting_runtime(
            mode, port, baud, config=ns["config"], save_config=ns["safe_save_config"],
            set_serial_state=set_state, set_status=lambda *_a: None,
            safe_close_serial=lambda: ns.update(serial_connection_generation=ns["serial_connection_generation"] + 1),
            wake_serial=lambda: None, system_ui=lambda *_a: None,
        )

    def assert_consistent(self, ns, mode, port, baud):
        disk = load_config_snapshot(ns["CONFIG_FILE"])["serial"]
        memory = dict(ns["config"]["serial"])
        self.assertEqual(memory, disk)
        self.assertEqual((disk["mode"], disk["port"], disk["baud"]), (mode, port, str(baud)))
        self.assertEqual((ns["MODE"], ns["PORT"], ns["BAUD"]), (mode, port, baud))
        self.assertEqual(disk["unknown"], "keep")
        self.assertFalse(ns["_CONFIG_SAVE_ACTIVE"])
        self.assertNotIn("_CONFIG_SAVE_THREAD", ns)
        self.assertEqual(self.fixture.errors, [])
        self.assertFalse(self.root.tk.call("after", "info"))
        self.assertFalse(self.root.tk.globalgetvar("background_errors"))

    def test_serial_ui_and_rebind_keep_independent_commit_results(self):
        for mode in ("Auto", "Manual"):
            for background_fails in (False, True):
                for ui_fails in (False, True):
                    with self.subTest(mode=mode, background_fails=background_fails, ui_fails=ui_fails):
                        ns = self.namespace()
                        entered, release = threading.Event(), threading.Event()
                        results = []

                        def save(**kwargs):
                            background = threading.current_thread() is not threading.main_thread()

                            def replace(source, target):
                                if background:
                                    entered.set()
                                    if not release.wait(3):
                                        raise TimeoutError("Test did not release writer")
                                if background_fails if background else ui_fails:
                                    raise PermissionError("synthetic write failure")
                                Path(source).replace(target)

                            return safe_save_config_runtime(**kwargs, replace_file=replace)

                        worker = threading.Thread(target=lambda: results.append(try_rebind_manual_port_namespace_runtime(ns)))
                        requested_port = "COM903" if mode == "Manual" else ""
                        try:
                            with patch("sms_ui.app_infrastructure_namespace_runtime.safe_save_config_runtime", new=save):
                                worker.start()
                                self.assertTrue(entered.wait(2))
                                self.root.after(40, release.set)
                                self.assertEqual(self.apply(ns, mode, requested_port, 57600), not ui_fails)
                        finally:
                            release.set()
                            worker.join(3)
                        self.assertFalse(worker.is_alive())
                        self.assertEqual(results, [not background_fails])
                        if ui_fails:
                            self.assert_consistent(ns, "Manual", "COM901" if background_fails else "COM902", 115200)
                        else:
                            self.assert_consistent(ns, mode, requested_port, 57600)

    def test_stale_scan_does_not_overwrite_committed_ui_choice(self):
        cases = [
            ("Auto", "", 57600, False),
            ("Manual", "COM903", 57600, False),
            ("Manual", "COM901", 115200, False),
            ("Manual", "COM903", 57600, True),
        ]
        for mode, port, baud, before_runtime in cases:
            with self.subTest(mode=mode, port=port, before_runtime=before_runtime):
                ns = self.namespace()
                entered, release = threading.Event(), threading.Event()
                results, hints = [], []

                def scan():
                    entered.set()
                    if not release.wait(3):
                        raise TimeoutError("Test did not release scan")
                    return "COM902", "synthetic"

                def release_scan():
                    release.set()
                    worker.join(2)
                    self.assertFalse(worker.is_alive())

                ns["find_luat_best_port"] = scan
                ns["_rebind_hint_notice"] = SimpleNamespace(reset=lambda: hints.append(True))
                worker = threading.Thread(target=lambda: results.append(try_rebind_manual_port_namespace_runtime(ns)))
                try:
                    worker.start()
                    self.assertTrue(entered.wait(2))
                    self.assertTrue(self.apply(ns, mode, port, baud,
                        before_apply=release_scan if before_runtime else lambda: None))
                    release_scan()
                    self.assertEqual(results, [False])
                    self.assertEqual(hints, [])
                    self.assertFalse(ns["serial_wakeup_event"].is_set())
                    self.assert_consistent(ns, mode, port, baud)
                finally:
                    release.set()
                    worker.join(3)


if __name__ == "__main__":
    unittest.main()
