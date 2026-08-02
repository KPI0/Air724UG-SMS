import json
import os
import tempfile
import unittest

from sms_core.autostart_instances import (
    AUTOSTART_CHILD_FLAG,
    MAX_AUTOSTART_INSTANCES,
    AutostartRegistrationResult,
    build_autostart_child_command,
    launch_autostart_companions,
    normalize_desired_count,
    register_autostart_instance,
    unregister_autostart_instance,
)
from sms_core.windows_runtime import close_windows_handle
from sms_ui.app_instance_runtime import (
    MAX_INSTANCE_NUMBER,
    claim_instance_number_app_runtime,
    is_instance_number_active_app_runtime,
)


class AutostartInstanceStateTests(unittest.TestCase):
    @staticmethod
    def unlocked(_app_dir, callback):
        return callback()

    def test_manual_starts_and_exit_track_current_instance_count(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "state.json")
            active = {1}

            first = register_autostart_instance(
                app_dir=temp_dir,
                state_path=state_path,
                instance_number=1,
                preserve_desired=False,
                allow_multi_instance=True,
                is_instance_active=lambda number: number in active,
                with_state_lock=self.unlocked,
            )
            self.assertEqual(first, AutostartRegistrationResult(1, True))

            active.add(2)
            second = register_autostart_instance(
                app_dir=temp_dir,
                state_path=state_path,
                instance_number=2,
                preserve_desired=False,
                allow_multi_instance=True,
                is_instance_active=lambda number: number in active,
                with_state_lock=self.unlocked,
            )
            self.assertEqual(second, AutostartRegistrationResult(2, True))

            active.remove(1)
            remaining = unregister_autostart_instance(
                app_dir=temp_dir,
                state_path=state_path,
                instance_number=1,
                is_instance_active=lambda number: number in active,
                with_state_lock=self.unlocked,
            )
            self.assertEqual(remaining, 1)

            with open(state_path, "r", encoding="utf-8") as file_obj:
                state = json.load(file_obj)
            self.assertEqual(state["desired_count"], 1)
            self.assertEqual(state["active_instances"], [2])

    def test_autostart_leader_preserves_previous_count_while_pruning_stale_entries(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "state.json")
            with open(state_path, "w", encoding="utf-8") as file_obj:
                json.dump(
                    {
                        "version": 1,
                        "desired_count": 3,
                        "active_instances": [1, 2, 3],
                    },
                    file_obj,
                )

            result = register_autostart_instance(
                app_dir=temp_dir,
                state_path=state_path,
                instance_number=1,
                preserve_desired=True,
                allow_multi_instance=True,
                is_instance_active=lambda number: number == 1,
                with_state_lock=self.unlocked,
            )

            self.assertEqual(result, AutostartRegistrationResult(3, True))
            with open(state_path, "r", encoding="utf-8") as file_obj:
                state = json.load(file_obj)
            self.assertEqual(state["desired_count"], 3)
            self.assertEqual(state["active_instances"], [1])

    def test_disabled_multi_instance_forces_desired_count_to_one(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "state.json")
            with open(state_path, "w", encoding="utf-8") as file_obj:
                json.dump(
                    {"version": 1, "desired_count": 8, "active_instances": [1]},
                    file_obj,
                )

            result = register_autostart_instance(
                app_dir=temp_dir,
                state_path=state_path,
                instance_number=1,
                preserve_desired=True,
                allow_multi_instance=False,
                is_instance_active=lambda _number: True,
                with_state_lock=self.unlocked,
            )

            self.assertEqual(result.desired_count, 1)

    def test_corrupt_state_falls_back_without_blocking_startup(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "state.json")
            with open(state_path, "w", encoding="utf-8") as file_obj:
                file_obj.write("not-json")
            logs = []

            result = register_autostart_instance(
                app_dir=temp_dir,
                state_path=state_path,
                instance_number=1,
                preserve_desired=True,
                allow_multi_instance=True,
                is_instance_active=lambda _number: False,
                log_error=logs.append,
                with_state_lock=self.unlocked,
            )

            self.assertEqual(result, AutostartRegistrationResult(1, True))
            self.assertTrue(any("Load autostart" in message for message in logs))

    def test_lock_failure_returns_saved_count_without_registering(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "state.json")
            with open(state_path, "w", encoding="utf-8") as file_obj:
                json.dump({"desired_count": 4, "active_instances": []}, file_obj)

            result = register_autostart_instance(
                app_dir=temp_dir,
                state_path=state_path,
                instance_number=1,
                preserve_desired=True,
                allow_multi_instance=True,
                is_instance_active=lambda _number: False,
                with_state_lock=lambda *_args: (_ for _ in ()).throw(RuntimeError("busy")),
            )

            self.assertEqual(result, AutostartRegistrationResult(4, False))

    def test_count_is_clamped_to_safe_range(self):
        self.assertEqual(normalize_desired_count(0), 1)
        self.assertEqual(normalize_desired_count("3"), 3)
        self.assertEqual(normalize_desired_count(100), 100)
        self.assertEqual(normalize_desired_count(1000), 1000)
        self.assertEqual(normalize_desired_count(10000), 9999)
        self.assertEqual(MAX_AUTOSTART_INSTANCES, MAX_INSTANCE_NUMBER)

    def test_state_preserves_one_hundred_active_instances(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "state.json")
            with open(state_path, "w", encoding="utf-8") as file_obj:
                json.dump(
                    {
                        "version": 1,
                        "desired_count": 100,
                        "active_instances": list(range(1, 101)),
                    },
                    file_obj,
                )

            result = register_autostart_instance(
                app_dir=temp_dir,
                state_path=state_path,
                instance_number=1,
                preserve_desired=True,
                allow_multi_instance=True,
                is_instance_active=lambda number: 1 <= number <= 100,
                with_state_lock=self.unlocked,
            )

            self.assertEqual(result, AutostartRegistrationResult(100, True))
            with open(state_path, "r", encoding="utf-8") as file_obj:
                state = json.load(file_obj)
            self.assertEqual(state["desired_count"], 100)
            self.assertEqual(state["active_instances"], list(range(1, 101)))

    @unittest.skipUnless(os.name == "nt", "requires Windows named mutexes")
    def test_real_windows_mutexes_track_two_instances(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_path = os.path.join(temp_dir, "state.json")
            first_number, first_handle = claim_instance_number_app_runtime(app_dir=temp_dir)
            second_handle = None
            try:
                first = register_autostart_instance(
                    app_dir=temp_dir,
                    state_path=state_path,
                    instance_number=first_number,
                    preserve_desired=False,
                    allow_multi_instance=True,
                    is_instance_active=lambda number: is_instance_number_active_app_runtime(
                        app_dir=temp_dir,
                        instance_number=number,
                    ),
                )
                second_number, second_handle = claim_instance_number_app_runtime(app_dir=temp_dir)
                second = register_autostart_instance(
                    app_dir=temp_dir,
                    state_path=state_path,
                    instance_number=second_number,
                    preserve_desired=False,
                    allow_multi_instance=True,
                    is_instance_active=lambda number: is_instance_number_active_app_runtime(
                        app_dir=temp_dir,
                        instance_number=number,
                    ),
                )

                self.assertEqual(first.desired_count, 1)
                self.assertEqual(second.desired_count, 2)
                self.assertEqual((first_number, second_number), (1, 2))
            finally:
                close_windows_handle(second_handle)
                close_windows_handle(first_handle)


class AutostartCompanionLaunchTests(unittest.TestCase):
    def test_build_child_command_supports_script_mode(self):
        command, workdir = build_autostart_child_command(
            "--autostart",
            launch_target_func=lambda: ("pythonw.exe", "sms.pyw", "E:/sms"),
        )

        self.assertEqual(
            command,
            ["pythonw.exe", "sms.pyw", "--autostart", AUTOSTART_CHILD_FLAG],
        )
        self.assertEqual(workdir, "E:/sms")

    def test_leader_launches_desired_companions_with_spacing(self):
        calls = []

        launched = launch_autostart_companions(
            desired_count=3,
            allow_multi_instance=True,
            is_leader=True,
            autostart_flag="--autostart",
            wait_before_launch=lambda seconds: calls.append(("wait", seconds)) or False,
            interval_seconds=0.5,
            prepare_launch=lambda *_args: (["sms.exe", "--autostart-child"], "E:/sms"),
            launch_process=lambda command, env, cwd: calls.append(
                ("launch", command, env, cwd)
            ),
            clean_env=lambda: {"clean": "1"},
        )

        self.assertEqual(launched, 2)
        self.assertEqual([item[0] for item in calls], ["wait", "launch", "wait", "launch"])
        self.assertEqual(calls[1], ("launch", ["sms.exe", "--autostart-child"], {"clean": "1"}, "E:/sms"))

    def test_child_and_single_instance_mode_never_expand(self):
        for allow_multi, is_leader in ((True, False), (False, True)):
            with self.subTest(allow_multi=allow_multi, is_leader=is_leader):
                launched = launch_autostart_companions(
                    desired_count=3,
                    allow_multi_instance=allow_multi,
                    is_leader=is_leader,
                    autostart_flag="--autostart",
                    launch_process=lambda *_args, **_kwargs: self.fail("unexpected launch"),
                )
                self.assertEqual(launched, 0)

    def test_shutdown_during_spacing_stops_launching(self):
        launched = launch_autostart_companions(
            desired_count=3,
            allow_multi_instance=True,
            is_leader=True,
            autostart_flag="--autostart",
            wait_before_launch=lambda _seconds: True,
            prepare_launch=lambda *_args: (["sms.exe"], "E:/sms"),
            launch_process=lambda *_args, **_kwargs: self.fail("unexpected launch"),
        )

        self.assertEqual(launched, 0)

    def test_leader_can_restore_one_hundred_instances(self):
        launches = []

        launched = launch_autostart_companions(
            desired_count=100,
            allow_multi_instance=True,
            is_leader=True,
            autostart_flag="--autostart",
            wait_before_launch=lambda _seconds: False,
            prepare_launch=lambda *_args: (["sms.exe", "--autostart-child"], "E:/sms"),
            launch_process=lambda command, env, cwd: launches.append((command, env, cwd)),
            clean_env=lambda: {"clean": "1"},
        )

        self.assertEqual(launched, 99)
        self.assertEqual(len(launches), 99)


if __name__ == "__main__":
    unittest.main()
