import unittest
from contextlib import ExitStack
from types import SimpleNamespace
from unittest.mock import patch

from sms_app import bootstrap
from sms_ui.app_instance_runtime import claim_instance_number_app_runtime


class BootstrapInstanceClaimTests(unittest.TestCase):
    def run_guards(self, acquire_mutex, *, limit=3):
        values = {
            'RESTART_HELPER_FLAG': '--restart-helper', 'APP_DIR': 'synthetic-client',
            'AUTOSTART_STATE_FILE': 'synthetic-state.json', 'ALLOW_MULTI_INSTANCE': True,
            'AUTOSTART_IS_LAUNCHER_CANDIDATE': False, 'APP_WINDOW_TITLE': '短信监听系统',
            'LOG_PREFIX': 'system', 'TTS_DIR': 'synthetic-tts',
        }
        with ExitStack() as stack:
            for name, value in values.items():
                stack.enter_context(patch.object(bootstrap, name, value, create=True))
            for name in ('APP_INSTANCE_NUMBER', 'instance_number_mutex',
                         'AUTOSTART_DESIRED_INSTANCE_COUNT', 'AUTOSTART_ACTIVE_INSTANCE_COUNT',
                         'AUTOSTART_INSTANCE_REGISTERED', 'APP_DISPLAY_TITLE',
                         'SERIAL_DEBUG_WINDOW_TITLE', 'TTS_FILE'):
                stack.enter_context(patch.object(bootstrap, name, None, create=True))
            for name in ('maybe_run_restart_helper_mode', 'check_single_instance',
                         'log_file_only', 'request_dpi_awareness'):
                stack.enter_context(patch.object(bootstrap, name, create=True))
            show = stack.enter_context(patch.object(bootstrap, 'show_message_box', create=True))
            register = stack.enter_context(patch.object(
                bootstrap, 'register_autostart_instance', return_value=SimpleNamespace(
                    desired_count=1, active_count=1, registered=True)))
            stack.enter_context(patch.object(bootstrap, 'claim_instance_number_app_runtime',
                side_effect=lambda **kwargs: claim_instance_number_app_runtime(
                    **kwargs, max_instance_number=limit, acquire_mutex=acquire_mutex,
                    close_handle=lambda _handle: None)))
            try:
                bootstrap._run_startup_guards()
                return None, show, register, bootstrap.instance_number_mutex
            except SystemExit as exc:
                return exc.code, show, register, bootstrap.instance_number_mutex

    def test_mutex_creation_failure_stops_before_registering_shared_runtime(self):
        code, show, register, _handle = self.run_guards(lambda _name: (None, 8))
        self.assertEqual(code, 1)
        show.assert_called_once()
        self.assertIn('实例编号', show.call_args.args[1])
        register.assert_not_called()

    def test_all_occupied_slots_stop_before_registering_instance_one(self):
        code, show, register, _handle = self.run_guards(lambda _name: (123, 183))
        self.assertEqual(code, 1)
        show.assert_called_once()
        register.assert_not_called()

    def test_claimed_number_is_retained_for_normal_startup(self):
        calls = []
        def acquire(name):
            calls.append(name)
            return (123, 183) if len(calls) == 1 else (456, 0)
        code, show, register, handle = self.run_guards(acquire)
        self.assertIsNone(code)
        show.assert_not_called()
        self.assertEqual(register.call_args.kwargs['instance_number'], 2)
        self.assertEqual(handle, 456)


if __name__ == '__main__':
    unittest.main()
