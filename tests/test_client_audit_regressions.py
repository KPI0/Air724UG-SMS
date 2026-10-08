"""Regressions from the client audit; all I/O uses synthetic data."""
import json
import threading
import tkinter as tk
import unittest
from unittest.mock import patch

from sms_core import serial_debug, third_push_sender
from sms_core.cloud_message_runtime import _cloud_transaction_parameters
from sms_core.serial_io_runtime import read_serial_line_safely_runtime
from sms_core.serial_runtime import SerialLineDecoder
from sms_core.serial_sender import SmsPduSendCoordinator, write_text_sms_pdu_locked
from sms_core.sms_pdu import encode_text_sms_pdus
from sms_core.third_push import dispatch_push_item
from sms_ui.config_sync_runtime import ConfigFileWatchState, schedule_config_file_watch_runtime
from sms_ui.serial_debug_pin_dialogs import (
    open_input_pin_dialog, open_input_puk_dialog, open_pin_lock_dialog, open_modify_pin_dialog,
)
from sms_ui.serial_debug_identity_dialogs import open_modify_number_dialog, open_modify_sn_dialog
from sms_ui.serial_debug_sms_call_dialogs import open_send_sms_dialog, open_dial_dialog
from tests.test_config_sync_runtime import FakeRootTimer


class SerialBoundsTests(unittest.TestCase):
    def test_read_uses_bounded_chunk_without_waiting_for_newline(self):
        class Burst:
            is_open = True
            in_waiting = 1024 * 1024

            def read(self, size):
                self.size = size
                return b'x' * size

            def readline(self):
                raise AssertionError('Unbounded line read')

        serial = Burst()
        result = read_serial_line_safely_runtime(threading.Lock(), lambda: serial, RuntimeError)
        self.assertGreater(len(result), 0)
        self.assertLessEqual(len(result), 4096)

    def test_unterminated_line_is_bounded_and_decoder_can_recover(self):
        decoder = SerialLineDecoder(max_line_chars=32)
        self.assertEqual(decoder.feed(b'x' * 32), [])
        with self.assertRaisesRegex(ValueError, '串口单行'):
            decoder.feed(b'x')
        self.assertEqual(decoder.text_buffer, '')
        self.assertEqual(decoder.feed(b'OK\r\n'), ['OK'])

    def test_newline_does_not_bypass_line_size_limit(self):
        with self.assertRaisesRegex(ValueError, '串口单行'):
            SerialLineDecoder(max_line_chars=32).feed(b'x' * 33 + b'\r\n')

    def test_coalesced_lines_and_trailing_sms_prompt_are_delivered(self):
        decoder = SerialLineDecoder()
        self.assertEqual(decoder.feed(b'OK\r\n> '), ['OK'])
        self.assertEqual(decoder.feed(b''), ['>'])

    def test_sms_body_starting_with_arrow_survives_chunk_boundary(self):
        for prefix in (b'>', b'> '):
            decoder = SerialLineDecoder()
            self.assertEqual(decoder.feed(prefix), [])
            self.assertEqual(decoder.feed(b'body\r\n'), [(prefix + b'body').decode()])

    def test_artificial_utf8_newline_is_handled_inside_burst(self):
        raw = b'prefix ' + '中'.encode()[:1] + b'\r\n' + '中'.encode()[1:] + b'\r\nOK\r\n'
        for index in range(1, len(raw)):
            with self.subTest(index=index):
                decoder = SerialLineDecoder()
                result = decoder.feed(raw[:index]) + decoder.feed(raw[index:])
                self.assertEqual(result, ['prefix 中', 'OK'])

    def test_many_valid_lines_do_not_count_as_one_oversized_line(self):
        decoder = SerialLineDecoder(max_line_chars=16)
        self.assertEqual(decoder.feed(b'valid\r\n' * 100), ['valid'] * 100)


class PushResponseValidationTests(unittest.TestCase):
    success = {
        'dingtalk': {'errcode': 0}, 'wecom': {'errcode': 0}, 'feishu': {'code': 0},
        'pushdeer': {'code': 0}, 'serverchan': {'code': 0}, 'bark': {'code': 200},
        'telegram': {'ok': True}, 'pushover': {'status': 1}, 'gotify': {'id': 123},
    }

    def test_known_channels_do_not_confirm_invalid_or_missing_result(self):
        for channel in self.success:
            for body in ('', '<html>unavailable</html>', '{}', '[]', 'null', 'true', '{'):
                with self.subTest(channel=channel, body=body):
                    ok, message = third_push_sender.api_ok(channel, True, 200, body)
                    self.assertFalse(ok)
                    self.assertIn('无法确认', message)

    def test_each_known_channel_accepts_its_own_success_shape(self):
        for channel, response in self.success.items():
            with self.subTest(channel=channel):
                self.assertTrue(third_push_sender.api_ok(channel, True, 200, json.dumps(response))[0])
        self.assertTrue(third_push_sender.api_ok('feishu', True, 200, '{"StatusCode":0}')[0])

    def test_explicit_failure_flags_are_not_confirmed_or_echoed(self):
        cases = {'telegram': {'ok': False}, 'bark': {'code': 400},
                 'pushover': {'status': 0}, 'gotify': {'error': 'synthetic-private-body'},
                 'wecom': {'errcode': 1}, 'feishu': {'code': 1}}
        for channel, response in cases.items():
            with self.subTest(channel=channel):
                response['message'] = 'synthetic-private-body'
                ok, message = third_push_sender.api_ok(channel, True, 200, json.dumps(response))
                self.assertFalse(ok)
                self.assertNotIn('synthetic-private-body', message)

    def test_custom_http_webhook_keeps_empty_and_text_success_compatibility(self):
        for body in ('', 'ok', '{}', 'null'):
            self.assertTrue(third_push_sender.api_ok('custom_post', True, 204, body)[0])

    def test_error_page_is_a_failed_channel_at_dispatch_boundary(self):
        with patch.object(third_push_sender, 'http_request', return_value=(True, 200, '<html>error</html>')):
            result = dispatch_push_item(dict(channels=['dingtalk'], message='synthetic',
                settings={'dingtalk_webhook': 'https://example.invalid/hook'}), third_push_sender.send_channel)
        self.assertEqual(result.ok_channels, [])
        self.assertEqual(len(result.fail_infos), 1)


class RestartWatchTests(unittest.TestCase):
    def options(self, root, state, lifecycle, changes):
        return dict(state=state, config_file='synthetic', interval_ms=1000,
                    root_after=root.after, root_after_cancel=root.after_cancel,
                    tk_alive=lambda: lifecycle['alive'], is_stopping=lambda: lifecycle['stopping'],
                    signature_func=lambda _: lifecycle['signature'], on_change=lambda: changes.append(True))

    def test_transient_exit_pauses_watch_then_resumes_changed_file(self):
        root, state, changes = FakeRootTimer(), ConfigFileWatchState(), []
        lifecycle = dict(alive=True, stopping=False, signature=1)
        schedule_config_file_watch_runtime(**self.options(root, state, lifecycle, changes))
        lifecycle.update(stopping=True, signature=2)
        root.run(state.after_id)
        self.assertEqual(changes, [])
        self.assertEqual(len(root.callbacks), 1)
        lifecycle['stopping'] = False
        root.run(state.after_id)
        self.assertEqual(changes, [True])
        lifecycle['alive'] = False
        root.run(state.after_id)
        self.assertIsNone(state.after_id)
        self.assertFalse(root.callbacks)

    def test_stale_callback_cannot_clear_new_generation_timer(self):
        root, state, changes = FakeRootTimer(), ConfigFileWatchState(), []
        lifecycle = dict(alive=True, stopping=False, signature=1)
        options = self.options(root, state, lifecycle, changes)
        old_id = schedule_config_file_watch_runtime(**options)
        old_callback = root.callbacks[old_id][1]
        new_id = schedule_config_file_watch_runtime(**options)
        old_callback()
        self.assertEqual(state.after_id, new_id)
        self.assertEqual(len(root.callbacks), 1)


class StructuredInputTests(unittest.TestCase):
    def test_pin_and_puk_builders_reject_invalid_fields(self):
        calls = [lambda value: serial_debug.build_pin_unlock_command(value),
                 lambda value: serial_debug.build_pin_lock_command(value, True),
                 lambda value: serial_debug.build_pin_change_command(value, '1234'),
                 lambda value: serial_debug.build_pin_change_command('1234', value),
                 lambda value: serial_debug.build_puk_unlock_command('12345678', value)]
        for build in calls:
            for value in ('', '123', '123456789', '12AB', '１２３４', '1234"\r\nAT\r\n"'):
                with self.subTest(value=value), self.assertRaises(ValueError):
                    build(value)
        for value in ('1234', '123456789', '１２３４５６７８', '1234\n5678'):
            with self.assertRaises(ValueError):
                serial_debug.build_puk_unlock_command(value, '1234')

    def test_other_structured_commands_reject_command_separators(self):
        for build in (serial_debug.build_dial_command, serial_debug.build_own_number_commands,
                      serial_debug.build_sn_command):
            for value in ('1234;AT', '1234\r\nAT', '1234"', '1234\x1aAT'):
                with self.subTest(build=build.__name__, value=value), self.assertRaises(ValueError):
                    build(value)

    def test_valid_dial_mmi_and_ascii_pin_remain_supported(self):
        self.assertEqual(serial_debug.build_dial_command('*#06#'), 'ATD*#06#;')
        self.assertEqual(serial_debug.build_dial_command('10086'), 'ATD10086;')
        self.assertEqual(serial_debug.build_pin_unlock_command('1234'), 'AT+CPIN="1234"')
        self.assertEqual(serial_debug.build_sn_command('SN-01_02'), 'AT+WISN=SN-01_02')

    def test_country_prefix_alone_cannot_become_an_empty_dial(self):
        with self.assertRaises(ValueError):
            serial_debug.build_dial_command('+86')

    def test_non_ascii_numbers_are_rejected_by_encoder_and_cloud(self):
        for phone in ('+８６１３８００００００００', '+٨٦١٣٨٠٠٠٠٠٠٠٠'):
            with self.subTest(phone=phone), self.assertRaises(ValueError):
                encode_text_sms_pdus(phone, 'synthetic')
            for command, key in (('CLOUD:SEND_SMS', 'sms_phone'), ('CLOUD:SET_OWN_NUMBER', 'own_number')):
                _, error = _cloud_transaction_parameters(command, {key: phone, 'sms_message': 'synthetic'})
                self.assertTrue(error)

    def test_invalid_number_never_touches_serial(self):
        for phone in ('', '+８６１３８００００００００', '1' * 256):
            touched = []
            result = write_text_sms_pdu_locked(threading.Lock(), lambda: touched.append(True),
                phone, 'synthetic', response_coordinator=SmsPduSendCoordinator())
            self.assertFalse(result)
            self.assertEqual(touched, [])


class StructuredFormTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(str(exc))
        self.root.withdraw()
        self.enabled = tk.BooleanVar(master=self.root, value=True)

    def tearDown(self):
        self.root.destroy()

    def descendants(self, widget):
        for child in widget.winfo_children():
            yield child
            yield from self.descendants(child)

    def test_invalid_fields_keep_form_open_and_do_not_send(self):
        sent = []
        common = (self.root, self.enabled)
        cases = [
            (lambda: open_input_pin_dialog(*common, sent.append, lambda *a: None), ['12AB'], ['1234'], '发送指令'),
            (lambda: open_input_puk_dialog(*common, sent.append, lambda *a: None), ['12345678', '12AB'], ['12345678', '1234'], '发送指令'),
            (lambda: open_pin_lock_dialog(*common, sent.append, lambda *a: None, True), ['12AB'], ['1234'], '发送指令'),
            (lambda: open_modify_pin_dialog(*common, sent.append, lambda *a: None), ['1234', '12AB'], ['1234', '5678'], '发送指令'),
            (lambda: open_modify_number_dialog(*common, sent.append, lambda *a: None), ['123\nAT'], ['+8613800000000'], '发送指令'),
            (lambda: open_modify_sn_dialog(*common, sent.append, lambda *a: None), ['SN;AT'], ['SN1234'], '发送指令'),
            (lambda: open_send_sms_dialog(*common, lambda *a: sent.append(a), lambda *a: None), ['+８６１３８００００００００'], ['10086'], '发送指令'),
            (lambda: open_dial_dialog(*common, sent.append, lambda: None, lambda: None, lambda *a: None), ['10086;AT'], ['10086'], '📞 拨号'),
        ]
        for opening, values, valid_values, button_text in cases:
            sent.clear()
            with self.subTest(values=values), patch('tkinter.messagebox.showerror') as error:
                opening()
                win = self.root.winfo_children()[0]
                win.withdraw()
                widgets = list(self.descendants(win))
                entries = [w for w in widgets if w.winfo_class() == 'TEntry']
                for entry, value in zip(entries, values):
                    entry.insert(0, value)
                for widget in widgets:
                    if widget.winfo_class() == 'Text':
                        widget.insert('1.0', 'synthetic')
                button = next(w for w in widgets if w.winfo_class() == 'TButton' and w.cget('text') == button_text)
                try:
                    button.invoke()
                    self.assertTrue(win.winfo_exists())
                    self.assertEqual([entry.get() for entry in entries], values)
                    error.assert_called_once()
                    self.assertEqual(sent, [])
                    for entry, value in zip(entries, valid_values):
                        entry.delete(0, 'end')
                        entry.insert(0, value)
                    button.invoke()
                    self.assertEqual(len(sent), 1)
                    self.assertFalse(win.winfo_exists())
                finally:
                    if win.winfo_exists():
                        win.destroy()


if __name__ == '__main__':
    unittest.main()
