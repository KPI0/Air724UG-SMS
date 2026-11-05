import asyncio
import configparser
import json
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import AsyncMock, Mock, patch
from urllib.parse import quote, unquote, urlsplit
from types import SimpleNamespace

from sms_core.cloud_auth import secret_match_result
from sms_core.cloud_control_settings import CloudControlSettings, read_cloud_control_settings
from sms_core.cloud_message_runtime import send_cloud_register_runtime
from sms_core.cloud_runtime import validate_cloud_start
from sms_core.cloud_ws_runtime import cloud_websocket_connect_kwargs
from sms_core.cloud_firmware_ota import RelayError, UsbOtaTransport
from sms_core.serial_debug import build_pin_unlock_command, build_puk_unlock_command
from sms_core.serial_debug import build_serial_command_payload, redact_sim_credentials_for_display
from sms_core import serial_sender
from sms_app.cloud_namespace_bindings import install_cloud_namespace_bindings
from sms_core.cloud_message_namespace_runtime import send_cloud_serial_command_namespace_runtime
from test_cloud_command_connection_context import Modem
from sms_core.config_runtime import (
    ConfigInitializationError, initialize_config_runtime, reload_config_runtime,
    safe_save_config_runtime, snapshot_config_runtime,
)
from sms_core.third_push import validate_push_settings
from sms_core.third_push_config import ThirdPushSettings, read_third_push_settings
from sms_core import third_push_sender as sender
from sms_ui.cloud_control_app_runtime import save_cloud_control_setting_runtime
from sms_ui.cloud_control_window_form import CloudControlFormController
from sms_ui.cloud_placeholder_entry import CloudPlaceholderEntry
from sms_ui.third_push_app_runtime import save_third_push_setting_runtime
from test_third_push_window_form import FakeVar, make_form


BAD_CREDENTIALS = ('synthetic中文', 'synthetic🔐', 'café', ' valid ', 'bad value',
                   '\tvalid', 'valid\n', 'valid\x00', 'valid\x7f', '\ud800')
PUSH_CREDENTIALS = ('pushdeer_key', 'bark_key', 'dingtalk_secret',
                    'pushover_api_token', 'pushover_user_key', 'gotify_token',
                    'wxpusher_app_token', 'email_password', 'next_smtp_proxy_password')


class CredentialPolicyTests(unittest.TestCase):
    def test_cloud_auth_rejects_invalid_matching_credentials_and_aliases(self):
        for credential in BAD_CREDENTIALS:
            for alias in ('secret', 'device_secret', 'password', 'pwd', 'token'):
                with self.subTest(alias=alias, credential=repr(credential)):
                    ok, reason = secret_match_result({alias: credential}, credential)
                    self.assertFalse(ok)
                    self.assertNotIn(credential, reason)

    def test_cloud_auth_accepts_all_printable_ascii_symbols_exactly(self):
        credential = ''.join(chr(code) for code in range(33, 127))
        self.assertEqual(secret_match_result({'secret': credential}, credential), (True, ''))
        self.assertFalse(secret_match_result({'secret': credential + '!'}, credential)[0])

    def test_cloud_start_rejects_invalid_credential_before_connection(self):
        for credential in BAD_CREDENTIALS:
            with self.subTest(credential=repr(credential)):
                result = validate_cloud_start(True, 'wss://example.test', credential)
                self.assertFalse(result.ok)
                self.assertNotIn(credential, result.warning_message)

    def test_cloud_read_rejects_invalid_raw_config_value(self):
        config = configparser.ConfigParser(interpolation=None)
        for credential in BAD_CREDENTIALS:
            config['cloud_control'] = {'device_secret': credential}
            before = snapshot_config_runtime(config)
            with self.subTest(credential=repr(credential)), self.assertRaises(ValueError):
                read_cloud_control_settings(config)
            self.assertEqual(snapshot_config_runtime(config), before)

    def test_cloud_save_rejects_without_persisting_or_publishing(self):
        config = configparser.ConfigParser(interpolation=None)
        config['cloud_control'] = {'device_secret': 'synthetic-old'}
        original = snapshot_config_runtime(config)
        save, apply, ui = Mock(), Mock(), Mock()
        for credential in BAD_CREDENTIALS:
            with self.subTest(credential=repr(credential)):
                result = save_cloud_control_setting_runtime(
                    current_settings=lambda: CloudControlSettings(device_secret='synthetic-old'),
                    apply_settings=apply, config=config, save_config=save, system_ui=ui,
                    device_secret=credential,
                )
                self.assertIsNone(result)
                self.assertEqual(snapshot_config_runtime(config), original)
        save.assert_not_called()
        apply.assert_not_called()

    def test_register_refuses_invalid_current_or_previous_credential_without_send(self):
        for credential in BAD_CREDENTIALS:
            for previous in (False, True):
                ws, log = AsyncMock(), Mock()
                with self.subTest(credential=repr(credential), previous=previous):
                    result = asyncio.run(send_cloud_register_runtime(
                        ws, auto_upload=True, build_payload=lambda *args: {'secret': args[3]},
                        timestamp=lambda: 0, identity_payload=lambda: {},
                        secret='synthetic-valid' if previous else credential,
                        previous_session_secret=credential if previous else '',
                        serial_port='', serial_baud=115200, serial_mode='Auto',
                        runtime_imei=lambda: '123456789012345', log=log,
                    ))
                    self.assertEqual(result, 'error')
                    ws.send.assert_not_called()
                    self.assertNotIn(credential, repr(log.call_args_list))

    def test_push_validation_rejects_invalid_unselected_credentials(self):
        for key in PUSH_CREDENTIALS:
            for credential in BAD_CREDENTIALS:
                with self.subTest(key=key, credential=repr(credential)):
                    errors = validate_push_settings([], {key: credential})
                    self.assertTrue(errors)
                    self.assertNotIn(credential, str(errors))

    def test_push_read_rejects_before_default_mutation(self):
        for key in PUSH_CREDENTIALS:
            config = configparser.ConfigParser(interpolation=None)
            config['third_push'] = {key: 'synthetic中文'}
            before = snapshot_config_runtime(config)
            with self.subTest(key=key), self.assertRaises(ValueError):
                read_third_push_settings(config)
            self.assertEqual(snapshot_config_runtime(config), before)

    def test_push_save_rejects_without_persisting_or_publishing(self):
        for key in PUSH_CREDENTIALS:
            config = configparser.ConfigParser(interpolation=None)
            config['third_push'] = {'notify_type': '[]'}
            before = snapshot_config_runtime(config)
            save, apply = Mock(), Mock()
            with self.subTest(key=key):
                result = save_third_push_setting_runtime(
                    current_settings=lambda: ThirdPushSettings(False, True, True, [], {}),
                    apply_settings=apply, config=config, save_config=save,
                    settings={key: 'synthetic中文'},
                )
                self.assertIsNone(result)
                self.assertEqual(snapshot_config_runtime(config), before)
                save.assert_not_called()
                apply.assert_not_called()

    def test_direct_push_handlers_do_not_submit_bad_credentials(self):
        cases = (
            (sender.send_pushdeer, {'pushdeer_api': 'https://example.test'}, 'pushdeer_key'),
            (sender.send_bark, {'bark_api': 'https://example.test'}, 'bark_key'),
            (sender.send_dingtalk, {'dingtalk_webhook': 'https://example.test'}, 'dingtalk_secret'),
            (sender.send_pushover, {'pushover_user_key': 'synthetic'}, 'pushover_api_token'),
            (sender.send_pushover, {'pushover_api_token': 'synthetic'}, 'pushover_user_key'),
            (sender.send_gotify, {'gotify_api': 'https://example.test'}, 'gotify_token'),
            (sender.send_wxpusher, {'wxpusher_uids': 'UID_synthetic'}, 'wxpusher_app_token'),
            (sender.send_next_smtp_proxy, {'next_smtp_proxy_api': 'https://example.test',
                 'next_smtp_proxy_user': 'synthetic', 'next_smtp_proxy_host': 'example.test',
                 'next_smtp_proxy_port': '587', 'next_smtp_proxy_to_email': 'a@example.test'},
                 'next_smtp_proxy_password'),
            (sender.send_email, {'email_smtp_host': 'example.test', 'email_smtp_port': '465',
                 'email_encryption': 'ssl', 'email_username': 'synthetic',
                 'email_from_address': 'a@example.test', 'email_to_addresses': 'b@example.test'},
                 'email_password'),
        )
        for handler, settings, key in cases:
            for credential in BAD_CREDENTIALS:
                with self.subTest(handler=handler.__name__, credential=repr(credential)), \
                        patch.object(sender, 'http_request') as http, \
                        patch.object(sender, 'DeadlineSMTP_SSL') as smtp:
                    ok, info = handler('通知🔔', dict(settings, **{key: credential}))
                    self.assertFalse(ok)
                    self.assertNotIn(credential, info)
                    http.assert_not_called()
                    smtp.assert_not_called()

    def test_startup_preserves_bad_config_file_and_refuses_default_save(self):
        for section, key in (('cloud_control', 'device_secret'), ('third_push', 'email_password')):
            with tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / 'config.ini'
                data = f'[{section}]\n{key}=synthetic中文\n'.encode('utf-8')
                path.write_bytes(data)
                save, backup = Mock(), Mock()
                with self.subTest(key=key), self.assertRaises(ConfigInitializationError):
                    initialize_config_runtime(
                        config=configparser.ConfigParser(interpolation=None), config_file=str(path),
                        defaults_by_section={'cloud_control': {'enabled': '0'}},
                        save_config=save, backup_file=backup,
                    )
                self.assertEqual(path.read_bytes(), data)
                save.assert_not_called()
                backup.assert_not_called()

    def test_shared_refresh_rejects_bad_snapshot_without_publish(self):
        config = configparser.ConfigParser(interpolation=None)
        config['cloud_control'] = {'device_secret': 'synthetic-old'}
        original = snapshot_config_runtime(config)
        commit = Mock()
        with self.assertRaises(ValueError):
            reload_config_runtime(
                config=config, config_file='unused', config_lock=threading.RLock(),
                load_snapshot=lambda _: {'cloud_control': {'device_secret': 'synthetic🔐'}},
                commit_config=commit,
            )
        self.assertEqual(snapshot_config_runtime(config), original)
        commit.assert_not_called()

    def test_final_disk_save_refuses_invalid_merge_and_preserves_file(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'config.ini'
            data = b'[cloud_control]\ndevice_secret=synthetic-old\n'
            path.write_bytes(data)
            config = configparser.ConfigParser(interpolation=None)
            config['cloud_control'] = {'device_secret': 'synthetic中文'}
            result = safe_save_config_runtime(
                config=config, config_file=str(path), config_lock=threading.RLock(),
                acquire_process_lock=lambda *a, **k: (object(), 0), release_process_lock=lambda _: None,
            )
            self.assertFalse(result)
            self.assertEqual(path.read_bytes(), data)
            self.assertEqual(list(Path(directory).glob('*.tmp')), [])

    def test_cloud_form_rejects_raw_secret_and_preserves_disabled_empty(self):
        form = CloudControlFormController.__new__(CloudControlFormController)
        form.win = 'synthetic-window'
        form.enabled_var, form.auto_upload_var = FakeVar(False), FakeVar(False)
        form.url_field = SimpleNamespace(get=lambda: 'wss://example.test', set_value=lambda _: None)
        form.secret_field = Mock()
        form._read_interval = lambda: 5
        for credential in BAD_CREDENTIALS:
            form.secret_field.get.return_value = credential
            with self.subTest(credential=repr(credential)), \
                    patch('sms_ui.cloud_control_window_form.messagebox.showerror') as error:
                self.assertIsNone(form.read())
                error.assert_called_once()
        form.secret_field.get.return_value = ''
        self.assertEqual(form.read()[3], '')

    def test_secret_widget_and_push_form_keep_raw_whitespace_for_rejection(self):
        entry = CloudPlaceholderEntry.__new__(CloudPlaceholderEntry)
        entry.variable, entry.placeholder_active, entry.random_secret = FakeVar(' synthetic '), False, True
        self.assertEqual(entry.get(), ' synthetic ')
        state = dict(enabled=False, sms_enabled=True, call_enabled=True, channels=[],
                     settings={'email_password': ' synthetic '})
        form = make_form(state)
        form.select_channel = Mock()
        self.assertEqual(form.collect(validate=False)[4]['email_password'], ' synthetic ')
        with patch('sms_ui.third_push_window_form.messagebox.showerror'):
            self.assertIsNone(form.collect(validate=True))

    def test_optional_empty_and_ordinary_unicode_message_fields_remain_valid(self):
        settings = dict.fromkeys(PUSH_CREDENTIALS, '')
        settings.update(custom_post_body='{"message":"中文🔔"}', dingtalk_keyword='通知',
                        gotify_title='中文🔔', email_subject='中文🔔',
                        custom_post_url='https://example.test/中文')
        self.assertEqual(validate_push_settings([], settings), [])
        config = configparser.ConfigParser(interpolation=None)
        config['cloud_control'] = {'device_secret': ''}
        self.assertEqual(read_cloud_control_settings(config).device_secret, '')

    def test_final_websocket_connect_refuses_decoded_password_or_token(self):
        factory = Mock()
        for credential in BAD_CREDENTIALS:
            encoded = quote(credential, safe='', errors='surrogatepass')
            for url in (f'wss://user:{encoded}@example.test', f'wss://example.test/ws?token={encoded}'):
                with self.subTest(url=url), self.assertRaises(ValueError):
                    cloud_websocket_connect_kwargs(url, ssl_context_factory=factory)
        factory.assert_not_called()
        from websockets.uri import parse_uri
        url = 'wss://user%3A%E4%B8%AD:synthetic-valid@example.test'
        parsed = parse_uri(url)
        self.assertEqual(parsed.username, 'user%3A%E4%B8%AD')
        self.assertEqual(parsed.password, 'synthetic-valid')
        self.assertEqual(cloud_websocket_connect_kwargs(url, ssl_context_factory=lambda _: 'synthetic-context')['ssl'],
                         'synthetic-context')

    def test_push_url_tokens_rejected_before_http_in_known_routes(self):
        cases = ((sender.send_telegram, 'telegram_api', '/bot{token}/sendMessage', {'telegram_chat_id': '123'}),
                 (sender.send_feishu, 'feishu_webhook', '/open-apis/bot/v2/hook/{token}', {}),
                 (sender.send_serverchan, 'serverchan_api', '/{token}.send', {'serverchan_title': '标题'}),
                 (sender.send_wecom, 'wecom_webhook', '/send?key={token}', {}))
        for handler, key, route, settings in cases:
            for credential in BAD_CREDENTIALS:
                token = quote(credential, safe='', errors='surrogatepass')
                with self.subTest(handler=handler.__name__, token=token), patch.object(sender, 'http_request') as http:
                    ok, _ = handler('中文🔔', dict(settings, **{key: 'https://example.test' + route.format(token=token)}))
                    self.assertFalse(ok)
                    http.assert_not_called()

    def test_push_url_all_named_query_credentials_and_basic_passwords_rejected(self):
        for name in ('access_token', 'auth_token', 'token', 'key', 'secret', 'password', 'api_key', 'apikey', 'app_key', 'appkey',
                     '%61ccess_token', '%61uth_token', '%70assword'):
            for raw in ('%E4%B8%AD', '%F0%9F%94%90', 'bad+value', 'bad%20value', 'bad%0Avalue'):
                with self.subTest(name=name, raw=raw):
                    self.assertTrue(validate_push_settings([], {'custom_post_url': f'https://example.test/send?{name}={raw}'}))
        self.assertTrue(validate_push_settings([], {'custom_post_url': 'https://user:%E4%B8%AD@example.test/'}))

    def test_push_urls_preserve_literal_and_encoded_plus_and_ordinary_unicode(self):
        for settings in ({'telegram_api': 'https://example.test/botabc+123/sendMessage'},
                         {'custom_post_url': 'https://example.test/中文?message=中文🔔&token=abc%2B123'},
                         {'custom_post_url': 'https://用户名:abc%2B123@example.test/中文'},
                         {'inotify_api': 'https://example.test/普通路由/中文?message=中文🔔'}):
            self.assertEqual(validate_push_settings([], settings), [])

    def test_pin_and_puk_keep_digit_lengths_and_reject_raw_whitespace_or_unicode(self):
        self.assertEqual(build_pin_unlock_command('1234'), 'AT+CPIN="1234"')
        self.assertEqual(build_puk_unlock_command('12345678', '5678'), 'AT+CPIN="12345678","5678"')
        for value in (' 1234 ', '1234\n', '１２３４', '密码', '🔐', '123'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                build_pin_unlock_command(value)
        for value in (' 12345678 ', '12345678\n', '１２３４５６７８', '🔐'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                build_puk_unlock_command(value, '5678')

    def test_usb_ota_final_write_refuses_invalid_secret_metadata(self):
        transport = UsbOtaTransport(SimpleNamespace(imei='123456789012345'), lambda: True)
        transport.port = Mock()
        for action in ('begin', 'commit'):
            for credential in BAD_CREDENTIALS:
                with self.subTest(action=action, credential=repr(credential)), self.assertRaises(RelayError):
                    transport._write({'action': action, 'secret': credential})
        transport.port.write.assert_not_called()

    def test_usb_ota_invalid_exchange_does_not_open_or_reserve_device(self):
        transport = UsbOtaTransport(SimpleNamespace(imei='123456789012345'), lambda: True)
        transport.connect = Mock()
        for action in ('begin', 'commit'):
            with self.subTest(action=action), self.assertRaises(RelayError):
                transport.exchange({'action': action, 'job_id': 'a' * 32, 'secret': 'synthetic中文'})
        transport.connect.assert_not_called()
        self.assertIsNone(transport.job_id)
        self.assertFalse(transport.commit_sent)

    def test_usb_transport_preserves_local_optional_empty_credential(self):
        transport = UsbOtaTransport(SimpleNamespace(imei='123456789012345'), lambda: True)
        frames = []
        def write(data):
            frames.append(json.loads(data))
            return len(data)
        transport.port = SimpleNamespace(write=write)
        for fields in ({}, {'secret': ''}):
            for action in ('begin', 'commit'):
                transport._write(dict(action=action, **fields))
        self.assertEqual(len(frames), 4)
        self.assertNotIn('secret', frames[0])
        self.assertEqual(frames[2]['secret'], '')

    def test_bark_keeps_all_ascii_credential_symbols_in_one_path_segment(self):
        credential = ''.join(chr(code) for code in range(33, 127))
        with patch.object(sender, 'http_request', return_value=(True, 200, '{"code":200}')) as http:
            ok, _ = sender.send_bark('通知🔔', {'bark_api': 'https://example.test', 'bark_key': credential})
        self.assertTrue(ok)
        url = http.call_args.args[0]
        self.assertEqual(unquote(urlsplit(url).path[1:]), credential)
        self.assertEqual(urlsplit(url).query, '')
        self.assertEqual(urlsplit(url).fragment, '')

    def test_ascii_symbols_round_trip_real_config_and_clear_invalid_old_credential(self):
        credential = ''.join(chr(code) for code in range(33, 127))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'config.ini'
            config = configparser.ConfigParser(interpolation=None)
            config['cloud_control'] = {'device_secret': credential}
            config['third_push'] = dict.fromkeys(PUSH_CREDENTIALS, credential)
            lock = threading.RLock()
            def save():
                return safe_save_config_runtime(
                    config=config, config_file=str(path), config_lock=lock,
                    acquire_process_lock=lambda *a, **k: (object(), 0), release_process_lock=lambda _: None,
                )
            self.assertTrue(save())
            loaded = configparser.ConfigParser(interpolation=None)
            loaded.read(path, encoding='utf-8')
            self.assertEqual(read_cloud_control_settings(loaded).device_secret, credential)
            push_settings = read_third_push_settings(loaded)
            for key in PUSH_CREDENTIALS:
                self.assertEqual(push_settings.settings[key], credential)
            path.write_text('[cloud_control]\ndevice_secret=synthetic中文\n', encoding='utf-8')
            config.clear()
            config['cloud_control'] = {'device_secret': ''}
            self.assertTrue(save())
            loaded.clear()
            loaded.read(path, encoding='utf-8')
            self.assertEqual(read_cloud_control_settings(loaded).device_secret, '')

    def test_raw_sim_credential_builder_rejects_unicode_and_quoted_inner_whitespace(self):
        for credential in BAD_CREDENTIALS:
            for command in (f'AT+CPIN="{credential}"', f'AT+CPIN2="{credential}","1234"',
                            f'AT+CPWD="SC","1234","{credential}"', f'AT+CLCK="AO",1,"{credential}"'):
                with self.subTest(command=repr(command)), self.assertRaises(ValueError):
                    build_serial_command_payload(command)
        with self.assertRaises(ValueError):
            build_serial_command_payload('AT+CPIN="12\ud80034"')
        with self.assertRaises(ValueError):
            build_serial_command_payload('ATI\ud800')
        for command in ('AT+CPIN=\u00a01234', 'AT+CPWD="SC",1234,\u00a05678',
                        'AT+CLCK="SC",1,"1234"\u00a0'):
            with self.subTest(command=repr(command)), self.assertRaises(ValueError):
                build_serial_command_payload(command)

    def test_raw_sim_valid_syntax_queries_and_ordinary_unicode_are_unchanged(self):
        commands = (' AT + CPIN = "1234"  ', 'AT+CPIN="1234" \t', 'AT+CPIN=?', 'AT+CPIN?', 'AT+CPIN2=?',
                    'AT+CPWD=?', 'AT+CLCK=?', 'AT+CLCK="SC",2',
                    'AT+CPWD="AO","old,!;\\\"key","new,!;\\\"key"',
                    'AT+CPIN="1234"\r\nAT+CLCK="SC",1,"5678"',
                    'AT+CPIN="1234";AT+CLCK="SC",1,"5678"',
                    'AT+CPBW=1,"10086",129,"中文🔔"')
        for command in commands:
            with self.subTest(command=command):
                self.assertEqual(build_serial_command_payload(command)[0], (command+'\r\n').encode('utf-8'))

    def test_unterminated_raw_password_preserves_its_control_character_tail(self):
        modem = Modem(serial_sender.AtCommandResponseCoordinator(), serial_sender.SmsPduSendCoordinator())
        for tail in ('\r\n', '\n', '\t'):
            for text in ('AT+CPIN="Good', 'AT+CPIN="Good\\"'):
                command = text + tail
                with self.subTest(command=repr(command)):
                    with self.assertRaises(ValueError):
                        build_serial_command_payload(command)
                    self.assertFalse(serial_sender.write_serial_command_result(modem, command).ok)
        self.assertEqual(modem.writes, [])

    def test_compound_at_inherits_prefix_and_refuses_bad_password_before_uart(self):
        modem = Modem(serial_sender.AtCommandResponseCoordinator(), serial_sender.SmsPduSendCoordinator())
        for command in ('AT+CPIN="1234";+CPWD="SC","1234","synthetic中文"',
                        'AT+CSQ;+CLCK="AO",1,"synthetic中文"',
                        'AT+CSQ;+CPIN="bad value"', 'AT+CSQ;+CPIN2="bad value"'):
            with self.subTest(command=repr(command)):
                with self.assertRaises(ValueError):
                    build_serial_command_payload(command)
                self.assertFalse(serial_sender.write_serial_command_result(modem, command).ok)
        self.assertEqual(modem.writes, [])
        command = 'AT+CSQ;+CLCK="SC",2'
        self.assertTrue(serial_sender.write_serial_command_result(modem, command).ok)
        self.assertEqual(modem.writes, [(command + '\r\n').encode('utf-8')])

    def test_compound_valid_passwords_are_redacted_and_newlines_reset_at_context(self):
        command = 'AT+CSQ;+CPWD="SC","synthetic-old","synthetic-new"'
        self.assertEqual(build_serial_command_payload(command)[0], (command + '\r\n').encode('utf-8'))
        display = redact_sim_credentials_for_display(command)
        self.assertNotIn('synthetic-old', display)
        self.assertNotIn('synthetic-new', display)
        for text in ('+CPIN="普通正文"', '中文;+CPIN="普通正文"',
                     'AT+CSQ\r\n+CPIN="普通正文"'):
            self.assertEqual(build_serial_command_payload(text)[0], (text + '\r\n').encode('utf-8'))

    def test_known_apn_password_rejected_without_limiting_apn_or_username(self):
        modem = Modem(serial_sender.AtCommandResponseCoordinator(), serial_sender.SmsPduSendCoordinator())
        for credential in BAD_CREDENTIALS:
            for prefix in ('AT', 'AT+CSQ;'):
                command = f'{prefix}+CPNETAPN=2,"中文APN","用户名","{credential}",2'
                with self.subTest(command=repr(command)):
                    with self.assertRaises(ValueError):
                        build_serial_command_payload(command)
                    self.assertFalse(serial_sender.write_serial_command_result(modem, command).ok)
        self.assertEqual(modem.writes, [])
        command = 'AT+CSQ;+CPNETAPN=2,"中文APN","用户名","synthetic-valid",2'
        self.assertTrue(serial_sender.write_serial_command_result(modem, command).ok)
        self.assertEqual(modem.writes, [(command + '\r\n').encode('utf-8')])
        self.assertNotIn('synthetic-valid', redact_sim_credentials_for_display(command))
        self.assertEqual(redact_sim_credentials_for_display('AT+CPNETAPN?'), 'AT+CPNETAPN?')

    def test_serial_writer_refuses_bad_credentials_and_next_at_still_works(self):
        at, sms = serial_sender.AtCommandResponseCoordinator(), serial_sender.SmsPduSendCoordinator()
        modem = Modem(at, sms)
        for command in ('AT+CPIN="12\ud80034"', 'AT+CLCK="AO",1,"synthetic中文"',
                        'AT+CPIN="1234"\r\nAT+CPWD="SC","1234","bad value"'):
            with self.subTest(command=repr(command)):
                self.assertFalse(serial_sender.write_serial_command_result(modem, command).ok)
        self.assertEqual(modem.writes, [])
        self.assertTrue(serial_sender.write_serial_command_result(modem, 'ATI').ok)
        self.assertEqual(modem.writes, [b'ATI\r\n'])

    def test_confirmed_serial_writer_refuses_before_waiter_and_continues(self):
        at, sms = serial_sender.AtCommandResponseCoordinator(), serial_sender.SmsPduSendCoordinator()
        modem, lock = Modem(at, sms), threading.RLock()
        result = serial_sender.write_serial_command_confirmed_locked(
            lock, lambda: modem, 'AT+CPIN="12\ud80034"', response_coordinator=at,
        )
        self.assertFalse(result.ok)
        self.assertEqual(modem.writes, [])
        self.assertTrue(serial_sender.write_serial_command_confirmed_locked(
            lock, lambda: modem, 'ATI', response_coordinator=at,
        ).ok)

    def test_formal_cloud_namespace_refuses_raw_credentials_with_zero_uart_bytes(self):
        at, sms = serial_sender.AtCommandResponseCoordinator(), serial_sender.SmsPduSendCoordinator()
        modem, websocket = Modem(at, sms), object()
        namespace = dict(serial_lock=threading.RLock(), serial_obj=modem, serial_connection_generation=1,
                         cloud_imei_verified=True, CLOUD_DEVICE_IMEI='000000000000001',
                         CLOUD_DEVICE_SECRET='synthetic-only', CLOUD_WS_URL='wss://example.test/ws/device',
                         CLOUD_CONTROL_ENABLED=True, cloud_ws_conn=websocket, cloud_connected=True,
                         cloud_device_authorized=True, SMS_SEND_COORDINATOR=sms,
                         SERIAL_COMMAND_RESPONSE_COORDINATOR=at,
                         _normalize_imei=lambda value: value,
                         CLOUD_SENSITIVE_COMMAND_PERMISSIONS={'pin': True, 'puk': True, 'call_control': True})
        install_cloud_namespace_bindings(namespace)
        messages = []
        namespace['_cloud_log'] = lambda message, **_: messages.append(message)
        for command in ('AT+CPIN="12\ud80034"', 'AT+CPWD="AO","1234","synthetic中文"',
                        'AT+CLCK="SC",1,"bad value"',
                        'AT+CSQ;+CPNETAPN=2,"APN","用户名","synthetic中文",2'):
            with self.subTest(command=repr(command)):
                self.assertFalse(send_cloud_serial_command_namespace_runtime(namespace, command)[0])
        self.assertEqual(modem.writes, [])
        self.assertNotIn('synthetic中文', str(messages))
        self.assertTrue(send_cloud_serial_command_namespace_runtime(namespace, 'AT+CPIN?')[0])
        self.assertEqual(modem.writes, [b'AT+CPIN?\r\n'])


if __name__ == '__main__':
    unittest.main()
