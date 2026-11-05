import asyncio
import json
import threading
import time
import unittest

from sms_core.cloud_message_runtime import handle_cloud_message_runtime
from sms_core.cloud_security import safe_preview
from sms_core.cloud_ws_runtime import wait_cloud_login_ack_runtime


class CloudAuthDiagnosticSafetyTests(unittest.IsolatedAsyncioTestCase):
    async def test_login_and_reauthorization_logs_hide_credentials_without_changing_ack(self):
        private = 'synthetic-private-key'
        messages = [f'授权说明 password={private}',
            f'授权说明 secret="synthetic-prefix,{private}"',
            f'授权说明 https://user:{private}@example.invalid/private/{private}',
            f'授权说明 AT+CPIN="{private}"',
            f'授权说明 AT+CPNETAPN="1","apn","user","{private}"']
        for status in ('authorized', 'waiting', 'failed'):
            for message in messages:
                for initial in (True, False):
                    with self.subTest(status=status, initial=initial, message=messages.index(message)):
                        ack = dict(type='device_login_ack', auth_status=status, message=message)
                        raw = json.dumps(ack, ensure_ascii=False)
                        logs, authorization, original_acks = [], [], []
                        def log(text, **_kwargs):
                            logs.append(text)
                        if initial:
                            class WebSocket:
                                async def recv(self):
                                    return raw
                            await wait_cloud_login_ack_runtime(WebSocket(), stop_event=threading.Event(),
                                set_authorized=authorization.append, set_auth_status_from_ack=original_acks.append,
                                log=log, safe_preview=safe_preview, monotonic=time.monotonic)
                        else:
                            async def reply(_payload):
                                self.fail('Login ACK cannot send a command reply')
                            await handle_cloud_message_runtime(raw, is_authorized=lambda: False,
                                set_authorized=authorization.append, reply=reply, log=log,
                                set_auth_status_from_ack=original_acks.append, set_cloud_status=lambda *_args: None,
                                check_replay_window=lambda *_args: None, auth_matches=lambda _data: True,
                                time_text=lambda: '', timestamp=lambda: 0, status_payload=lambda: {},
                                send_serial_command=lambda *_args: None, show_window=lambda: None,
                                hide_window=lambda: None)
                        self.assertEqual(authorization, [status == 'authorized'])
                        self.assertEqual(original_acks, [ack])
                        self.assertNotIn(private, '\n'.join(logs))
                        self.assertTrue(logs)

    async def test_plain_auth_diagnostics_and_ordinary_sms_body_remain_visible(self):
        ack = dict(type='device_login_ack', auth_status='waiting', message='设备正在等待网页端绑定')
        self.assertIn(ack['message'], safe_preview(json.dumps(ack, ensure_ascii=False)))
        body = '正常短信正文 password=synthetic-body-value AT+CPIN="synthetic-pin-text"'
        sms = dict(type='sms', message=body)
        self.assertEqual(json.loads(safe_preview(json.dumps(sms, ensure_ascii=False)))['message'], body)
