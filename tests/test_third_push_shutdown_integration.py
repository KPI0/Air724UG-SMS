import json
import queue
import threading
import unittest
from unittest.mock import patch

from sms_ui.third_push_namespace_runtime import (
    enqueue_third_push_namespace_runtime, third_push_worker_namespace_runtime,
)


class ThirdPushShutdownIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.requests, self.logs, self.ui, self.popups = [], [], [], []
        stop = threading.Event()
        stop.set()
        self.namespace = dict(
            THIRD_PUSH_Q=queue.Queue(), third_push_stop=stop, TK_SHUTDOWN=threading.Event(),
            THIRD_PUSH_ENABLED=True, THIRD_PUSH_SMS_ENABLED=True, THIRD_PUSH_CALL_ENABLED=True,
            THIRD_PUSH_TYPES=['custom_post'], LOG_PREFIX='COM_OLD', APP_VERSION='test',
            THIRD_PUSH_SETTINGS={'custom_post_url': 'https://example.invalid/push',
                                 'custom_post_body': '{"source":"{port}","body":"{msg}"}'},
            system_ui=lambda *args: self.ui.append(args), log_file_only=self.logs.append,
            show_third_push_test_result=lambda *args: self.popups.append(args),
        )

    def enqueue(self, body, **kwargs):
        return enqueue_third_push_namespace_runtime(
            self.namespace, body, template='{port}: {msg}', show_result=True, **kwargs,
        )

    def drain(self, ok=True):
        def request(url, method, headers, data, **kwargs):
            self.requests.append(json.loads(data))
            return ok, 200 if ok else 503, 'ok' if ok else 'synthetic unavailable'
        with patch('sms_core.third_push_sender.http_request', side_effect=request):
            third_push_worker_namespace_runtime(self.namespace)
        self.assertEqual(self.namespace['THIRD_PUSH_Q'].unfinished_tasks, 0)

    def test_queued_sms_and_call_keep_source_in_message_and_custom_post_body(self):
        self.assertTrue(self.enqueue('sms'))
        self.namespace['LOG_PREFIX'] = 'COM_NEW'
        self.assertTrue(self.enqueue('call', event_type='call'))
        self.namespace['LOG_PREFIX'] = 'system'
        self.drain()
        self.assertEqual(self.requests, [
            {'source': 'COM_OLD', 'body': 'COM_OLD: sms'},
            {'source': 'COM_NEW', 'body': 'COM_NEW: call'},
        ])

    def test_existing_explicit_port_variables_are_preserved(self):
        for variables in ({'port': 'COM_EXPLICIT'}, {'{port}': 'COM_EXPLICIT'}):
            self.enqueue('synthetic', variables=variables)
        self.drain()
        self.assertEqual(self.requests, [{'source': 'COM_EXPLICIT', 'body': 'COM_EXPLICIT: synthetic'}] * 2)

    def test_legacy_items_without_source_still_use_current_port(self):
        self.enqueue('legacy')
        self.namespace['THIRD_PUSH_Q'].queue[0]['variables'].pop('port')
        self.drain()
        self.assertEqual(self.requests[0]['source'], 'COM_OLD')

    def test_shutdown_failure_is_logged_without_ui_or_popup(self):
        self.enqueue('synthetic')
        self.namespace['TK_SHUTDOWN'].set()
        self.drain(ok=False)
        self.assertEqual(len(self.logs), 1)
        self.assertIn('失败', self.logs[0])
        self.assertIn('503', self.logs[0])
        self.assertEqual((self.ui, self.popups), ([], []))

    def test_shutdown_success_does_not_log_failure_or_open_popup(self):
        self.enqueue('synthetic')
        self.namespace['TK_SHUTDOWN'].set()
        self.drain()
        self.assertEqual((self.logs, self.ui, self.popups), ([], [], []))

    def test_bad_shutdown_item_is_logged_and_does_not_block_next_notification(self):
        self.namespace['THIRD_PUSH_Q'].put(None)
        self.enqueue('next')
        self.namespace['TK_SHUTDOWN'].set()
        self.drain()
        self.assertEqual(len(self.logs), 1)
        self.assertIn('未确认送达', self.logs[0])
        self.assertEqual(self.requests[0]['body'], 'COM_OLD: next')
