from contextlib import contextmanager
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import socket
import ssl
import threading
from types import SimpleNamespace
import unittest
import urllib.error
from unittest.mock import patch

from sms_core import updates
from sms_core.threading_runtime import WorkerThreadRegistry
from sms_ui.update_check_runtime import check_update_and_prompt_runtime, handle_update_check_plan


PRIVATE = 'synthetic-diagnostic-private'


@contextmanager
def rejected_server():
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_args):
            pass
        def do_GET(self):
            self.send_response(503, f'upstream rejected token={PRIVATE}')
            self.end_headers()
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    server.daemon_threads = True
    worker = threading.Thread(target=lambda: server.serve_forever(poll_interval=0.01))
    worker.start()
    try:
        yield f'http://127.0.0.1:{server.server_port}'
    finally:
        server.shutdown()
        server.server_close()
        worker.join(2)


class UpdateDiagnosticSafetyTests(unittest.TestCase):
    def test_real_http_reason_does_not_reach_update_error_ui_or_registered_log(self):
        shown, logs = [], []
        registry = WorkerThreadRegistry()
        with rejected_server() as base, patch.object(updates, 'build_release_api_urls', return_value=[base]), \
                patch.object(updates, '_sleep_backoff'):
            worker = check_update_and_prompt_runtime(
                owner='synthetic', repo='synthetic', current_version='4.0.2',
                get_update_config=lambda: ('', base), ui_post=lambda callback: callback(),
                show_info=lambda *_a: None, show_warning=lambda *_a: None,
                show_error=lambda *_a: shown.append(_a), ask_open_download=lambda *_a: False,
                open_url=lambda *_a: None, thread_registry=registry, log_error=logs.append)
            worker.join(3)
            self.assertFalse(worker.is_alive())
        self.assertEqual(shown, [('检测更新失败', '获取最新版本信息失败：HTTP 503 请求失败')])
        self.assertEqual(logs, [])
        self.assertEqual(registry.snapshot(), ())
        self.assertNotIn(PRIVATE, str(shown))

    def test_proxy_result_hides_origin_credentials_path_and_http_error_reason(self):
        with rejected_server() as base, patch.object(updates, '_sleep_backoff'):
            configured = base.replace('http://', f'http://synthetic:{PRIVATE}@') + f'/{PRIVATE}?token={PRIVATE}'
            result = updates.test_update_proxy_connectivity(
                'synthetic', 'synthetic', configured, '',
                get_json=lambda _url, **_kwargs: updates.http_get_json(base, timeout=0.3, retries=1))
        rendered = updates.format_proxy_test_result(result)
        self.assertNotIn(PRIVATE, rendered)
        self.assertNotIn('synthetic:', rendered)
        self.assertIn(base, rendered)
        self.assertIn('HTTP 503', rendered)

    def test_successful_proxy_tests_preserve_real_bases_for_requests(self):
        configured = f'https://synthetic:{PRIVATE}@proxy.example.invalid/{PRIVATE}/'
        downloaded = []
        release = {'tag_name': 'v4.0.3', 'assets': [
            {'name': 'sms.zip', 'size': 1, 'browser_download_url': 'https://example.invalid/sms.zip'}]}
        result = updates.test_update_proxy_connectivity(
            'synthetic', 'synthetic', configured, configured,
            get_json=lambda _url, **_kwargs: release,
            probe=lambda url, **_kwargs: downloaded.append(url) or (True, 'HTTP 200'))
        self.assertEqual(result['ok_bases'], [configured])
        self.assertEqual(downloaded, [configured + 'https://example.invalid/sms.zip'])
        self.assertTrue(result['download_ok'])
        self.assertNotIn(PRIVATE, updates.format_proxy_test_result(result))
        self.assertIn('API 代理可用', updates.format_proxy_test_result(result))

    def test_typed_errors_remain_useful_without_formatting_unknown_exception(self):
        class ExplosiveError(Exception):
            def __str__(self):
                raise AssertionError('Exception text must not be formatted')
            def __repr__(self):
                raise AssertionError('Exception repr must not be formatted')
        for error, expected in (
            (urllib.error.URLError(TimeoutError(PRIVATE)), '连接超时'),
            (urllib.error.URLError(socket.gaierror(PRIVATE)), 'DNS'),
            (urllib.error.URLError(ssl.SSLError(PRIVATE)), 'TLS握手失败'),
            (ssl.SSLCertVerificationError(PRIVATE), '证书'),
            (ExplosiveError(PRIVATE), 'ExplosiveError'),
            (ValueError([PRIVATE]), 'ValueError'),
        ):
            with self.subTest(error_type=type(error).__name__):
                displayed = updates.classify_update_error(error)
                self.assertIn(expected, displayed)
                self.assertNotIn(PRIVATE, displayed)

    def test_known_release_validation_message_is_preserved(self):
        message = '更新服务响应过大，已停止读取'
        self.assertEqual(updates.classify_update_error(ValueError(message)), message)
        self.assertNotIn(PRIVATE, str(updates.UpdateRequestError(RuntimeError(PRIVATE))))

    def test_browser_failure_hides_error_and_url_credentials_and_keeps_real_open_value(self):
        url = f'https://synthetic:{PRIVATE}@proxy.example.invalid/sms.zip?%74oken={PRIVATE}&file=normal#fragment'
        shown, opened = [], []
        def fail_open(value):
            opened.append(value)
            raise RuntimeError(PRIVATE)
        handle_update_check_plan(SimpleNamespace(kind='update', latest_tag='v4.0.3', download_url=url), '4.0.2',
            show_info=lambda *_a: None, show_warning=lambda *_a: shown.append(_a),
            ask_open_download=lambda *_a: True, open_url=fail_open)
        self.assertEqual(opened, [url])
        self.assertNotIn(PRIVATE, str(shown))
        self.assertIn('/sms.zip?%74oken=***&file=normal', shown[0][1])
        self.assertIn('更新代理设置', shown[0][1])

    def test_normal_download_link_remains_copyable_after_browser_failure(self):
        url = 'https://example.invalid/releases/sms.zip?file=normal'
        shown = []
        handle_update_check_plan(SimpleNamespace(kind='update', latest_tag='v4.0.3', download_url=url), '4.0.2',
            show_info=lambda *_a: None, show_warning=lambda *_a: shown.append(_a),
            ask_open_download=lambda *_a: True, open_url=lambda _url: False)
        self.assertIn(url, shown[0][1])
        self.assertIn('请手动复制以下链接下载', shown[0][1])

    def test_config_read_failure_is_caught_without_exception_format_or_log_bypass(self):
        class ExplosiveError(Exception):
            def __str__(self):
                raise AssertionError('Original exception must not be formatted')
        shown, logs = [], []
        registry = WorkerThreadRegistry()
        def bad_config():
            raise ExplosiveError(PRIVATE)
        worker = check_update_and_prompt_runtime(
            owner='synthetic', repo='synthetic', current_version='4.0.2', get_update_config=bad_config,
            ui_post=lambda callback: callback(), show_info=lambda *_a: None, show_warning=lambda *_a: None,
            show_error=lambda *_a: shown.append(_a), ask_open_download=lambda *_a: False,
            open_url=lambda *_a: None, thread_registry=registry, log_error=logs.append)
        worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertEqual(len(shown), 1)
        self.assertIn('ExplosiveError', shown[0][1])
        self.assertEqual(logs, [])
        self.assertEqual(registry.snapshot(), ())


if __name__ == '__main__':
    unittest.main()
