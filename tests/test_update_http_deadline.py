from contextlib import contextmanager
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import threading
import time
import unittest
import urllib.error
from unittest.mock import patch

from sms_core import updates


class UpdateHttpDeadlineTests(unittest.TestCase):
    @contextmanager
    def server(self):
        self.error_connections_closed = {'GET': threading.Event(), 'HEAD': threading.Event()}
        error_connections_closed = self.error_connections_closed

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *_args):
                pass

            def drip(self, data):
                try:
                    for byte in data:
                        self.wfile.write(bytes([byte]))
                        self.wfile.flush()
                        time.sleep(0.025)
                except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
                    pass

            def unfinished_error(self, method):
                self.send_response(503)
                self.send_header('Content-Length', '1000')
                self.end_headers()
                self.wfile.flush()
                self.connection.settimeout(2)
                try:
                    if not self.connection.recv(1):
                        error_connections_closed[method].set()
                except TimeoutError:
                    pass

            def do_HEAD(self):
                if self.path in ('/error', '/head-error'):
                    self.unfinished_error('HEAD')
                elif self.path == '/probe-headers':
                    self.drip(b'HTTP/1.0 200 OK\r\nX-Slow: ' + b'x' * 30 + b'\r\n\r\n')
                else:
                    self.send_error(405)

            def do_GET(self):
                body = b'{"tag_name":"v4.0.2","assets":[]}'
                if self.path == '/head-error' and not error_connections_closed['HEAD'].wait(0.5):
                    self.send_error(502)
                    return
                if self.path == '/error':
                    self.unfinished_error('GET')
                    return
                if self.path in ('/headers', '/probe-headers'):
                    self.drip(b'HTTP/1.0 200 OK\r\nX-Slow: ' + b'x' * 30 + b'\r\n\r\n')
                    return
                if self.path == '/redirect':
                    time.sleep(0.08)
                    self.send_response(302)
                    self.send_header('Location', '/body')
                    self.end_headers()
                    return
                self.send_response(200)
                if self.path == '/chunked':
                    self.send_header('Transfer-Encoding', 'chunked')
                    self.end_headers()
                    self.drip(b'%X\r\n' % len(body) + body + b'\r\n0\r\n\r\n')
                else:
                    self.send_header('Content-Length', str(len(body)))
                    self.end_headers()
                    if self.path in ('/body', '/probe-body'):
                        self.drip(body)
                    else:
                        self.wfile.write(body)

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

    def test_release_read_deadline_covers_headers_body_chunking_and_redirects(self):
        with self.server() as base:
            for path in ('headers', 'body', 'chunked', 'redirect'):
                with self.subTest(path=path), patch.object(updates, '_sleep_backoff'):
                    started = time.monotonic()
                    with self.assertRaises((TimeoutError, urllib.error.URLError)):
                        updates.http_get_json(f'{base}/{path}', timeout=0.12, retries=1)
                    self.assertLess(time.monotonic() - started, 0.5)

    def test_probe_deadline_covers_head_and_get_fallback(self):
        with self.server() as base:
            for path in ('probe-headers', 'probe-body'):
                with self.subTest(path=path), patch.object(updates, '_sleep_backoff'):
                    started = time.monotonic()
                    with self.assertRaises((TimeoutError, urllib.error.URLError)):
                        updates.http_probe(f'{base}/{path}', timeout=0.12, retries=1)
                    self.assertLess(time.monotonic() - started, 0.7)

    def test_normal_reads_and_head_fallback_work_without_system_proxies(self):
        with self.server() as base, patch('urllib.request.getproxies') as getproxies:
            self.assertEqual(updates.http_get_json(f'{base}/normal', timeout=1, retries=1),
                             {'tag_name': 'v4.0.2', 'assets': []})
            self.assertEqual(updates.http_probe(f'{base}/normal', timeout=1, retries=1),
                             (True, 'HTTP 200'))
            getproxies.assert_not_called()

    def test_retries_have_fresh_deadlines_and_no_final_backoff(self):
        with self.server() as base, patch.object(updates, '_sleep_backoff') as backoff:
            started = time.monotonic()
            with self.assertRaises((TimeoutError, urllib.error.URLError)):
                updates.http_get_json(f'{base}/body', timeout=0.12, retries=2)
            elapsed = time.monotonic() - started
            self.assertGreaterEqual(elapsed, 0.2)
            self.assertLess(elapsed, 0.7)
            backoff.assert_called_once_with(0.6, 0)

    def test_failed_release_releases_socket_even_when_error_is_retained(self):
        with self.server() as base:
            with self.assertRaises(urllib.error.HTTPError) as caught:
                updates.http_get_json(f'{base}/error', timeout=0.3, retries=1)
            try:
                self.assertTrue(caught.exception.closed)
                self.assertTrue(self.error_connections_closed['GET'].wait(0.5))
            finally:
                caught.exception.close()

    def test_failed_head_releases_socket_before_successful_get_fallback(self):
        with self.server() as base:
            self.assertEqual(updates.http_probe(f'{base}/head-error', timeout=0.3, retries=1),
                             (True, 'HTTP 200'))
            self.assertTrue(self.error_connections_closed['HEAD'].wait(0.5))

    def test_failed_probe_releases_both_head_and_get_sockets(self):
        with self.server() as base:
            with self.assertRaises(urllib.error.HTTPError) as caught:
                updates.http_probe(f'{base}/error', timeout=0.3, retries=1)
            try:
                self.assertTrue(caught.exception.closed)
                self.assertTrue(self.error_connections_closed['HEAD'].wait(0.5))
                self.assertTrue(self.error_connections_closed['GET'].wait(0.5))
            finally:
                caught.exception.close()


if __name__ == '__main__':
    unittest.main()
