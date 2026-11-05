from contextlib import contextmanager
import queue
import socket
import socketserver
import threading
import time
import unittest
from unittest.mock import patch

from sms_core import third_push_sender as sender
from sms_core.smtp_deadline import DeadlineSMTP, DeadlineSMTP_SSL
from sms_core import http_deadline
from sms_ui.third_push_app_runtime import third_push_worker_app_runtime


class PlainContext:
    """Protocol fixture only; actual TLS is covered by the isolated integration."""
    def __init__(self, delay=0):
        self.calls = []
        self.delay = delay

    def wrap_socket(self, sock, server_hostname):
        self.calls.append((server_hostname, sock.gettimeout()))
        if self.delay:
            time.sleep(self.delay)
        return sock


@contextmanager
def smtp_server(mode='normal', *, context=None, ssl_immediate=False, delay=0):
    state = {'recipients': 0, 'delivered': 0, 'tls': 0, 'closed': threading.Event()}
    stopped = threading.Event()

    class Handler(socketserver.BaseRequestHandler):
        def handle(self):
            conn = self.request
            conn.settimeout(2)
            conn.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
            reader = None
            try:
                if ssl_immediate:
                    conn = context.wrap_socket(conn, server_side=True)
                    state['tls'] += 1
                def reply(data, slow=False):
                    if delay and stopped.wait(delay):
                        return
                    if slow:
                        for byte in data:
                            if stopped.wait(0.025):
                                return
                            conn.sendall(bytes([byte]))
                    else:
                        conn.sendall(data)
                if mode == 'greeting':
                    reply(b'220 ' + b'x' * 60 + b'\r\n', slow=True)
                elif mode == 'stalled_greeting':
                    reply(b'220 pending')
                    stopped.wait(0.4)
                    return
                elif mode == 'multiline':
                    for _ in range(60):
                        reply(b'220-another line\r\n')
                        if stopped.wait(0.025):
                            return
                else:
                    reply(b'220 loopback fixture\r\n')
                reader = conn.makefile('rb')
                while not stopped.is_set():
                    line = reader.readline(65537)
                    if not line:
                        break
                    command = line.split(b' ', 1)[0].strip().upper()
                    if command in (b'EHLO', b'HELO'):
                        reply(b'250-loopback fixture\r\n250-STARTTLS\r\n250 AUTH PLAIN\r\n')
                    elif command == b'STARTTLS':
                        reply(b'220 ready for TLS\r\n')
                        if context is not None:
                            reader.close()
                            conn = context.wrap_socket(conn, server_side=True)
                            reader = conn.makefile('rb')
                            state['tls'] += 1
                    elif command == b'AUTH':
                        reply(b'235 authentication accepted\r\n')
                    elif command == b'MAIL':
                        reply(b'250 sender accepted\r\n')
                    elif command == b'RCPT':
                        state['recipients'] += 1
                        reply(b'550 refused\r\n' if mode == 'refuse' and state['recipients'] == 2
                              else b'250 recipient accepted\r\n')
                    elif command == b'DATA':
                        reply(b'354 send message\r\n')
                        while reader.readline(65537) not in (b'.\r\n', b''):
                            pass
                        state['delivered'] += 1
                        reply(b'250 delivered\r\n')
                    elif command == b'QUIT':
                        reply(b'221 ' + b'x' * 60 + b'\r\n', slow=mode == 'quit')
                        break
                    elif command == b'RSET':
                        reply(b'250 reset\r\n')
                    else:
                        reply(b'500 unknown command\r\n')
            except (OSError, TimeoutError):
                pass
            finally:
                if reader is not None:
                    reader.close()
                conn.close()
                state['closed'].set()

    class Server(socketserver.ThreadingTCPServer):
        allow_reuse_address = True
        daemon_threads = True

    server = Server(('127.0.0.1', 0), Handler)
    worker = threading.Thread(target=lambda: server.serve_forever(poll_interval=0.01))
    worker.start()
    try:
        yield server.server_address[1], state
    finally:
        stopped.set()
        server.shutdown()
        server.server_close()
        worker.join(2)


def email_settings(port, *, encryption='starttls', recipients=1):
    return {
        'email_smtp_host': '127.0.0.1', 'email_smtp_port': str(port),
        'email_encryption': encryption, 'email_username': 'synthetic@example.invalid',
        'email_password': 'synthetic-valid', 'email_from_address': 'sender@example.invalid',
        'email_to_addresses': ';'.join(f'receiver{index}@example.invalid' for index in range(recipients)),
    }


@contextmanager
def bounded_smtp(timeout=0.15, context=None):
    context = context or PlainContext()
    def plain(*args, **kwargs):
        kwargs['timeout'] = timeout
        return DeadlineSMTP(*args, **kwargs)
    def secure(*args, **kwargs):
        kwargs['timeout'] = timeout
        return DeadlineSMTP_SSL(*args, **kwargs)
    with patch.object(sender, 'DeadlineSMTP', plain), patch.object(sender, 'DeadlineSMTP_SSL', secure), \
            patch.object(sender.ssl, 'create_default_context', return_value=context):
        yield context


class SmtpDeadlineTests(unittest.TestCase):
    def test_multiple_tcp_addresses_share_deadline_and_close_every_failed_socket(self):
        for connection in ('smtp', 'smtp_ssl', 'http'):
            with self.subTest(connection=connection):
                opened = []
                class SlowSocket:
                    def __init__(self, *_args):
                        self.timeout = None
                        self.closed = False
                        self.bound = None
                        self.connected = False
                        opened.append(self)
                    def settimeout(self, timeout):
                        self.timeout = timeout
                    def bind(self, address):
                        self.bound = address
                    def connect(self, _address):
                        self.connected = True
                        time.sleep(min(self.timeout, 0.045))
                        raise TimeoutError('Synthetic connect refused after waiting')
                    def close(self):
                        self.closed = True
                addresses = [(socket.AF_INET, socket.SOCK_STREAM, 0, '', ('127.0.0.1', 1))] * 3
                with patch.object(http_deadline.socket, 'getaddrinfo', return_value=addresses), \
                        patch.object(http_deadline.socket, 'socket', SlowSocket):
                    start = time.monotonic()
                    with self.assertRaises(TimeoutError):
                        if connection == 'smtp':
                            DeadlineSMTP('localhost', 1, timeout=0.07, source_address=('127.0.0.1', 0))
                        elif connection == 'smtp_ssl':
                            DeadlineSMTP_SSL('localhost', 1, timeout=0.07, context=PlainContext(),
                                             source_address=('127.0.0.1', 0))
                        else:
                            http_deadline._HTTPConnection('localhost', deadline=http_deadline._Deadline(0.07),
                                                          source_address=('127.0.0.1', 0)).connect()
                    self.assertLess(time.monotonic() - start, 0.115)
                self.assertGreaterEqual(sum(sock.connected for sock in opened), 2)
                self.assertTrue(all(sock.closed for sock in opened))
                self.assertTrue(all(sock.bound == ('127.0.0.1', 0) for sock in opened))
                self.assertLess(opened[1].timeout, opened[0].timeout)
                self.assertTrue(all(right.timeout < left.timeout for left, right in zip(opened, opened[1:])))

    def test_failed_address_can_fall_back_to_real_loopback_with_source_binding(self):
        with socket.socket() as listener:
            listener.bind(('127.0.0.1', 0))
            listener.listen(1)
            with socket.socket() as unavailable:
                unavailable.bind(('127.0.0.1', 0))
                refused = unavailable.getsockname()
                addresses = [(socket.AF_INET, socket.SOCK_STREAM, 0, '', refused),
                             (socket.AF_INET, socket.SOCK_STREAM, 0, '', listener.getsockname())]
                with patch.object(http_deadline.socket, 'getaddrinfo', return_value=addresses):
                    start = time.monotonic()
                    client = http_deadline.create_connection_with_deadline(
                        ('localhost', listener.getsockname()[1]), http_deadline._Deadline(5), ('127.0.0.1', 0))
                    self.assertLess(time.monotonic() - start, 5)
                peer, _ = listener.accept()
                with client, peer:
                    self.assertEqual(client.getpeername(), listener.getsockname())
                    self.assertEqual(client.getsockname()[0], '127.0.0.1')
                    client.sendall(b'synthetic')
                    self.assertEqual(peer.recv(9), b'synthetic')

    def test_slow_partial_greeting_stops_real_send_within_whole_budget(self):
        with smtp_server('greeting') as (port, state), bounded_smtp():
            start = time.monotonic()
            ok, info = sender.send_email('synthetic content', email_settings(port))
            self.assertFalse(ok)
            self.assertIn('超时', info)
            self.assertLess(time.monotonic() - start, 0.6)
            self.assertTrue(state['closed'].wait(0.5))
            self.assertEqual(state['delivered'], 0)
        native_clock = time.monotonic
        with self.subTest(native_timeout_while_deadline_clock_still_positive=True), \
                smtp_server('stalled_greeting') as (port, state), bounded_smtp():
            held_clock = native_clock()
            with patch.object(http_deadline.time, 'monotonic', return_value=held_clock):
                ok, info = sender.send_email('synthetic content', email_settings(port))
            self.assertFalse(ok)
            self.assertIn('超时', info)
            self.assertGreaterEqual(native_clock() - held_clock, 0.12)
            self.assertLess(native_clock() - held_clock, 0.6)
            self.assertTrue(state['closed'].wait(0.5))
            self.assertEqual(state['delivered'], 0)

    def test_multiline_ssl_greeting_uses_one_deadline(self):
        with smtp_server('multiline') as (port, state), bounded_smtp() as context:
            start = time.monotonic()
            ok, info = sender.send_email('synthetic content', email_settings(port, encryption='ssl'))
            self.assertFalse(ok)
            self.assertIn('超时', info)
            self.assertLess(time.monotonic() - start, 0.6)
            self.assertTrue(context.calls)
            self.assertEqual(state['delivered'], 0)

    def test_deadline_is_shared_across_greeting_ehlo_and_starttls(self):
        with smtp_server(delay=0.07) as (port, state), bounded_smtp(timeout=0.18):
            start = time.monotonic()
            ok, info = sender.send_email('synthetic content', email_settings(port))
            self.assertFalse(ok)
            self.assertIn('超时', info)
            self.assertLess(time.monotonic() - start, 0.6)
            self.assertEqual(state['delivered'], 0)

    def test_starttls_wrap_cannot_reset_session_budget(self):
        context = PlainContext(delay=0.15)
        with smtp_server() as (port, state), bounded_smtp(timeout=0.12, context=context):
            start = time.monotonic()
            ok, info = sender.send_email('synthetic content', email_settings(port))
            self.assertFalse(ok)
            self.assertIn('超时', info)
            self.assertLess(time.monotonic() - start, 0.6)
            self.assertEqual(state['delivered'], 0)
            self.assertLessEqual(context.calls[0][1], 0.12 + 1e-8)

    def test_normal_starttls_and_ssl_keep_one_hundred_recipients(self):
        for encryption in ('starttls', 'ssl'):
            with self.subTest(encryption=encryption), smtp_server() as (port, state), bounded_smtp(timeout=5):
                ok, info = sender.send_email('synthetic content', email_settings(port, encryption=encryption, recipients=100))
                self.assertTrue(ok)
                self.assertEqual(info, 'SMTP 邮件已发送')
                self.assertEqual(state['recipients'], 100)
                self.assertEqual(state['delivered'], 1)

    def test_partial_recipient_refusal_retains_actual_delivery_result(self):
        with smtp_server('refuse') as (port, state), bounded_smtp(timeout=2):
            ok, info = sender.send_email('synthetic content', email_settings(port, recipients=3))
            self.assertFalse(ok)
            self.assertEqual(info, 'SMTP 接收方已接受 2，拒收 1')
            self.assertEqual(state['delivered'], 1)

    def test_confirmed_delivery_remains_success_when_quit_response_times_out(self):
        with smtp_server('quit') as (port, state), bounded_smtp(timeout=0.2):
            start = time.monotonic()
            ok, info = sender.send_email('synthetic content', email_settings(port))
            self.assertTrue(ok)
            self.assertEqual(info, 'SMTP 邮件已发送')
            self.assertEqual(state['delivered'], 1)
            self.assertLess(time.monotonic() - start, 0.65)
            self.assertTrue(state['closed'].wait(0.5))

    def test_real_sendall_uses_remaining_budget_and_preserves_successful_write(self):
        with socket.socket() as listener:
            listener.setsockopt(socket.SOL_SOCKET, socket.SO_RCVBUF, 4096)
            listener.bind(('127.0.0.1', 0))
            listener.listen(1)
            client_sock = socket.create_connection(listener.getsockname())
            peer, _ = listener.accept()
            with peer:
                client_sock.setsockopt(socket.SOL_SOCKET, socket.SO_SNDBUF, 4096)
                client = DeadlineSMTP(timeout=1)
                client.sock = client_sock
                try:
                    start = time.monotonic()
                    payload = b'x' * 65536
                    client.send(payload)
                    self.assertLess(time.monotonic() - start, 0.6)
                    self.assertGreater(client_sock.gettimeout(), 0)
                    self.assertLessEqual(client_sock.gettimeout(), 1)
                    received = bytearray()
                    peer.settimeout(1)
                    while len(received) < len(payload):
                        received.extend(peer.recv(len(payload) - len(received)))
                    self.assertEqual(bytes(received), payload)
                finally:
                    client.close()
                self.assertIsNone(client.sock)
        class FailingSocket:
            closed = False
            def settimeout(self, _timeout):
                pass
            def sendall(self, _data):
                raise TimeoutError('synthetic-private-adapter-timeout')
            def close(self):
                self.closed = True
        with self.subTest(stdlib_preserves_typed_write_timeout=True):
            client = DeadlineSMTP(timeout=1)
            failing_socket = FailingSocket()
            client.sock = failing_socket
            try:
                with self.assertRaises(TimeoutError) as caught:
                    client.send(b'synthetic')
                self.assertNotIn('synthetic-private-adapter-timeout', str(caught.exception))
                self.assertTrue(failing_socket.closed)
                self.assertIsNone(client.sock)
            finally:
                client.close()

    def test_shutdown_worker_drains_next_notice_after_real_timeout_without_ui(self):
        notices = queue.Queue()
        stop = threading.Event()
        stop.set()
        logs, ui, results = [], [], []
        with smtp_server('greeting') as (slow_port, slow_state), smtp_server() as (normal_port, normal_state), \
                bounded_smtp(timeout=0.2):
            for port in (slow_port, normal_port):
                notices.put({'message': 'synthetic content', 'channels': ['email'],
                             'settings': email_settings(port), 'show_result': True})
            start = time.monotonic()
            worker = threading.Thread(target=lambda: third_push_worker_app_runtime(
                stop_event=stop, push_queue=notices, get_log_prefix=lambda: 'synthetic',
                app_version='4.0.2', system_ui=lambda *args: ui.append(args),
                show_result=lambda *args: results.append(args), log_error=logs.append))
            worker.start()
            worker.join(1.5)
            self.assertFalse(worker.is_alive())
            self.assertLess(time.monotonic() - start, 0.8)
            self.assertEqual(notices.unfinished_tasks, 0)
            self.assertEqual(slow_state['delivered'], 0)
            self.assertEqual(normal_state['delivered'], 1)
            self.assertEqual(ui, [])
            self.assertEqual(results, [])
            self.assertEqual(len(logs), 1)
            self.assertNotIn('synthetic-valid', logs[0])


if __name__ == '__main__':
    unittest.main()
