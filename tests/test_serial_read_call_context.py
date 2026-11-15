import threading
import unittest
from contextlib import contextmanager
from types import SimpleNamespace

from sms_app.serial_namespace_bindings import install_serial_namespace_bindings
from sms_core.call_command_context import SerialCallStateSnapshot, call_state_lock, set_local_dial_number
from sms_core.call_send_namespace_runtime import send_local_serial_command_namespace_runtime
from sms_core.serial_app_runtime import run_serial_reader_namespace_runtime
from sms_core.serial_io_runtime import read_serial_line_safely_runtime
from sms_core.serial_runtime import SerialLineDecoder, SerialReadChunk
from tests import test_serial_dial_generation


class BufferedSerial:
    def __init__(self, data=b"", timeout=0.05):
        self.is_open = True
        self.timeout = timeout
        self.buffer = bytearray(data)
        self.read_sizes = []

    @property
    def in_waiting(self):
        return len(self.buffer)

    def read(self, size):
        self.read_sizes.append(size)
        data = bytes(self.buffer[:size])
        del self.buffer[:size]
        return data


class SerialReadCallContextTests(unittest.TestCase):
    def test_complete_old_terminal_keeps_origin_after_read_before_runtime_binding(self):
        for terminal in ("NO CARRIER", "ERROR", "[I]-[ril.proatc] NO CARRIER", "[I]-[ril.proatc] ERROR"):
            for number in ("10086", "10010"):
                with self.subTest(terminal=terminal, number=number):
                    self.exercise_old_terminal(terminal, number)

    def exercise_old_terminal(self, terminal, number):
        fixture = test_serial_dial_generation.SerialDialGenerationTests()
        fixture.setUp()
        ns = fixture.namespace
        ns["serial_lock"] = threading.Lock()
        ns["serial_read_lock"] = threading.Lock()
        ns["serial_stop_event"] = threading.Event()
        ns["serial"] = SimpleNamespace(SerialException=RuntimeError)
        workers = []
        wrote = threading.Event()

        def send_local(command, **kwargs):
            def start(_name, worker, **_kwargs):
                if fixture.pending:
                    fixture.workers.append(worker)
                    return object()
                return worker()
            return send_local_serial_command_namespace_runtime(
                ns, command, **kwargs, start_worker=start, response_timeout=1,
            )

        fixture.send_local = send_local

        def raw_loop(_kwargs, dial):
            original = dict(ns)
            install_serial_namespace_bindings(ns)
            real_read = ns["read_serial_line_safely"]
            ns.clear()
            ns.update(original)
            serial = ns["serial_obj"]
            rx = BufferedSerial((terminal + "\r\n").encode(), timeout=0)
            serial.timeout = 0
            serial.read = rx.read
            type(serial).in_waiting = property(lambda _serial: rx.in_waiting)
            first = True

            def read():
                nonlocal first
                raw = real_read()
                if first:
                    first = False
                    fixture.pending = True
                    serial.response = None
                    serial.after_write = wrote.set
                    dial(number)
                    worker = threading.Thread(target=fixture.workers.pop())
                    workers.append(worker)
                    worker.start()
                    self.assertTrue(wrote.wait(1))
                    rx.buffer.extend(b"OK\r\nCONNECT\r\n")
                elif not raw:
                    ns["serial_running"] = False
                return raw

            ns["read_serial_line_safely"] = read
            ns["open_and_initialize_serial"] = lambda _port: dial("10086")
            def disconnected(error, *_args):
                self.fail("fixture reader disconnected: " + str(error))

            try:
                run_serial_reader_namespace_runtime(
                    ns, parse_callback_head=lambda value: ("", value),
                    apply_disconnect_effects=disconnected,
                )
                for worker in workers:
                    worker.join(1)
                    self.assertFalse(worker.is_alive())
                fixture.drain_ui()
                self.assertEqual(ns["current_dial_num"], number)
                self.assertEqual(fixture.popups[-1].markers, ["connected"])
                self.assertEqual(serial.writes, [b"ATD10086;\r\n", ("ATD" + number + ";\r\n").encode()])
                self.assertEqual(ns["CLOUD_SMS_EVENT_Q"].qsize(), 2)
            finally:
                ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"].cancel_active("test cleanup")
                for worker in workers:
                    worker.join(2)
                del type(serial).in_waiting

        fixture.raw_loop = raw_loop
        fixture.run_scenario(lambda *_args: None)

    def read_contextual(self, namespace, *, wait_for_data=None, monotonic=None):
        options = dict(
            read_lock=namespace["serial_read_lock"],
            read_context_guard=lambda: call_state_lock(namespace),
            capture_read_chunk=lambda raw: SerialReadChunk(raw, SerialCallStateSnapshot(namespace)),
            wait_for_data=wait_for_data,
        )
        if monotonic is not None:
            options["monotonic"] = monotonic
        return read_serial_line_safely_runtime(
            namespace["serial_lock"], lambda: namespace["serial_obj"], RuntimeError, **options,
        )

    def namespace(self, serial):
        return dict(serial_lock=threading.Lock(), serial_read_lock=threading.Lock(),
                    serial_obj=serial, serial_connection_generation=1,
                    ring_timeout_target=0.0, current_dial_num="")

    def test_idle_wait_releases_locks_and_binds_actual_new_reply(self):
        serial = BufferedSerial(timeout=0.2)
        ns = self.namespace(serial)
        entered = threading.Event()
        accepted = threading.Event()
        errors = []

        def sender():
            try:
                self.assertTrue(entered.wait(1))
                with ns["serial_lock"], call_state_lock(ns):
                    request = set_local_dial_number(ns, "10086")
                    request.submitted = True
                    serial.buffer.extend(b"OK\r\n")
            except Exception as exc:
                errors.append(exc)
            finally:
                accepted.set()

        def wait(_seconds):
            entered.set()
            self.assertTrue(accepted.wait(1))
            return False

        worker = threading.Thread(target=sender)
        worker.start()
        try:
            raw = self.read_contextual(ns, wait_for_data=wait)
        finally:
            worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertEqual(errors, [])
        self.assertEqual(raw, b"OK\r\n")
        self.assertIs(raw.call_state_snapshot.generation, ns["_local_dial_request"])
        self.assertTrue(raw.call_state_snapshot.is_response_current())

    def test_partial_old_line_retains_first_atomic_read_snapshot(self):
        serial = BufferedSerial(b"NO ", timeout=0)
        ns = self.namespace(serial)
        decoder = SerialLineDecoder()
        first = self.read_contextual(ns)
        self.assertEqual(decoder.feed(first), [])
        set_local_dial_number(ns, "10010")
        serial.buffer.extend(b"CARRIER\r\nOK\r\n")
        second = self.read_contextual(ns)
        lines = decoder.feed(second)
        self.assertEqual(lines, ["NO CARRIER", "OK"])
        self.assertIs(lines[0].call_state_snapshot, first.call_state_snapshot)
        self.assertIs(lines[1].call_state_snapshot, second.call_state_snapshot)
        self.assertFalse(lines[0].call_state_snapshot.is_response_current())
        self.assertTrue(lines[1].call_state_snapshot.is_response_current())

    def test_replaced_serial_during_idle_wait_is_rejected(self):
        original = BufferedSerial(timeout=0.2)
        replacement = BufferedSerial(b"OK\r\n")
        ns = self.namespace(original)

        def wait(_seconds):
            with ns["serial_lock"]:
                ns["serial_obj"] = replacement
            return False

        with self.assertRaisesRegex(RuntimeError, "连接已变化"):
            self.read_contextual(ns, wait_for_data=wait)
        self.assertEqual(original.read_sizes, [])
        self.assertEqual(replacement.read_sizes, [])

    def test_stop_wait_returns_empty_without_reading_late_bytes(self):
        serial = BufferedSerial(timeout=0.2)
        ns = self.namespace(serial)

        def wait(_seconds):
            serial.buffer.extend(b"NO CARRIER\r\n")
            return True

        raw = self.read_contextual(ns, wait_for_data=wait)
        self.assertEqual(raw, b"")
        self.assertEqual(serial.read_sizes, [])
        self.assertIsInstance(raw.call_state_snapshot, SerialCallStateSnapshot)

    def test_idle_budget_uses_serial_timeout_without_os_empty_read(self):
        serial = BufferedSerial(timeout=0.03)
        ns = self.namespace(serial)
        now = [0.0]
        waits = []

        def wait(seconds):
            waits.append(seconds)
            now[0] += seconds
            return False

        raw = self.read_contextual(ns, wait_for_data=wait, monotonic=lambda: now[0])
        self.assertEqual(raw, b"")
        self.assertAlmostEqual(sum(waits), serial.timeout)
        self.assertEqual(serial.read_sizes, [])

    def test_buffered_read_and_capture_share_serial_then_call_locks(self):
        serial = BufferedSerial(b"A" * 5000, timeout=0)
        ns = self.namespace(serial)
        seen = []
        context_held = []

        @contextmanager
        def context_guard():
            context_held.append(True)
            try:
                yield
            finally:
                context_held.pop()

        def capture(raw):
            self.assertFalse(ns["serial_lock"].acquire(blocking=False))
            self.assertEqual(context_held, [True])
            seen.append(len(raw))
            return raw

        raw = read_serial_line_safely_runtime(
            ns["serial_lock"], lambda: serial, RuntimeError,
            read_lock=ns["serial_read_lock"], read_context_guard=context_guard,
            capture_read_chunk=capture,
        )
        self.assertEqual(len(raw), 4096)
        self.assertEqual(serial.read_sizes, [4096])
        self.assertEqual(seen, [4096])


if __name__ == "__main__":
    unittest.main()
