import threading
import unittest
from unittest.mock import patch

from sms_core.serial_sender import (
    AtCommandResponseWaiter,
    AtCommandResponseCoordinator,
    SerialCommandResult,
    SmsPduSendCoordinator,
    SmsPduSendResponse,
    send_command_async,
    send_command_sequence_async,
    send_command_with_result_async,
    send_text_sms_pdu_async,
    write_serial_command_sequence_locked,
    write_serial_command_sequence_confirmed_locked,
    write_text_sms_pdu,
    write_text_sms_pdu_locked,
    write_serial_command_result,
)
from sms_core.threading_runtime import WorkerThreadRegistry


class FakeSerial:
    def __init__(self, is_open=True, fail_write=False):
        self.is_open = is_open
        self.fail_write = fail_write
        self.writes = []
        self.flush_count = 0

    def write(self, payload):
        if self.fail_write:
            raise RuntimeError("write failed")
        self.writes.append(payload)

    def flush(self):
        self.flush_count += 1


class CallbackSerial(FakeSerial):
    def __init__(self, on_write):
        super().__init__()
        self.on_write = on_write

    def write(self, payload):
        super().write(payload)
        self.on_write(payload)


class TrackingLock:
    def __init__(self):
        self.depth = 0

    def __enter__(self):
        self.depth += 1
        return self

    def __exit__(self, exc_type, exc, tb):
        self.depth -= 1

    @property
    def locked(self):
        return self.depth > 0


class FakeSmsWaiter:
    def __init__(self, response=None, prompt_response=None):
        self.response = response or SmsPduSendResponse(True)
        self.prompt_response = prompt_response or SmsPduSendResponse(True)
        self.waits = []
        self.prompt_waits = []
        self.cancelled = []

    def wait_prompt(self, timeout):
        self.prompt_waits.append(timeout)
        return self.prompt_response

    def mark_pdu_send_started(self):
        return True

    def wait(self, timeout):
        self.waits.append(timeout)
        return self.response

    def cancel(self, error):
        self.cancelled.append(error)


class FakeSmsCoordinator:
    def __init__(self, responses, prompt_responses=None):
        self.responses = list(responses)
        self.prompt_responses = list(prompt_responses or [])
        self.waiters = []
        self.finished = []

    def begin_segment(self, label=""):
        response = self.responses.pop(0) if self.responses else SmsPduSendResponse(True)
        prompt_response = (
            self.prompt_responses.pop(0)
            if self.prompt_responses
            else SmsPduSendResponse(True)
        )
        waiter = FakeSmsWaiter(response, prompt_response)
        waiter.label = label
        self.waiters.append(waiter)
        return waiter

    def finish(self, waiter):
        self.finished.append(waiter)


class SerialSenderResultTests(unittest.TestCase):
    def test_write_serial_command_result_reports_success(self):
        serial_obj = FakeSerial()

        result = write_serial_command_result(serial_obj, "ATA")

        self.assertEqual(result, SerialCommandResult(True))
        self.assertEqual(serial_obj.writes, [b"ATA\r\n"])
        self.assertEqual(serial_obj.flush_count, 1)

    def test_write_serial_command_result_reports_closed_port(self):
        result = write_serial_command_result(FakeSerial(is_open=False), "ATA")

        self.assertFalse(result.ok)
        self.assertEqual(result.error, "串口未连接")

    def test_write_serial_command_result_reports_write_error(self):
        result = write_serial_command_result(FakeSerial(fail_write=True), "ATA")

        self.assertFalse(result.ok)
        self.assertEqual(result.error, "write failed")

    def test_at_response_waiter_recognizes_terminal_voice_dial_response(self):
        waiter = AtCommandResponseWaiter("ATD10086;")
        waiter.observe_line("NO CARRIER")
        response = waiter.wait(0)
        self.assertFalse(response.ok)
        self.assertEqual(response.error, "NO CARRIER")

    def test_at_response_waiter_does_not_treat_call_terminal_as_result_for_other_command(self):
        waiter = AtCommandResponseWaiter("AT+CSQ")
        waiter.observe_line("NO CARRIER")
        self.assertFalse(waiter.done())

    def test_send_command_with_result_async_invokes_callback(self):
        serial_obj = FakeSerial()
        lock = threading.RLock()
        done = threading.Event()
        results = []

        thread = send_command_with_result_async(
            lock,
            lambda: serial_obj,
            "ATH",
            on_result=lambda result: (results.append(result), done.set()),
        )

        thread.join(1)
        self.assertTrue(done.is_set())
        self.assertEqual(results, [SerialCommandResult(True)])
        self.assertEqual(serial_obj.writes, [b"ATH\r\n"])

    def test_write_text_sms_pdu_sends_long_message_segments(self):
        serial_obj = FakeSerial()
        debug = []

        result = write_text_sms_pdu(
            serial_obj,
            "+1234",
            "A" * 100,
            push_debug=debug.append,
            sleep_func=lambda _seconds: None,
            response_coordinator=FakeSmsCoordinator([SmsPduSendResponse(True), SmsPduSendResponse(True)]),
        )

        self.assertTrue(result)
        self.assertEqual(serial_obj.writes[0], b"AT+CMGF=0\r\n")
        cmgs_commands = [
            payload for payload in serial_obj.writes
            if payload.startswith(b"AT+CMGS=")
        ]
        pdu_payloads = [
            payload for payload in serial_obj.writes
            if payload.endswith(b"\x1a")
        ]
        self.assertEqual(len(cmgs_commands), 2)
        self.assertEqual(len(pdu_payloads), 2)
        self.assertTrue(any("(1/2)" in line for line in debug))
        self.assertTrue(any("(2/2)" in line for line in debug))

    def test_write_text_sms_pdu_without_response_coordinator_fails_before_pdu(self):
        serial_obj = FakeSerial()
        debug = []

        result = write_text_sms_pdu(
            serial_obj,
            "+1234",
            "hello",
            push_debug=debug.append,
            sleep_func=lambda _seconds: None,
        )

        self.assertFalse(result)
        self.assertEqual(serial_obj.writes, [])
        self.assertTrue(any("未配置短信发送确认器" in line for line in debug))

    def test_write_text_sms_pdu_does_not_send_pdu_without_prompt(self):
        serial_obj = FakeSerial()
        debug = []

        result = write_text_sms_pdu(
            serial_obj,
            "+1234",
            "hello",
            push_debug=debug.append,
            sleep_func=lambda _seconds: None,
            response_coordinator=SmsPduSendCoordinator(),
            segment_timeout=0,
        )

        self.assertFalse(result)
        self.assertFalse(any(payload.endswith(b"\x1a") for payload in serial_obj.writes))
        self.assertTrue(any("等待 Modem 的 > 提示符超时" in line for line in debug))

    def test_write_text_sms_pdu_waits_for_prompt_before_writing_body(self):
        coordinator = SmsPduSendCoordinator()

        def on_write(payload):
            if payload.startswith(b"AT+CMGS="):
                coordinator.observe_line("> ")
            elif payload.endswith(b"\x1a"):
                coordinator.observe_line("+CMGS: 17")
                coordinator.observe_line("OK")

        serial_obj = CallbackSerial(on_write)

        result = write_text_sms_pdu(
            serial_obj,
            "+1234",
            "hello",
            sleep_func=lambda _seconds: None,
            response_coordinator=coordinator,
            prompt_timeout=0.1,
            segment_timeout=0.1,
        )

        self.assertTrue(result)
        self.assertTrue(serial_obj.writes[1].startswith(b"AT+CMGS="))
        self.assertTrue(serial_obj.writes[2].endswith(b"\x1a"))

    def test_write_text_sms_pdu_blocks_body_until_prompt_arrives(self):
        coordinator = SmsPduSendCoordinator()
        cmgs_written = threading.Event()
        pdu_written = threading.Event()

        def on_write(payload):
            if payload.startswith(b"AT+CMGS="):
                cmgs_written.set()
            elif payload.endswith(b"\x1a"):
                pdu_written.set()

        serial_obj = CallbackSerial(on_write)
        result_holder = []

        thread = threading.Thread(
            target=lambda: result_holder.append(
                write_text_sms_pdu(
                    serial_obj,
                    "+1234",
                    "hello",
                    sleep_func=lambda _seconds: None,
                    response_coordinator=coordinator,
                    prompt_timeout=1.0,
                    segment_timeout=1.0,
                )
            ),
            daemon=True,
        )
        thread.start()

        self.assertTrue(cmgs_written.wait(1))
        self.assertFalse(pdu_written.is_set())
        self.assertEqual(len(serial_obj.writes), 2)

        coordinator.observe_line(">")
        self.assertTrue(pdu_written.wait(1))
        coordinator.observe_line("+CMGS: 17")
        coordinator.observe_line("OK")
        thread.join(1)

        self.assertFalse(thread.is_alive())
        self.assertEqual(result_holder, [True])

    def test_write_text_sms_pdu_stops_when_cmgs_fails_before_prompt(self):
        coordinator = SmsPduSendCoordinator()

        def on_write(payload):
            if payload.startswith(b"AT+CMGS="):
                coordinator.observe_line("+CMS ERROR: 500")

        serial_obj = CallbackSerial(on_write)
        debug = []

        result = write_text_sms_pdu(
            serial_obj,
            "+1234",
            "hello",
            push_debug=debug.append,
            sleep_func=lambda _seconds: None,
            response_coordinator=coordinator,
            prompt_timeout=0.1,
        )

        self.assertFalse(result)
        self.assertFalse(any(payload.endswith(b"\x1a") for payload in serial_obj.writes))
        self.assertTrue(any("+CMS ERROR: 500" in line for line in debug))

    def test_write_text_sms_pdu_does_not_send_after_prompt_then_error(self):
        coordinator = SmsPduSendCoordinator()

        def on_write(payload):
            if payload.startswith(b"AT+CMGS="):
                coordinator.observe_line(">")
                coordinator.observe_line("+CMS ERROR: 500")

        serial_obj = CallbackSerial(on_write)
        debug = []

        result = write_text_sms_pdu(
            serial_obj,
            "+1234",
            "hello",
            push_debug=debug.append,
            sleep_func=lambda _seconds: None,
            response_coordinator=coordinator,
            prompt_timeout=1.0,
        )

        self.assertFalse(result)
        self.assertFalse(any(payload.endswith(b"\x1a") for payload in serial_obj.writes))
        self.assertTrue(any("+CMS ERROR: 500" in line for line in debug))

    def test_sms_pdu_send_coordinator_wakes_prompt_wait_on_error(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(label="(1/1)")
        started = threading.Event()
        result_holder = []

        def wait_for_prompt():
            started.set()
            result_holder.append(waiter.wait_prompt(1.0))

        thread = threading.Thread(target=wait_for_prompt, daemon=True)
        thread.start()
        self.assertTrue(started.wait(1))
        coordinator.observe_line("+CMS ERROR: 500")
        thread.join(0.2)
        coordinator.finish(waiter)

        self.assertFalse(thread.is_alive())
        self.assertEqual(result_holder[0].error, "+CMS ERROR: 500")

    def test_sms_pdu_send_coordinator_recognizes_prompt(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(label="(1/1)")

        coordinator.observe_line("[I]-[ril.proatc] > ")
        response = waiter.wait_prompt(0)
        coordinator.finish(waiter)

        self.assertTrue(response.ok)

    def test_sms_pdu_send_coordinator_ignores_untrusted_prompt_logs(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(label="(1/1)")

        coordinator.observe_line("[I]-[app.note] >")
        coordinator.observe_line("[I]-[ril.note] >")
        self.assertFalse(waiter.wait_prompt(0).ok)

        coordinator.observe_line("[I]-[ril.proatc] >")
        self.assertTrue(waiter.wait_prompt(0).ok)
        coordinator.finish(waiter)

    def test_waiter_atomic_pdu_commit_rejects_cancel_after_prompt(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(label="(1/1)")
        writes = []

        coordinator.observe_line(">")
        self.assertTrue(waiter.wait_prompt(0).ok)
        coordinator.cancel_active("串口连接已断开")
        result = waiter.write_pdu_if_ready(
            lambda: writes.append(b"PDU") or SerialCommandResult(True)
        )
        coordinator.finish(waiter)

        self.assertFalse(result.ok)
        self.assertEqual(result.error, "串口连接已断开")
        self.assertEqual(writes, [])

    def test_waiter_atomic_pdu_commit_blocks_concurrent_cancel_until_write_commits(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(label="(1/1)")
        cancel_done = threading.Event()
        cancel_thread_holder = []

        coordinator.observe_line(">")
        self.assertTrue(waiter.wait_prompt(0).ok)

        def writer():
            thread = threading.Thread(
                target=lambda: (
                    coordinator.cancel_active("串口连接已断开"),
                    cancel_done.set(),
                ),
                daemon=True,
            )
            cancel_thread_holder.append(thread)
            thread.start()
            self.assertFalse(cancel_done.wait(0.05))
            return SerialCommandResult(True)

        result = waiter.write_pdu_if_ready(writer)
        cancel_thread_holder[0].join(1)
        coordinator.finish(waiter)

        self.assertTrue(result.ok)
        self.assertTrue(cancel_done.is_set())

    def test_sms_pdu_send_coordinator_waits_for_cmgs_and_ok(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(label="(1/1)")

        coordinator.observe_line(">")
        self.assertTrue(waiter.wait_prompt(0).ok)
        self.assertTrue(waiter.mark_pdu_send_started())
        coordinator.observe_line("+CMGS: 17")
        self.assertFalse(waiter.done())
        coordinator.observe_line("[I]-[ril.proatc] OK")
        response = waiter.wait(0)
        coordinator.finish(waiter)

        self.assertTrue(response.ok)

    def test_sms_pdu_send_coordinator_ignores_unrelated_lines_by_phase(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(label="(1/1)")

        coordinator.observe_line("OK")
        coordinator.observe_line("[I]-[app] previous +CMGS: 99 result")
        coordinator.observe_line("[I]-[app] documentation says +CMS ERROR: 500")
        coordinator.observe_line("[E]-[usbmsc.write] mount ERROR")
        self.assertFalse(waiter.done())

        coordinator.observe_line(">")
        self.assertTrue(waiter.wait_prompt(0).ok)
        self.assertTrue(waiter.mark_pdu_send_started())
        coordinator.observe_line("OK")
        coordinator.observe_line("[I]-[app] unrelated +CMGS: 99")
        coordinator.observe_line("[E]-[socket] connect false ERROR")
        self.assertFalse(waiter.done())

        coordinator.observe_line("+CMGS: 17")
        coordinator.observe_line("OK")
        response = waiter.wait(0)
        coordinator.finish(waiter)

        self.assertTrue(response.ok)

    def test_sms_pdu_send_coordinator_ignores_ril_and_lib_sms_history_errors(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(cmgs_len=12)

        coordinator.observe_line("[I]-[ril.note] previous +CMS ERROR: 500")
        coordinator.observe_line("[I]-[lib_sms rsp] documentation +CME ERROR: 10")
        coordinator.observe_line("[I]-[lib_sms rsp] previous +CMGS AT+CMGS=12 false ERROR")
        self.assertFalse(waiter.done())

        coordinator.observe_line(">")
        self.assertTrue(waiter.wait_prompt(0).ok)
        self.assertTrue(waiter.mark_pdu_send_started())
        coordinator.observe_line("[I]-[app] OK")
        coordinator.observe_line("[I]-[ril.note] +CMS ERROR: 500")
        coordinator.observe_line("[I]-[app] +CMGS: 17")
        self.assertFalse(waiter.done())

        coordinator.observe_line("+CMGS: 17")
        coordinator.observe_line("[I]-[ril.proatc] OK")
        self.assertTrue(waiter.wait(0).ok)
        coordinator.finish(waiter)

    def test_sms_pdu_send_coordinator_accepts_exact_trusted_ril_error(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(label="(1/1)")

        coordinator.observe_line("[I]-[ril.proatc] +CMS ERROR: 500")
        response = waiter.wait(0)
        coordinator.finish(waiter)

        self.assertFalse(response.ok)
        self.assertEqual(response.error, "+CMS ERROR: 500")

    def test_sms_pdu_send_coordinator_ignores_other_cmgs_transaction_length(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(cmgs_len=12)

        coordinator.observe_line(">")
        self.assertTrue(waiter.wait_prompt(0).ok)
        self.assertTrue(waiter.mark_pdu_send_started())
        coordinator.observe_line("[I]-[lib_sms rsp] +CMGS AT+CMGS=99 true OK")
        self.assertFalse(waiter.done())
        coordinator.observe_line("[I]-[lib_sms rsp] +CMGS AT+CMGS=12 true OK")
        self.assertTrue(waiter.wait(0).ok)
        coordinator.finish(waiter)

    def test_sms_pdu_send_coordinator_binds_lines_to_current_connection(self):
        coordinator = SmsPduSendCoordinator()
        connection_a = object()
        connection_b = object()
        waiter = coordinator.begin_segment(connection=connection_a)

        coordinator.observe_line(">", connection=connection_b)
        self.assertFalse(waiter.wait_prompt(0).ok)
        coordinator.observe_line(">", connection=connection_a)
        self.assertTrue(waiter.wait_prompt(0).ok)
        self.assertTrue(waiter.mark_pdu_send_started())
        coordinator.observe_line("+CMGS: 17", connection=connection_b)
        coordinator.observe_line("OK", connection=connection_b)
        self.assertFalse(waiter.done())
        coordinator.observe_line("+CMGS: 17", connection=connection_a)
        coordinator.observe_line("OK", connection=connection_a)
        self.assertTrue(waiter.wait(0).ok)
        coordinator.finish(waiter)

    def test_sms_pdu_send_coordinator_reports_cms_error(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(label="(1/1)")

        coordinator.observe_line("+CMS ERROR: 500")
        response = waiter.wait(0)
        coordinator.finish(waiter)

        self.assertFalse(response.ok)
        self.assertEqual(response.error, "+CMS ERROR: 500")

    def test_sms_pdu_send_coordinator_reports_luat_false_error(self):
        coordinator = SmsPduSendCoordinator()
        waiter = coordinator.begin_segment(label="(1/1)")

        coordinator.observe_line("[I]-[lib_sms rsp] +CMGS AT+CMGS=12 false ERROR")
        response = waiter.wait(0)
        coordinator.finish(waiter)

        self.assertFalse(response.ok)
        self.assertIn("false ERROR", response.error)

    def test_write_serial_command_sequence_locked_releases_lock_while_waiting(self):
        serial_obj = FakeSerial()
        lock = TrackingLock()
        debug = []
        sleep_locked_states = []

        result = write_serial_command_sequence_locked(
            lock,
            lambda: serial_obj,
            ["AT", "ATI"],
            push_debug=debug.append,
            sleep_func=lambda _seconds: sleep_locked_states.append(lock.locked),
        )

        self.assertTrue(result)
        self.assertEqual(serial_obj.writes, [b"AT\r\n", b"ATI\r\n"])
        self.assertEqual(sleep_locked_states, [False])
        self.assertIn(">>> 发送: ATI\\r\\n", debug)

    def test_confirmed_command_sequence_waits_for_each_ok(self):
        coordinator = AtCommandResponseCoordinator()

        def on_write(_payload):
            coordinator.observe_line("[I]-[ril.proatc] OK")

        serial_obj = CallbackSerial(on_write)
        result = write_serial_command_sequence_confirmed_locked(
            threading.RLock(),
            lambda: serial_obj,
            ('AT+CPBS="ON"', 'AT+CPBW=1,"+8613123123123",145', "AT+CNUM"),
            response_coordinator=coordinator,
            response_timeout=0.1,
        )

        self.assertTrue(result.ok)
        self.assertEqual(
            serial_obj.writes,
            [
                b'AT+CPBS="ON"\r\n',
                b'AT+CPBW=1,"+8613123123123",145\r\n',
                b"AT+CNUM\r\n",
            ],
        )

    def test_confirmed_command_sequence_stops_after_first_error(self):
        coordinator = AtCommandResponseCoordinator()

        def on_write(_payload):
            coordinator.observe_line("[I]-[ril.proatc] +CME ERROR: 3")

        serial_obj = CallbackSerial(on_write)
        result = write_serial_command_sequence_confirmed_locked(
            threading.RLock(),
            lambda: serial_obj,
            ('AT+CPBS="ON"', 'AT+CPBW=1,"+8613123123123",145', "AT+CNUM"),
            response_coordinator=coordinator,
            response_timeout=0.1,
        )

        self.assertFalse(result.ok)
        self.assertEqual(result.error, "+CME ERROR: 3")
        self.assertEqual(serial_obj.writes, [b'AT+CPBS="ON"\r\n'])

    def test_confirmed_command_sequence_ignores_untrusted_ok(self):
        coordinator = AtCommandResponseCoordinator()

        def on_write(_payload):
            coordinator.observe_line("[I]-[app.note] OK")

        serial_obj = CallbackSerial(on_write)
        result = write_serial_command_sequence_confirmed_locked(
            threading.RLock(),
            lambda: serial_obj,
            ["AT"],
            response_coordinator=coordinator,
            response_timeout=0,
        )

        self.assertFalse(result.ok)
        self.assertIn("超时", result.error)
        self.assertEqual(serial_obj.writes, [b"AT\r\n"])

    def test_write_text_sms_pdu_locked_releases_lock_while_waiting(self):
        serial_obj = FakeSerial()
        lock = TrackingLock()
        debug = []
        ui_lines = []
        sleep_calls = []

        result = write_text_sms_pdu_locked(
            lock,
            lambda: serial_obj,
            "+1234",
            "A" * 100,
            push_debug=debug.append,
            port_ui=lambda *args: ui_lines.append(args),
            sleep_func=lambda seconds: sleep_calls.append((seconds, lock.locked)),
            response_coordinator=FakeSmsCoordinator([SmsPduSendResponse(True), SmsPduSendResponse(True)]),
        )

        self.assertTrue(result)
        self.assertEqual([locked for _seconds, locked in sleep_calls], [False, False])
        self.assertEqual([seconds for seconds, _locked in sleep_calls], [0.3, 1.5])
        self.assertEqual(ui_lines[0], ("📤 发送短信至 +1234：", "normal"))
        self.assertIn(">>> 发送: AT+CMGF=0\\r\\n", debug)

    def test_cloud_sms_transaction_confirms_cmgf_before_cmgs(self):
        command_coordinator = AtCommandResponseCoordinator()
        sms_coordinator = SmsPduSendCoordinator()

        def on_write(payload):
            if payload == b"AT+CMGF=0\r\n":
                command_coordinator.observe_line("[I]-[ril.proatc] OK")
            elif payload.startswith(b"AT+CMGS="):
                sms_coordinator.observe_line("[I]-[ril.proatc] >")
            elif payload.endswith(b"\x1a"):
                sms_coordinator.observe_line("+CMGS: 17")
                sms_coordinator.observe_line("[I]-[ril.proatc] OK")

        serial_obj = CallbackSerial(on_write)
        result = write_text_sms_pdu_locked(
            threading.RLock(),
            lambda: serial_obj,
            "+1234",
            "hello",
            response_coordinator=sms_coordinator,
            command_response_coordinator=command_coordinator,
            prompt_timeout=0.1,
            segment_timeout=0.1,
        )

        self.assertTrue(result)
        self.assertEqual(serial_obj.writes[0], b"AT+CMGF=0\r\n")
        self.assertTrue(serial_obj.writes[1].startswith(b"AT+CMGS="))
        self.assertTrue(serial_obj.writes[2].endswith(b"\x1a"))

    def test_cloud_sms_transaction_stops_when_cmgf_fails(self):
        command_coordinator = AtCommandResponseCoordinator()
        serial_obj = CallbackSerial(
            lambda _payload: command_coordinator.observe_line(
                "[I]-[ril.proatc] ERROR"
            )
        )

        result = write_text_sms_pdu_locked(
            threading.RLock(),
            lambda: serial_obj,
            "+1234",
            "hello",
            response_coordinator=SmsPduSendCoordinator(),
            command_response_coordinator=command_coordinator,
            segment_timeout=0.1,
        )

        self.assertFalse(result)
        self.assertEqual(serial_obj.writes, [b"AT+CMGF=0\r\n"])

    def test_write_text_sms_pdu_locked_fails_on_segment_error_and_stops(self):
        serial_obj = FakeSerial()
        lock = TrackingLock()
        debug = []
        ui_lines = []
        coordinator = FakeSmsCoordinator([
            SmsPduSendResponse(True),
            SmsPduSendResponse(False, "+CMS ERROR: 500"),
        ])

        result = write_text_sms_pdu_locked(
            lock,
            lambda: serial_obj,
            "+1234",
            "A" * 100,
            push_debug=debug.append,
            port_ui=lambda *args: ui_lines.append(args),
            sleep_func=lambda _seconds: None,
            response_coordinator=coordinator,
        )

        self.assertFalse(result)
        self.assertEqual(len(coordinator.waiters), 2)
        self.assertTrue(any("+CMS ERROR: 500" in line for line in debug))
        self.assertTrue(any("+CMS ERROR: 500" in item[0] for item in ui_lines if item))

    def test_write_text_sms_pdu_locked_does_not_send_without_prompt(self):
        serial_obj = FakeSerial()
        lock = TrackingLock()
        coordinator = SmsPduSendCoordinator()

        result = write_text_sms_pdu_locked(
            lock,
            lambda: serial_obj,
            "+1234",
            "hello",
            sleep_func=lambda _seconds: None,
            response_coordinator=coordinator,
            prompt_timeout=0,
        )

        self.assertFalse(result)
        self.assertFalse(any(payload.endswith(b"\x1a") for payload in serial_obj.writes))

    def test_async_sms_send_is_registered_and_cancelled_without_waiting_for_timeout(self):
        class WaitingSerial(FakeSerial):
            def __init__(self):
                super().__init__()
                self.cmgs_written = threading.Event()

            def write(self, payload):
                super().write(payload)
                if payload.startswith(b"AT+CMGS="):
                    self.cmgs_written.set()

        serial_obj = WaitingSerial()
        coordinator = SmsPduSendCoordinator()
        registry = WorkerThreadRegistry()
        ui_lines = []

        thread = send_text_sms_pdu_async(
            threading.Lock(),
            lambda: serial_obj,
            "+1234",
            "hello",
            port_ui=lambda *args: ui_lines.append(args),
            sleep_func=lambda _seconds: None,
            response_coordinator=coordinator,
            prompt_timeout=5,
            segment_timeout=5,
            thread_registry=registry,
        )

        self.assertTrue(serial_obj.cmgs_written.wait(1))
        self.assertIn(thread, registry.snapshot())
        self.assertTrue(coordinator.cancel_active("软件正在退出，短信发送已取消"))
        thread.join(1)

        self.assertFalse(thread.is_alive())
        self.assertEqual(registry.snapshot(), ())
        self.assertTrue(any("软件正在退出" in item[0] for item in ui_lines))

    def test_all_async_serial_command_types_register_until_write_finishes(self):
        class BlockingSerial(FakeSerial):
            def __init__(self):
                super().__init__()
                self.write_entered = threading.Event()
                self.release_write = threading.Event()

            def write(self, payload):
                self.write_entered.set()
                self.release_write.wait(2)
                super().write(payload)

        starters = (
            lambda serial_obj, registry: send_command_async(
                threading.Lock(),
                lambda: serial_obj,
                "AT",
                thread_registry=registry,
            ),
            lambda serial_obj, registry: send_command_with_result_async(
                threading.Lock(),
                lambda: serial_obj,
                "ATI",
                thread_registry=registry,
            ),
            lambda serial_obj, registry: send_command_sequence_async(
                threading.Lock(),
                lambda: serial_obj,
                ["AT", "ATI"],
                sleep_func=lambda _seconds: None,
                thread_registry=registry,
            ),
        )

        for start in starters:
            with self.subTest(start=start):
                serial_obj = BlockingSerial()
                registry = WorkerThreadRegistry()
                thread = start(serial_obj, registry)

                self.assertTrue(serial_obj.write_entered.wait(1))
                self.assertIn(thread, registry.snapshot())
                serial_obj.release_write.set()
                thread.join(2)

                self.assertFalse(thread.is_alive())
                self.assertEqual(registry.snapshot(), ())

    def test_serial_command_thread_unregisters_when_result_callback_raises(self):
        registry = WorkerThreadRegistry()
        logs = []

        thread = send_command_with_result_async(
            threading.Lock(),
            lambda: FakeSerial(),
            "ATA",
            on_result=lambda _result: (_ for _ in ()).throw(RuntimeError("callback failed")),
            log_error=logs.append,
            thread_registry=registry,
        )
        thread.join(2)

        self.assertFalse(thread.is_alive())
        self.assertEqual(registry.snapshot(), ())
        self.assertTrue(any("callback failed" in message for message in logs))

    def test_serial_command_thread_unregisters_when_thread_start_fails(self):
        registry = WorkerThreadRegistry()
        fake_thread = object()

        def fail_start(_name, _target, *, before_start, **_kwargs):
            before_start(fake_thread)
            raise RuntimeError("start failed")

        with patch("sms_core.serial_sender.start_daemon_thread", side_effect=fail_start):
            with self.assertRaisesRegex(RuntimeError, "start failed"):
                send_command_async(
                    threading.Lock(),
                    lambda: FakeSerial(),
                    "AT",
                    thread_registry=registry,
                )

        self.assertEqual(registry.snapshot(), ())


if __name__ == "__main__":
    unittest.main()
