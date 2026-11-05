import asyncio
import json
import queue
import unittest

import websockets
from websockets.exceptions import ConnectionClosedError
from websockets.frames import Close

from sms_core.cloud_message_namespace_runtime import drain_cloud_serial_log_queue_namespace_runtime
from sms_core.cloud_serial_log_runtime import (
    CloudSerialLogDrainState,
    clear_cloud_serial_log_queue,
    drain_cloud_serial_log_queue,
    reset_cloud_serial_log_state,
    schedule_cloud_serial_log_drain,
    send_cloud_serial_log_runtime,
)


def private_close_error():
    close = Close(1008, 'password=synthetic-close-secret path=synthetic-close-path')
    return ConnectionClosedError(close, close, True)


class OpaqueDiagnosticError(Exception):
    def __str__(self):
        raise AssertionError('Diagnostic must not format untrusted exception text')

    def __repr__(self):
        raise AssertionError('Diagnostic must not format untrusted exception text')


class CloudSerialLogDiagnosticSafetyTests(unittest.IsolatedAsyncioTestCase):
    async def test_live_close_reason_is_hidden_by_namespace_drain(self):
        private = 'synthetic-close-secret'
        ordinary_log = 'synthetic ordinary serial event'
        received_login = []
        logs = []

        async def peer(socket):
            received_login.append(json.loads(await socket.recv()))
            await socket.close(code=1008, reason='password=' + private + ' path=synthetic-close-path')

        async with websockets.serve(peer, '127.0.0.1', 0) as server:
            port = server.sockets[0].getsockname()[1]
            async with websockets.connect('ws://127.0.0.1:' + str(port)) as socket:
                await socket.send(json.dumps({'type': 'device_login', 'secret': private}))
                await asyncio.wait_for(socket.wait_closed(), 5)
                items = queue.Queue()
                items.put({'type': 'log', 'text': ordinary_log})
                state = CloudSerialLogDrainState(drain_scheduled=True)
                namespace = {'CLOUD_SERIAL_LOG_Q': items, 'CLOUD_SERIAL_LOG_DRAIN_BATCH': 25,
                             'CLOUD_SERIAL_LOG_DRAIN_STATE': state, 'cloud_ws_conn': socket,
                             'cloud_connected': True, 'cloud_device_authorized': True,
                             'log_file_only': logs.append}
                await drain_cloud_serial_log_queue_namespace_runtime(namespace, socket)
        self.assertEqual(received_login, [{'type': 'device_login', 'secret': private}])
        self.assertEqual(logs, ['Drain cloud serial log queue failed: ConnectionClosedError'])
        self.assertTrue(items.empty())
        self.assertEqual(items.unfinished_tasks, 0)
        self.assertFalse(state.drain_scheduled)
        self.assertNotIn(private, '\n'.join(logs))
        self.assertNotIn('synthetic-close-path', '\n'.join(logs))
        self.assertNotIn(ordinary_log, '\n'.join(logs))

    async def test_successful_namespace_batch_preserves_original_payloads_and_history(self):
        sent = []
        logs = []

        class Socket:
            async def send(self, payload):
                sent.append(json.loads(payload))

        socket = Socket()
        originals = [{'type': 'log', 'text': 'ordinary message ' + str(number)} for number in range(3)]
        items = queue.Queue()
        for payload in originals:
            items.put(payload)
        state = CloudSerialLogDrainState(drain_scheduled=True)
        namespace = {'CLOUD_SERIAL_LOG_Q': items, 'CLOUD_SERIAL_LOG_DRAIN_BATCH': 25,
                     'CLOUD_SERIAL_LOG_DRAIN_STATE': state, 'cloud_ws_conn': socket,
                     'cloud_connected': True, 'cloud_device_authorized': True, 'log_file_only': logs.append}
        await drain_cloud_serial_log_queue_namespace_runtime(namespace, socket)
        self.assertEqual(sent, [{**payload, 'history_only': True} for payload in originals])
        self.assertTrue(all('history_only' not in payload for payload in originals))
        self.assertEqual(items.unfinished_tasks, 0)
        self.assertFalse(state.drain_scheduled)
        self.assertEqual(logs, [])

    async def test_old_generation_transport_failure_preserves_new_queue_and_schedule(self):
        items = queue.Queue()
        items.put({'type': 'log', 'text': 'old'})
        state = CloudSerialLogDrainState(drain_scheduled=True)
        logs = []
        new_payload = {'type': 'log', 'text': 'new'}

        class Socket:
            async def send(self, _payload):
                reset_cloud_serial_log_state(items, state)
                items.put(new_payload)
                state.drain_scheduled = True
                raise private_close_error()

        socket = Socket()
        await drain_cloud_serial_log_queue(socket, log_queue=items, batch_size=25, state=state,
                                          is_current_connection=lambda current: current is socket,
                                          is_connected=lambda: True, log_error=logs.append)
        self.assertIs(items.get_nowait(), new_payload)
        items.task_done()
        self.assertEqual(items.unfinished_tasks, 0)
        self.assertEqual(state.generation, 1)
        self.assertTrue(state.drain_scheduled)
        self.assertEqual(logs, ['Drain cloud serial log queue failed: ConnectionClosedError'])

    async def test_batch_continuation_failure_closes_coroutine_and_releases_schedule(self):
        items = queue.Queue()
        for number in range(2):
            items.put({'number': number})
        sent, coroutines, logs = [], [], []
        state = CloudSerialLogDrainState(drain_scheduled=True)

        class Socket:
            async def send(self, payload):
                sent.append(json.loads(payload))

        def refuse(coroutine):
            coroutines.append(coroutine)
            raise private_close_error()

        socket = Socket()
        await drain_cloud_serial_log_queue(socket, log_queue=items, batch_size=1, state=state,
                                          is_current_connection=lambda _current: True,
                                          is_connected=lambda: True, create_task=refuse, log_error=logs.append)
        self.assertEqual(sent, [{'number': 0}])
        self.assertEqual(items.get_nowait(), {'number': 1})
        items.task_done()
        self.assertEqual(items.unfinished_tasks, 0)
        self.assertIsNone(coroutines[0].cr_frame)
        self.assertFalse(state.drain_scheduled)
        self.assertEqual(logs, ['Schedule next cloud serial log drain failed: ConnectionClosedError'])

    async def test_submission_failure_preserves_new_generation_while_closing_old_coroutine(self):
        state = CloudSerialLogDrainState()
        coroutines, logs = [], []

        async def work():
            return None

        def factory(_socket, _generation):
            coroutine = work()
            coroutines.append(coroutine)
            return coroutine

        def refuse(_coroutine, _loop):
            state.generation += 1
            state.drain_scheduled = True
            raise private_close_error()

        self.assertFalse(schedule_cloud_serial_log_drain(object(), object(), state=state,
                         drain_coro_factory=factory, run_coroutine_threadsafe=refuse, log_error=logs.append))
        self.assertIsNone(coroutines[0].cr_frame)
        self.assertEqual(state.generation, 1)
        self.assertTrue(state.drain_scheduled)
        self.assertEqual(logs, ['Schedule cloud serial log drain failed: ConnectionClosedError'])

    async def test_queue_and_reset_failures_never_format_exception_text(self):
        logs = []

        class BrokenQueue:
            def get_nowait(self):
                raise OpaqueDiagnosticError()

        class BrokenLock:
            def __enter__(self):
                raise OpaqueDiagnosticError()

            def __exit__(self, *_args):
                return False

        clear_cloud_serial_log_queue(BrokenQueue(), log_error=logs.append)
        reset_cloud_serial_log_state(queue.Queue(), CloudSerialLogDrainState(lock=BrokenLock()), log_error=logs.append)
        self.assertEqual(logs, ['Clear cloud serial log queue failed: OpaqueDiagnosticError',
                                'Reset cloud serial log state failed: OpaqueDiagnosticError'])

    async def test_producer_failure_never_formats_exception_or_queues_payload(self):
        logs = []
        items = queue.Queue()

        class Loop:
            def is_running(self):
                return True

        loop, socket = Loop(), object()

        def build(_line):
            raise OpaqueDiagnosticError()

        self.assertEqual(send_cloud_serial_log_runtime('ordinary event', authorized=True,
                         get_loop=lambda: loop, get_ws=lambda: socket, is_connected=lambda: True,
                         runtime_imei=lambda: '000000000000001', build_payload=build, log_queue=items,
                         schedule_drain=lambda *_args: None, log_error=logs.append), 'error')
        self.assertTrue(items.empty())
        self.assertEqual(items.unfinished_tasks, 0)
        self.assertEqual(logs, ['Send cloud serial log failed: OpaqueDiagnosticError'])


if __name__ == '__main__':
    unittest.main()
