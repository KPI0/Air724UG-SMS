import asyncio
from contextlib import suppress
import socket
import unittest
from unittest.mock import patch

import websockets

from sms_core.cloud_ws_runtime import cloud_websocket_connect_kwargs


class CloudSocksProxyTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.received, self.proxy_requests, self.proxy_errors = [], [], []
        self.proxy_tasks = set()
        self.echo_server = await websockets.serve(self._echo, "127.0.0.1", 0, close_timeout=1)
        self.addAsyncCleanup(self._close_server, self.echo_server)
        self.echo_port = self.echo_server.sockets[0].getsockname()[1]
        self.proxy_server = await asyncio.start_server(self._proxy, "127.0.0.1", 0)
        self.addAsyncCleanup(self._close_server, self.proxy_server)
        self.proxy_port = self.proxy_server.sockets[0].getsockname()[1]
        self.bypass_patch = patch("websockets.proxy.urllib.request.proxy_bypass", return_value=False)
        self.proxies_patch = patch("websockets.proxy.urllib.request.getproxies", return_value={})
        self.bypass_patch.start()
        self.addCleanup(self.bypass_patch.stop)
        self.getproxies = self.proxies_patch.start()
        self.addCleanup(self.proxies_patch.stop)

    async def _close_server(self, server):
        server.close()
        await server.wait_closed()

    async def asyncTearDown(self):
        tasks = list(self.proxy_tasks)
        for task in tasks:
            task.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)

    async def _echo(self, websocket):
        async for message in websocket:
            self.received.append(message)
            await websocket.send(message)

    async def _pipe(self, reader, writer):
        while data := await reader.read(4096):
            writer.write(data)
            await writer.drain()
        writer.close()

    async def _proxy(self, reader, writer):
        task = asyncio.current_task()
        self.proxy_tasks.add(task)
        peer_writer = None
        try:
            version, method_count = await asyncio.wait_for(reader.readexactly(2), 2)
            self.assertEqual(version, 5)
            methods = await asyncio.wait_for(reader.readexactly(method_count), 2)
            self.assertIn(0, methods)
            writer.write(b"\x05\x00")
            await writer.drain()
            version, command, reserved, address_type = await asyncio.wait_for(reader.readexactly(4), 2)
            self.assertEqual((version, command, reserved), (5, 1, 0))
            self.assertEqual(address_type, 3)
            size = (await reader.readexactly(1))[0]
            host = (await reader.readexactly(size)).decode("ascii")
            port = int.from_bytes(await reader.readexactly(2), "big")
            self.assertEqual((host, port), ("device.invalid", self.echo_port))
            self.proxy_requests.append((host, port))
            peer_reader, peer_writer = await asyncio.open_connection("127.0.0.1", self.echo_port)
            writer.write(b"\x05\x00\x00\x01" + socket.inet_aton("127.0.0.1") + self.echo_port.to_bytes(2, "big"))
            await writer.drain()
            await asyncio.gather(self._pipe(reader, peer_writer), self._pipe(peer_reader, writer))
        except (AssertionError, asyncio.IncompleteReadError, OSError) as exc:
            self.proxy_errors.append(type(exc).__name__)
        finally:
            for connection in (writer, peer_writer):
                if connection is not None:
                    connection.close()
                    with suppress(OSError):
                        await connection.wait_closed()
            self.proxy_tasks.discard(task)

    async def _roundtrip(self, url):
        kwargs = cloud_websocket_connect_kwargs(url)
        self.assertNotIn("proxy", kwargs)
        async with websockets.connect(url, **kwargs, open_timeout=2, close_timeout=1) as websocket:
            await websocket.send("synthetic-device-message")
            self.assertEqual(await asyncio.wait_for(websocket.recv(), 2), "synthetic-device-message")

    async def test_default_proxy_selection_completes_socks_handshake_and_websocket_echo(self):
        self.getproxies.return_value = {"socks": f"socks5h://127.0.0.1:{self.proxy_port}"}
        await self._roundtrip(f"ws://device.invalid:{self.echo_port}/ws/device")
        self.assertEqual(self.proxy_requests, [("device.invalid", self.echo_port)])
        self.assertEqual(self.received, ["synthetic-device-message"])
        self.assertEqual(self.proxy_errors, [])
        self.getproxies.assert_called()

    async def test_direct_cloud_connection_keeps_websocket_echo_without_proxy(self):
        await self._roundtrip(f"ws://127.0.0.1:{self.echo_port}/ws/device")
        self.assertEqual(self.proxy_requests, [])
        self.assertEqual(self.received, ["synthetic-device-message"])
        self.assertEqual(self.proxy_errors, [])
        self.getproxies.assert_called()


class CloudSocksProxySetupCleanupTests(unittest.TestCase):
    def test_proxy_start_failure_closes_created_echo_listener(self):
        ports = []
        original_serve = websockets.serve

        async def create_echo(*args, **kwargs):
            server = await original_serve(*args, **kwargs)
            ports.append(server.sockets[0].getsockname()[1])
            return server

        case = CloudSocksProxyTests("test_direct_cloud_connection_keeps_websocket_echo_without_proxy")
        result = unittest.TestResult()
        with patch("websockets.serve", side_effect=create_echo), \
                patch("asyncio.start_server", side_effect=OSError("synthetic proxy setup failure")):
            case.run(result)
        self.assertEqual(len(result.errors), 1)
        self.assertIn("synthetic proxy setup failure", result.errors[0][1])
        self.assertEqual(result.failures, [])
        self.assertEqual(len(ports), 1)
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
            listener.bind(("127.0.0.1", ports[0]))


if __name__ == "__main__":
    unittest.main()
