import hashlib
from pathlib import Path
import threading
import unittest
from unittest.mock import patch

from sms_core.cloud_firmware_ota import RelayError
from sms_core.firmware_package import CORE, CAPACITY, MAX_PACKAGE_BYTES
from sms_core.local_firmware_update import LocalFirmwareUpdater, UpdateError, capture_local_context
from sms_core.threading_runtime import WorkerThreadRegistry
from tests.test_cloud_firmware_ota import namespace

FIXTURES = Path(__file__).parent / "fixtures/firmware_ota"


class Transport:
    def __init__(self, owner, context, current):
        self.owner, self.context, self.current = owner, context, current
        self.commit_sent = False
        self.closed = False
        self.aborted = False

    def exchange(self, request):
        if not self.current():
            raise RelayError("connection_changed")
        self.owner.requests.append(dict(request))
        action = request["action"]
        reply = dict(ok=True)
        if action == "info":
            reply["ota"] = dict(self.owner.capability)
        if action == "chunk":
            block = bytes.fromhex(request["data"])
            if hashlib.sha256(block).hexdigest() != request["sha256"]:
                raise AssertionError("Wrong block digest")
            reply["offset"] = request["offset"] + len(block)
        if action == "commit":
            self.commit_sent = True
        if self.owner.handler:
            return self.owner.handler(self, request, reply)
        return reply

    def wait_install_result(self, request):
        if self.owner.install_handler:
            return self.owner.install_handler(self, request)
        self.commit_sent = False
        return dict(ok=True, phase="rebooted")

    def close(self):
        self.closed = True

    def abort(self):
        self.aborted = not self.commit_sent


class LocalFirmwareUpdateTests(unittest.TestCase):
    def setUp(self):
        self.ns = namespace()
        self.ns.update(cloud_connected=False, cloud_device_authorized=False, CLOUD_CONTROL_ENABLED=False,
                       CLOUD_DEVICE_SECRET="", TK_SHUTDOWN=threading.Event(), is_exiting=False)
        self.requests, self.transports = [], []
        self.handler = self.install_handler = None
        self.capability = dict(protocol=1, supported=True, core=CORE, capacity=CAPACITY,
                               version="1.0.7", config_md5="a" * 32, main_md5="b" * 32, last_job="")
        self.updater = LocalFirmwareUpdater(self.ns, transport_factory=self.make_transport,
                                             start_worker=lambda work: work(), confirmation_timeout=0)

    def make_transport(self, context, current):
        transport = Transport(self, context, current)
        self.transports.append(transport)
        return transport

    def ready(self, suffix="dfota.bin"):
        self.assertTrue(self.updater.prepare(FIXTURES / ("synthetic." + suffix)))
        self.assertEqual(self.updater.snapshot()["phase"], "ready", self.updater.snapshot()["message"])

    def test_both_package_formats_work_without_cloud_or_device_secret(self):
        for suffix in ("dfota.bin", "air724ota"):
            self.ready(suffix)
            self.assertTrue(self.updater.start())
            self.assertEqual(self.updater.snapshot()["phase"], "success")
            self.assertEqual(self.updater.snapshot()["current_version"], "1.0.8")
        self.assertTrue(all("secret" not in request for request in self.requests))
        self.assertTrue(all(transport.closed for transport in self.transports))

    def test_selecting_file_only_probes_and_never_starts_installation(self):
        self.ready()
        self.assertTrue(all(request["action"] == "info" for request in self.requests))
        self.assertTrue(self.updater.snapshot()["config_changed"])

    def test_refresh_clears_previous_selection_and_start_permission(self):
        self.ready()
        self.updater.prepare()
        self.assertEqual(self.updater.snapshot()["phase"], "idle")
        self.assertFalse(self.updater.start())

    def test_bad_or_missing_file_clears_previous_ready_state(self):
        self.ready()
        self.updater.prepare(FIXTURES / "missing.bin")
        self.assertEqual(self.updater.snapshot()["phase"], "failed")
        self.assertFalse(self.updater.start())
        self.updater.prepare(__file__)
        self.assertEqual(self.updater.snapshot()["phase"], "failed")

    def test_oversized_read_is_bounded_and_rejected(self):
        from io import BytesIO
        stream = BytesIO(b"x" * (MAX_PACKAGE_BYTES + 20))
        with patch("sms_core.local_firmware_update.open", return_value=stream):
            self.updater.prepare("oversized.dfota.bin")
        self.assertEqual(self.updater.snapshot()["phase"], "failed")
        self.assertFalse(self.requests)

    def test_no_modem_or_unverified_identity_prevents_usb_probe(self):
        for key, value in (("serial_obj", None), ("imei", "")):
            with self.subTest(key=key):
                before = self.ns[key]
                self.ns[key] = value
                with self.assertRaises(UpdateError):
                    capture_local_context(self.ns)
                self.ns[key] = before
        self.assertFalse(self.transports)

    def test_unsupported_core_version_and_missing_hashes_are_rejected(self):
        for key, value in (("core", "wrong"), ("capacity", 999), ("supported", False),
                           ("main_md5", ""), ("version", "broken"), ("version", "1.0.8"), ("version", "2.0.0")):
            with self.subTest(key=key, value=value):
                before = self.capability[key]
                self.capability[key] = value
                self.updater.prepare(FIXTURES / "synthetic.dfota.bin")
                self.assertEqual(self.updater.snapshot()["phase"], "failed")
                self.assertFalse(self.updater.start())
                self.capability[key] = before
        self.assertTrue(all(request["action"] == "info" for request in self.requests))

    def test_connection_change_after_preflight_prevents_any_transfer(self):
        self.ready()
        self.requests.clear()
        self.ns["serial_connection_generation"] += 1
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "failed")
        self.assertFalse(self.requests)

    def test_device_version_is_checked_again_at_start(self):
        self.ready()
        self.capability["version"] = "1.0.8"
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "failed")
        self.assertTrue(all(request["action"] == "info" for request in self.requests))

    def test_payload_snapshot_cannot_change_after_selection(self):
        self.ready()
        with patch("sms_core.local_firmware_update.open", side_effect=AssertionError("Must not reread selected path")):
            self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "success")

    def test_timeout_retries_same_chunk_request_id_but_not_commit(self):
        dropped = set()
        def handler(transport, request, reply):
            action = request["action"]
            if action in {"chunk", "commit"} and action not in dropped:
                dropped.add(action)
                raise RelayError("usb_timeout")
            return reply
        self.handler = handler
        self.ready()
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "success")
        chunks = [r for r in self.requests if r["action"] == "chunk"]
        self.assertEqual(chunks[0], chunks[1])
        self.assertEqual(sum(r["action"] == "commit" for r in self.requests), 1)

    def test_wrong_offset_stops_before_install(self):
        self.handler = lambda transport, req, reply: dict(reply, offset=-1) if req["action"] == "chunk" else reply
        self.ready()
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "failed")
        self.assertFalse(any(r["action"] == "commit" for r in self.requests))

    def test_cancel_mid_transfer_aborts_without_commit(self):
        def handler(transport, request, reply):
            if request["action"] == "chunk":
                self.assertTrue(self.updater.cancel())
            return reply
        self.handler = handler
        self.ready()
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "cancelled")
        self.assertTrue(self.transports[-1].aborted)
        self.assertFalse(any(r["action"] == "commit" for r in self.requests))

    def test_cancel_is_disabled_once_commit_is_admitted(self):
        def handler(transport, request, reply):
            if request["action"] == "commit":
                self.assertFalse(self.updater.cancel())
            return reply
        self.handler = handler
        self.ready()
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "success")

    def test_cancel_during_preflight_never_leaves_a_startable_package(self):
        def cancel(transport, request, reply):
            self.assertTrue(self.updater.cancel())
            return reply
        self.handler = cancel
        self.updater.prepare(FIXTURES / "synthetic.dfota.bin")
        self.assertEqual(self.updater.snapshot()["phase"], "cancelled")
        self.assertFalse(self.updater.start())
        self.assertTrue(all(request["action"] == "info" for request in self.requests))

    def test_cancel_accepted_at_preflight_publication_cannot_be_overwritten(self):
        original_lock = self.updater.lock
        before_lock = []

        class PublicationLock:
            def __enter__(self):
                if before_lock:
                    before_lock.pop()()
                return original_lock.__enter__()

            def __exit__(self, *args):
                return original_lock.__exit__(*args)

        self.updater.lock = PublicationLock()
        read_device = self.updater._read_device

        def read_and_schedule_cancel(context):
            result = read_device(context)
            before_lock.append(lambda: self.assertTrue(self.updater.cancel()))
            return result

        with patch.object(self.updater, "_read_device", side_effect=read_and_schedule_cancel):
            for path in (None, FIXTURES / "synthetic.dfota.bin"):
                with self.subTest(path=path):
                    self.assertTrue(self.updater.prepare(path))
                    self.assertEqual(self.updater.snapshot()["phase"], "cancelled")
                    self.assertFalse(self.updater.start())
                    self.assertIsNone(self.updater.package)

    def test_timeouts_exhaust_bounded_retries_without_installing(self):
        self.ready()
        def timeout(transport, request, reply):
            if request["action"] == "chunk":
                raise RelayError("usb_timeout")
            return reply
        self.handler = timeout
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "failed")
        chunks = [request for request in self.requests if request["action"] == "chunk"]
        self.assertEqual(len(chunks), 3)
        self.assertEqual(chunks, [chunks[0]] * 3)
        self.assertFalse(any(request["action"] == "commit" for request in self.requests))

    def test_shutdown_during_confirmation_is_unconfirmed_and_closes_transport(self):
        self.ready()
        def stopping(transport, request):
            self.ns["TK_SHUTDOWN"].set()
            raise RelayError("connection_changed")
        self.install_handler = stopping
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "unconfirmed")
        self.assertFalse(self.updater.snapshot()["busy"])
        self.assertTrue(all(transport.closed for transport in self.transports))
        self.assertEqual(sum(request["action"] == "commit" for request in self.requests), 1)

    def test_native_failure_and_commit_rejection_are_not_reported_as_success(self):
        self.ready()
        self.install_handler = lambda *_: dict(ok=False, phase="failed", reason="integrity")
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "failed")
        self.ready()
        self.handler = lambda transport, req, reply: dict(ok=False, reason="busy") if req["action"] == "commit" else reply
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "failed")

    def test_missing_reboot_evidence_is_unconfirmed_and_cannot_restart_automatically(self):
        self.install_handler = lambda *_: (_ for _ in ()).throw(RelayError("usb_timeout"))
        self.ready()
        self.updater.start()
        self.assertEqual(self.updater.snapshot()["phase"], "unconfirmed")
        self.assertFalse(self.updater.start())
        self.assertEqual(sum(r["action"] == "commit" for r in self.requests), 1)

    def test_parallel_actions_are_rejected_and_thread_start_failure_unlocks(self):
        waiting = []
        self.updater.start_worker = waiting.append
        self.assertTrue(self.updater.prepare())
        self.assertFalse(self.updater.prepare())
        self.assertFalse(self.updater.start())
        waiting.pop()()
        self.assertFalse(self.updater.snapshot()["busy"])
        self.updater.start_worker = lambda _: (_ for _ in ()).throw(RuntimeError("thread unavailable"))
        self.assertFalse(self.updater.prepare())
        self.assertFalse(self.updater.snapshot()["busy"])

    def test_shutdown_stops_worker_and_prevents_new_tasks(self):
        waiting = []
        self.updater.start_worker = waiting.append
        self.updater.prepare()
        self.ns["TK_SHUTDOWN"].set()
        waiting.pop()()
        self.assertEqual(self.updater.snapshot()["phase"], "cancelled")
        self.assertFalse(self.updater.prepare())
        self.assertFalse(self.requests)

    def test_real_worker_is_registered_cancellable_and_joins_cleanly(self):
        entered, release = threading.Event(), threading.Event()
        registry = WorkerThreadRegistry()
        self.ns["UPDATE_THREAD_REGISTRY"] = registry
        self.updater.start_worker = self.updater._start_worker
        def wait_for_release(transport, request, reply):
            entered.set()
            release.wait(3)
            return reply
        self.handler = wait_for_release
        try:
            self.assertTrue(self.updater.prepare())
            self.assertTrue(entered.wait(2))
            threads = registry.snapshot()
            self.assertEqual(len(threads), 1)
            self.assertFalse(self.updater.prepare())
            self.assertTrue(self.updater.cancel())
        finally:
            release.set()
            for worker in registry.snapshot():
                worker.join(3)
        self.assertEqual(self.updater.snapshot()["phase"], "cancelled")
        self.assertFalse(self.updater.snapshot()["busy"])
        self.assertFalse(registry.snapshot())
        self.assertTrue(all(transport.closed for transport in self.transports))

    def test_canonical_parser_copy_has_not_diverged(self):
        client = Path(__file__).resolve().parents[1]
        canonical = client.parents[1] / "air724ug-sms-server/server_modules/server_firmware_package.py"
        if not canonical.exists():
            self.skipTest("Standalone desktop checkout has no server source")
        bundled = client / "sms/sms_core/firmware_package.py"
        self.assertEqual(bundled.read_text(encoding="utf-8-sig").split("\n", 1)[1], canonical.read_text(encoding="utf-8-sig"))


if __name__ == "__main__":
    unittest.main()
