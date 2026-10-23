import gc
import os
import tempfile
import threading
import tkinter as tk
import unittest
import asyncio
from dataclasses import replace
from types import SimpleNamespace
from unittest.mock import patch

from sms_core.cloud_command_context import (
    capture_cloud_serial_command_context, set_cloud_command_authorized,
)
from sms_core.cloud_runtime import CloudControlSettings, read_cloud_control_settings
from sms_core.cloud_ws_runtime import cloud_ws_main_runtime
from sms_core.cloud_sms_event_runtime import CloudSmsEventDrainState
from sms_core.config_runtime import (
    remember_config_snapshot, safe_save_config_runtime, snapshot_config_runtime,
)
from sms_ui.cloud_control_namespace_runtime import (
    apply_cloud_control_settings_namespace_runtime,
    cloud_control_settings_snapshot_namespace_runtime,
    restart_cloud_control_namespace_runtime,
    save_cloud_control_setting_namespace_runtime,
    start_cloud_control_namespace_runtime,
    stop_cloud_control_namespace_runtime,
    sync_cloud_control_settings_namespace_runtime,
)
from sms_ui.cloud_control_window import open_cloud_control_window_dialog, open_cloud_control_window_runtime
from sms_ui.config_sync_namespace_bindings import install_config_sync_namespace_bindings
from sms_ui.config_sync_namespace_runtime import reload_shared_ui_config_namespace_runtime
from tests import test_config_sync_namespace_runtime as sync_tests
from tests.test_config_sync_runtime import FakeRootTimer


INITIAL = CloudControlSettings(True, "wss://old.example.invalid/ws/device", 5, "synthetic-old", False)


def make_namespace(settings=INITIAL):
    namespace = sync_tests.ConfigSyncNamespaceRuntimeTests().make_namespace()
    namespace.update(
        cloud_ws_lock=threading.Lock(), cloud_restart_seq=0,
        cloud_stop_event=threading.Event(), cloud_ws_thread=None,
        cloud_connected=True, cloud_device_authorized=True,
        cloud_ws_loop=object(), cloud_ws_conn=object(),
        CLOUD_SMS_EVENT_DRAIN_STATE=CloudSmsEventDrainState(),
    )
    namespace["apply_cloud_control_settings"] = lambda value: apply_cloud_control_settings_namespace_runtime(namespace, value)
    namespace["apply_cloud_control_settings"](settings)
    namespace["cloud_stop_event"].clear()
    namespace["config"]["cloud_control"] = {
        "enabled": "1" if settings.enabled else "0", "url": settings.url,
        "reconnect_interval": str(settings.reconnect_interval),
        "device_secret": settings.device_secret, "auto_upload": "1" if settings.auto_upload else "0",
    }
    remember_config_snapshot(namespace["config"])
    namespace["stop_cloud_control"] = lambda **kwargs: namespace["calls"].append(("stop", kwargs))
    namespace["restart_cloud_control"] = lambda **kwargs: namespace["calls"].append(("restart", kwargs))
    namespace["_cloud_send_register"] = lambda ws: ("register", ws)
    namespace["_cloud_log"] = lambda *_args, **_kwargs: None
    namespace["_cloud_unregister_then_close"] = lambda _ws: None
    namespace["asyncio"] = SimpleNamespace(
        run_coroutine_threadsafe=lambda coro, loop: namespace["calls"].append(("register", coro, loop)),
    )
    namespace["_cloud_schedule_unregister"] = lambda reason: namespace["calls"].append(("unregister", reason))
    namespace["safe_save_config"] = lambda: True
    return namespace


class CloudConfigSyncTests(unittest.TestCase):
    def test_all_five_fields_apply_and_select_connection_action(self):
        cases = (
            ({"enabled": False}, "云端控制开关", "stop"),
            ({"url": "wss://next.example.invalid/ws/device"}, "云端服务器地址", "restart"),
            ({"reconnect_interval": 13}, "云端重连间隔", "restart"),
            ({"device_secret": "synthetic-next"}, "云端设备密钥", "restart"),
            ({"auto_upload": True}, "主动公开设备", "register"),
        )
        for update, label, action in cases:
            with self.subTest(field=label):
                namespace = make_namespace()
                target = replace(INITIAL, **update)
                changes = sync_cloud_control_settings_namespace_runtime(namespace, target)
                self.assertEqual(changes, (label,))
                self.assertEqual(cloud_control_settings_snapshot_namespace_runtime(namespace), target)
                self.assertEqual(namespace["calls"][0][0], action)
                self.assertEqual(sync_cloud_control_settings_namespace_runtime(namespace, target), ())
                self.assertEqual(len(namespace["calls"]), 1)

    def test_publication_only_keeps_authorization_and_unpublishes_without_restart(self):
        namespace = make_namespace(replace(INITIAL, auto_upload=True))
        set_cloud_command_authorized(namespace, True)
        epoch = namespace["_cloud_authorization_epoch"]
        sync_cloud_control_settings_namespace_runtime(namespace, INITIAL)
        self.assertEqual(namespace["calls"], [("unregister", "auto_upload_disabled")])
        self.assertTrue(namespace["cloud_device_authorized"])
        self.assertIs(namespace["_cloud_authorization_epoch"], epoch)
        self.assertFalse(namespace["cloud_stop_event"].is_set())

    def test_connection_change_stops_old_handshake_before_publishing_new_values(self):
        namespace = make_namespace()
        namespace["apply_cloud_control_settings"](replace(INITIAL, url="wss://next.example.invalid/ws/device"))
        self.assertTrue(namespace["cloud_stop_event"].is_set())

    def test_handshake_finishing_after_sync_stop_cannot_register(self):
        stop_event = threading.Event()
        calls = []

        class LateConnection:
            async def __aenter__(self):
                stop_event.set()
                return object()

            async def __aexit__(self, *_args):
                calls.append("closed")

        async def register(_ws):
            self.fail("a cancelled connection must not register")

        with patch("sms_core.cloud_ws_runtime.cloud_websocket_connect_kwargs", return_value={}):
            asyncio.run(cloud_ws_main_runtime(
                "ws://127.0.0.1/ws/device", 5, stop_event=stop_event, runtime_imei=lambda: "synthetic",
                request_cloud_device_imei=lambda: None, set_cloud_status=lambda *_: None,
                log=lambda *_args, **_kwargs: None, connect=lambda *_args, **_kwargs: LateConnection(),
                set_connection_state=lambda ws, **_kwargs: calls.append(ws),
                reset_serial_log_state=lambda: None, send_register=register, wait_login_ack=register,
                handle_message=lambda *_: None, cloud_control_enabled=False, monotonic=lambda: 1,
            ))
        self.assertEqual(calls, ["closed", None])

    def test_disabled_cloud_accepts_other_fields_without_starting(self):
        namespace = make_namespace(replace(INITIAL, enabled=False))
        target = CloudControlSettings(False, "", 1, "", True)
        self.assertEqual(len(sync_cloud_control_settings_namespace_runtime(namespace, target)), 4)
        self.assertEqual(namespace["calls"], [])
        self.assertFalse(start_cloud_control_namespace_runtime(namespace))

    def test_reconnect_revokes_original_command_even_after_reauthorization(self):
        for field, value in (("enabled", False), ("url", "wss://next.example.invalid/ws/device"),
                             ("device_secret", "synthetic-next")):
            with self.subTest(field=field):
                namespace = make_namespace()
                namespace.update(serial_lock=threading.RLock(), serial_obj=object(),
                                 serial_connection_generation=1, _cloud_runtime_imei=lambda: "synthetic-imei")
                set_cloud_command_authorized(namespace, True)
                command = capture_cloud_serial_command_context(namespace)
                self.assertEqual(command.rejection_reason(namespace), "")
                namespace["apply_cloud_control_settings"](replace(INITIAL, **{field: value}))
                self.assertFalse(namespace["cloud_device_authorized"])
                namespace["apply_cloud_control_settings"](INITIAL)
                set_cloud_command_authorized(namespace, True)
                self.assertTrue(command.rejection_reason(namespace))

    def test_settings_update_takes_event_lock_before_command_lock(self):
        namespace = make_namespace()
        acquired = threading.Event()
        finished = threading.Event()
        event_lock = namespace["CLOUD_SMS_EVENT_DRAIN_STATE"].lock
        peer_lock = namespace["_call_peer_state_lock"]
        errors = []

        def change_settings():
            acquired.set()
            try:
                namespace["apply_cloud_control_settings"](replace(INITIAL, device_secret="synthetic-next"))
            except Exception as exc:
                errors.append(type(exc).__name__)
            finally:
                finished.set()

        with event_lock:
            worker = threading.Thread(target=change_settings, daemon=True)
            worker.start()
            self.assertTrue(acquired.wait(1))
            self.assertTrue(peer_lock.acquire(timeout=1), "settings update must not hold peer lock while waiting for event lock")
            peer_lock.release()
        self.assertTrue(finished.wait(2))
        worker.join(1)
        self.assertEqual(errors, [])

    def test_pending_restart_is_cancelled_by_disable(self):
        namespace = make_namespace()
        queued = []
        namespace.update(tk_alive=lambda: True, ui_post=queued.append)
        namespace["_reset_cloud_serial_log_state"] = lambda: None
        namespace["set_cloud_status"] = lambda *_: None

        class ImmediateThread:
            def __init__(self, target, **_kwargs):
                self.target = target

            def start(self):
                self.target()

        namespace["threading"] = SimpleNamespace(Thread=ImmediateThread)
        namespace["stop_cloud_control"] = lambda **kwargs: stop_cloud_control_namespace_runtime(namespace, **kwargs)
        namespace["start_cloud_control"] = lambda **kwargs: namespace["calls"].append(("start", kwargs))
        restart_cloud_control_namespace_runtime(namespace)
        self.assertEqual(len(queued), 1)
        namespace["apply_cloud_control_settings"](replace(INITIAL, enabled=False))
        namespace["stop_cloud_control"]()
        queued.pop()()
        self.assertEqual(namespace["calls"], [])
        self.assertTrue(namespace["cloud_stop_event"].is_set())

    def test_sync_errors_and_window_errors_do_not_expose_configuration(self):
        namespace = make_namespace()
        namespace["restart_cloud_control"] = lambda **_: (_ for _ in ()).throw(RuntimeError("synthetic-sensitive-value"))
        namespace["cloud_control_win"] = SimpleNamespace(
            winfo_exists=lambda: True,
            _refresh_synced_fields=lambda: (_ for _ in ()).throw(ValueError("synthetic-sensitive-value")),
        )
        snapshot = snapshot_config_runtime(namespace["config"])
        snapshot["cloud_control"]["device_secret"] = "synthetic-sensitive-value"
        changes = reload_shared_ui_config_namespace_runtime(namespace, load_snapshot=lambda _: snapshot)
        self.assertIn("云端设备密钥", changes)
        self.assertTrue(all("synthetic-sensitive-value" not in str(call) for call in namespace["calls"]))

    def test_shared_file_watch_and_partial_saves_preserve_other_instances(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "config.ini")
            first, second, third = [make_namespace() for _ in range(3)]
            with open(path, "w", encoding="utf-8") as file_obj:
                first["config"].write(file_obj)
            timers = []
            for namespace in (first, second, third):
                namespace["CONFIG_FILE"] = path
                namespace["safe_save_config"] = lambda ns=namespace: safe_save_config_runtime(
                    config=ns["config"], config_file=path, config_lock=ns["CONFIG_LOCK"],
                )
                timer = FakeRootTimer()
                namespace["root"] = SimpleNamespace(after=timer.after, after_cancel=timer.after_cancel,
                                                    winfo_rgb=sync_tests.FakeRoot().winfo_rgb)
                timers.append(timer)
                install_config_sync_namespace_bindings(namespace)
                namespace["start_config_file_watch"]()
            save_cloud_control_setting_namespace_runtime(first, enabled=False, auto_upload=True,
                                                       url="wss://next.example.invalid", device_secret="synthetic-next",
                                                       reconnect_interval=17)
            timers[1].run(second["CONFIG_FILE_WATCH_STATE"].after_id)
            expected = read_cloud_control_settings(first["config"])
            self.assertEqual(cloud_control_settings_snapshot_namespace_runtime(second), expected)
            self.assertTrue(any(call[0] == "stop" for call in second["calls"]))
            self.assertEqual((second["PORT"], second["BAUD"], second["MODE"]), ("COM1", 115200, "Manual"))
            # Saving a different category merges disk changes before cloud runtime refresh.
            third["config"].set("ui", "popup_enabled", "0")
            self.assertTrue(third["safe_save_config"]())
            self.assertEqual(cloud_control_settings_snapshot_namespace_runtime(third), INITIAL)
            saved = save_cloud_control_setting_namespace_runtime(third, auto_upload=False)
            self.assertEqual(saved, replace(expected, auto_upload=False))
            self.assertEqual(cloud_control_settings_snapshot_namespace_runtime(third), saved)
            timers[0].run(first["CONFIG_FILE_WATCH_STATE"].after_id)
            self.assertEqual(cloud_control_settings_snapshot_namespace_runtime(first), saved)
            save_cloud_control_setting_namespace_runtime(first, enabled=True)
            timers[1].run(second["CONFIG_FILE_WATCH_STATE"].after_id)
            self.assertTrue(any(call[0] == "restart" for call in second["calls"]))
            self.assertTrue(second["CLOUD_CONTROL_ENABLED"])

    def test_local_publication_save_restarts_for_merged_credentials(self):
        callbacks = {}
        actions = []
        state = {"enabled": True, "auto_upload": False, "url": INITIAL.url,
                 "secret": INITIAL.device_secret, "reconnect_interval": 5}

        def dialog(_parent, _provider, _status, on_public, *_args, **_kwargs):
            callbacks["public"] = on_public
            return "window"

        with patch("sms_ui.cloud_control_window.open_cloud_control_window_dialog", side_effect=dialog):
            open_cloud_control_window_runtime(
                None, None, lambda: state, None, lambda: None,
                lambda **_: replace(INITIAL, device_secret="synthetic-merged", auto_upload=True),
                lambda: (True, True, True), lambda: actions.append("register"),
                lambda _: actions.append("unregister"), lambda **_: actions.append("restart"),
                lambda: actions.append("stop"), lambda *_: None, lambda *_: False,
                lambda _: None, lambda *_: None,
            )
        self.assertTrue(callbacks["public"](True, None))
        self.assertEqual(actions, ["restart"])

    def test_full_form_save_uses_only_changes_when_config_already_merged(self):
        namespace = make_namespace()
        namespace["config"]["cloud_control"].update(
            url="wss://merged.example.invalid/ws/device", device_secret="synthetic-merged",
            reconnect_interval="19", auto_upload="1",
        )
        callbacks = {}
        applied = []

        def dialog(_parent, _provider, _status, _public, on_save, *_args, **_kwargs):
            callbacks["save"] = on_save
            return "window"

        def save(**kwargs):
            applied.append(kwargs)
            return save_cloud_control_setting_namespace_runtime(namespace, **kwargs)

        with patch("sms_ui.cloud_control_window.open_cloud_control_window_dialog", side_effect=dialog):
            open_cloud_control_window_runtime(
                None, None, lambda: {"enabled": INITIAL.enabled, "auto_upload": INITIAL.auto_upload,
                                    "url": INITIAL.url, "secret": INITIAL.device_secret, "reconnect_interval": 5},
                None, lambda: None, save, lambda: (True, True, True), lambda: None, lambda _: None,
                namespace["restart_cloud_control"], namespace["stop_cloud_control"], lambda *_: None,
                lambda *_: False, lambda _: None, lambda *_: None,
            )
        with patch("sms_ui.cloud_control_window.messagebox.showinfo"):
            self.assertTrue(callbacks["save"]((True, INITIAL.url, 7, INITIAL.device_secret, False), None))
        self.assertEqual(applied, [{"reconnect_interval": 7}])
        saved = cloud_control_settings_snapshot_namespace_runtime(namespace)
        self.assertEqual(saved.url, "wss://merged.example.invalid/ws/device")
        self.assertEqual(saved.device_secret, "synthetic-merged")
        self.assertTrue(saved.auto_upload)
        self.assertEqual(saved.reconnect_interval, 7)


class CloudSyncFormTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(str(exc))
        self.root.withdraw()
        self.state = {"enabled": False, "auto_upload": False, "url": "", "secret": "", "reconnect_interval": 5}
        self.callbacks = []
        self.fail_save = False
        self.window = open_cloud_control_window_dialog(
            self.root, lambda: dict(self.state), tk.StringVar(self.root),
            lambda value, _: self.save(auto_upload=value),
            lambda values, _: self.save_values(values), lambda values, _: self.save_values(values),
            lambda _values, _: self.save(enabled=False), lambda win: win.destroy(), lambda *_: None,
            on_enabled_changed=lambda value, values, _: self.save_values(values) if value else self.save(enabled=False),
        )
        self.window.withdraw()

    def tearDown(self):
        self.root.destroy()
        self.window = self.root = None
        gc.collect()

    def save(self, **kwargs):
        if self.fail_save:
            return False
        self.state.update(kwargs)
        self.callbacks.append(kwargs)
        return True

    def save_values(self, values):
        enabled, url, interval, secret, public = values
        return self.save(enabled=enabled, url=url, reconnect_interval=interval, secret=secret, auto_upload=public)

    def buttons(self):
        def walk(widget):
            yield widget
            for child in widget.winfo_children():
                yield from walk(child)
        return {str(widget.cget("text")): widget for widget in walk(self.window)
                if widget.winfo_class() in ("TButton", "TCheckbutton")}

    def test_external_refresh_preserves_drafts_and_does_not_invoke_public_confirmation(self):
        self.window._url_var.set("wss://draft.example.invalid/ws/device")
        self.state.update(enabled=True, auto_upload=True, url="wss://external.example.invalid/ws/device",
                          secret="synthetic-next", reconnect_interval=13)
        with patch("sms_ui.cloud_control_window_form.confirm_public_device") as confirm:
            self.window._refresh_synced_fields()
            self.window._sync_form_from_state()
        confirm.assert_not_called()
        self.assertTrue(self.window._enabled_var.get())
        self.assertTrue(self.window._auto_upload_var.get())
        self.assertEqual(self.window._url_var.get(), "wss://draft.example.invalid/ws/device")
        self.assertEqual(self.window._reconnect_var.get(), "13")
        self.assertEqual(self.window._secret_var.get(), "synthetic-next")
        self.assertEqual(self.callbacks, [])

    def test_partial_toggle_preserves_drafts_and_failed_disable_restores_checkbox(self):
        self.window._reconnect_var.set("invalid-draft")
        with patch("sms_ui.cloud_control_window_form.confirm_public_device", return_value=True):
            self.buttons()["主动公开设备"].invoke()
        self.assertTrue(self.state["auto_upload"])
        self.assertEqual(self.window._reconnect_var.get(), "invalid-draft")
        self.state.update(enabled=True)
        self.window._refresh_synced_fields()
        self.fail_save = True
        self.buttons()["启用云端控制"].invoke()
        self.assertTrue(self.window._enabled_var.get())
        self.assertEqual(self.window._reconnect_var.get(), "invalid-draft")
        self.fail_save = False
        self.buttons()["断开"].invoke()
        self.assertFalse(self.state["enabled"])
        self.assertEqual(self.window._reconnect_var.get(), "invalid-draft")

    def test_successful_full_save_rebases_draft_for_later_external_refresh(self):
        self.window._url_var.set("wss://draft.example.invalid/ws/device")
        self.window._reconnect_var.set("11")
        self.buttons()["保存"].invoke()
        self.assertEqual(self.state["reconnect_interval"], 11)
        self.state.update(url="wss://external.example.invalid/ws/device", reconnect_interval=19)
        self.window._refresh_synced_fields()
        self.assertEqual(self.window._url_var.get(), self.state["url"])
        self.assertEqual(self.window._reconnect_var.get(), "19")

    def test_save_rereads_fields_synced_while_confirmation_is_open(self):
        for scheme in ("wss", "ws"):
            with self.subTest(scheme=scheme):
                self.state.update(enabled=True, url="ws://old.example.invalid/ws/device",
                                  secret="synthetic-old", reconnect_interval=5)
                self.window._refresh_synced_fields()
                next_url = f"{scheme}://next.example.invalid/ws/device"
                confirmations = []

                def confirm(_win):
                    confirmations.append(True)
                    self.state.update(url=next_url, secret="synthetic-next", reconnect_interval=13)
                    self.window._refresh_synced_fields()
                    return True

                with patch("sms_ui.cloud_control_window_form.confirm_insecure_ws", side_effect=confirm):
                    self.buttons()["保存"].invoke()
                self.assertEqual(self.state["url"], next_url)
                self.assertEqual(self.state["secret"], "synthetic-next")
                self.assertEqual(self.state["reconnect_interval"], 13)
                self.assertEqual(len(confirmations), 1 if scheme == "wss" else 2)

    def test_confirmation_refresh_keeps_draft_and_revalidates_changed_fields(self):
        self.state.update(enabled=True, url="ws://old.example.invalid/ws/device",
                          secret="synthetic-old", reconnect_interval=5)
        self.window._refresh_synced_fields()
        self.window._reconnect_var.set("7")

        def confirm(_win):
            self.state.update(url="wss://next.example.invalid/ws/device", secret="", reconnect_interval=13)
            self.window._refresh_synced_fields()
            return True

        with patch("sms_ui.cloud_control_window_form.confirm_insecure_ws", side_effect=confirm), \
                patch("sms_ui.cloud_control_window_form.messagebox.showerror") as show_error:
            self.buttons()["保存"].invoke()
        self.assertEqual(self.callbacks, [])
        self.assertEqual(self.window._reconnect_var.get(), "7")
        show_error.assert_called_once()


if __name__ == "__main__":
    unittest.main()
