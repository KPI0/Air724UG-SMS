import asyncio
import configparser
import gc
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import io
import json
from pathlib import Path
import queue
import ssl
import tempfile
import threading
import time
import tkinter as tk
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import wave
import weakref

from sms_core.call_recordings import CallRecordingRepository, CloudCallRecordingUploader
from sms_core.config_runtime import config_mutex_name, safe_save_config_runtime
from sms_core.http_deadline import _Deadline, _HTTPSConnection
from sms_core.serial_namespace_runtime import try_rebind_manual_port_namespace_runtime
from sms_core.third_push_sender import http_request
from sms_core.tts_runtime import enqueue_tts_request, generate_tts_file, tts_worker_loop
from sms_core.windows_runtime import acquire_named_mutex_lock, release_named_mutex_lock
from sms_ui.app_infrastructure_namespace_runtime import safe_save_config_namespace_runtime
from sms_ui.app_lifecycle_namespace_runtime import cleanup_and_exit_namespace_runtime
from sms_ui.audio_namespace_runtime import generate_alert_voice_namespace_runtime, play_alert_namespace_runtime
from sms_ui.audio_runtime import play_alert_runtime
from sms_ui.config_sync_namespace_runtime import reload_shared_ui_config_namespace_runtime
from sms_ui.config_save_runtime import UiConfigSave
from sms_ui.settings_runtime import open_voice_text_dialog_runtime, toggle_voice_broadcast_runtime
from sms_ui.shutdown_progress import ShutdownProgress
from sms_ui.utility_dialogs import open_voice_text_dialog


def wave_data(text="saved"):
    output = io.BytesIO()
    with wave.open(output, "wb") as audio:
        audio.setnchannels(1)
        audio.setsampwidth(2)
        audio.setframerate(8000)
        audio.writeframes(text.encode("utf-16-le"))
    return output.getvalue()


class WaveEngine:
    def setProperty(self, *_args):
        pass

    def save_to_file(self, text, path):
        Path(path).write_bytes(wave_data(text))

    def runAndWait(self):
        pass

    def stop(self):
        pass


def drain_tts(tasks, path, *, engine=WaveEngine, played=None, errors=None, beeps=None, preview=None):
    class StopWhenDone:
        def is_set(self):
            return tasks.unfinished_tasks == 0

    state = [str(path)]
    tts_worker_loop(
        StopWhenDone(), tasks, threading.Lock(), lambda: state[0], lambda value: state.__setitem__(0, value),
        str(path.parent), "default", lambda **_kw: played.append(Path(state[0]).read_bytes()) if played is not None else None,
        errors.append if errors is not None else lambda exc: None,
        fallback_beep=lambda: beeps.append(True) if beeps is not None else None,
        engine_factory=engine,
        preview_play_callback=lambda file: preview.append(Path(file).read_bytes()) if preview is not None else None,
    )
    return Path(state[0])


class TtsOutputIntegrityTests(unittest.TestCase):
    def test_preview_does_not_suppress_the_next_real_notification(self):
        with tempfile.TemporaryDirectory() as directory:
            target, preview = Path(directory) / "alert.wav", Path(directory) / "preview.wav"
            target.write_bytes(wave_data("saved"))
            preview.write_bytes(wave_data("draft"))
            plays = []
            sound = SimpleNamespace(
                SND_FILENAME=1, SND_ASYNC=2, MB_ICONASTERISK=4,
                PlaySound=lambda path, flags: plays.append(Path(path).read_bytes()),
                MessageBeep=lambda flag: self.fail("Unexpected beep"),
            )
            namespace = {"TTS_FILE": str(target), "VOICE_ENABLED": True, "_last_play_time": 0,
                         "winsound": sound, "time": SimpleNamespace(monotonic=lambda: 10)}
            self.assertEqual(play_alert_namespace_runtime(namespace, force=True, tts_file=str(preview)), "played")
            self.assertEqual(play_alert_namespace_runtime(namespace), "played")
            self.assertEqual(plays, [wave_data("draft"), wave_data("saved")])

    def test_invalid_generation_preserves_good_audio_and_reports_failure(self):
        for output in (None, b"", b"invalid wave", wave_data()[:-2]):
            with self.subTest(output=output), tempfile.TemporaryDirectory() as directory:
                target = Path(directory) / "alert.wav"
                target.write_bytes(wave_data())
                original = target.read_bytes()

                class BrokenEngine(WaveEngine):
                    def save_to_file(self, text, path):
                        if output is not None:
                            Path(path).write_bytes(output)

                tasks = queue.Queue()
                tasks.put(("changed", True, True))
                played, errors, beeps = [], [], []
                drain_tts(tasks, target, engine=BrokenEngine, played=played, errors=errors, beeps=beeps)
                self.assertEqual(target.read_bytes(), original)
                self.assertEqual(played, [])
                self.assertEqual(len(errors), 1)
                self.assertEqual(beeps, [True])
                self.assertEqual(list(Path(directory).glob("*.tmp.wav")), [])

    def test_preview_preserves_pending_saved_generation_and_current_cache(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "alert_2.wav"
            target.write_bytes(wave_data())
            tasks = queue.Queue()
            enqueue_tts_request(tasks, "new saved", force=True)
            enqueue_tts_request(tasks, "draft", force=True, play_after=True, preview=True)
            self.assertEqual(tasks.qsize(), 2)
            preview = []
            final = drain_tts(tasks, target, preview=preview)
            self.assertEqual(final, target)
            self.assertEqual(target.read_bytes(), wave_data("new saved"))
            self.assertEqual(preview, [wave_data("draft")])
            self.assertEqual(tasks.unfinished_tasks, 0)

    def test_invalid_existing_cache_is_regenerated(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "alert.wav"
            target.write_bytes(b"")
            tasks = queue.Queue()
            tasks.put(("saved", False, True))
            played = []
            drain_tts(tasks, target, played=played)
            self.assertEqual(played, [wave_data("saved")])

    def test_replacement_failure_keeps_previous_wave(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "alert.wav"
            target.write_bytes(wave_data())
            with patch("sms_core.tts_runtime.os.replace", side_effect=PermissionError("locked")):
                with self.assertRaises(PermissionError):
                    generate_tts_file("draft", str(target), directory, threading.Lock(), WaveEngine)
            self.assertEqual(target.read_bytes(), wave_data())
            self.assertEqual(list(Path(directory).glob("*.tmp.wav")), [])


class TkQualityTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        self.root.withdraw()
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.addCleanup(self.destroy_root)
        self.errors = []
        self.root.report_callback_exception = lambda *args, errors=self.errors: errors.append(args)
        self.root.tk.eval("set ::background_errors {}; proc bgerror {message} {lappend ::background_errors $message}")

    def destroy_root(self):
        try:
            self.root.destroy()
        except tk.TclError:
            pass
        finally:
            self.root = None

    def namespace(self):
        config = configparser.ConfigParser(interpolation=None)
        config["ui"] = {"voice_enabled": "1"}
        path = str(Path(self.temp.name) / "config.ini")
        lock = threading.RLock()
        self.assertTrue(safe_save_config_runtime(config=config, config_file=path, config_lock=lock))
        namespace = {"root": self.root, "config": config, "CONFIG_FILE": path, "CONFIG_LOCK": lock,
                     "log_file_only": lambda _message: None}
        self.addCleanup(namespace.clear)
        return namespace

    def test_preview_then_cancel_does_not_change_next_notification(self):
        target = Path(self.temp.name) / "alert.wav"
        target.write_bytes(wave_data("saved"))
        tasks = queue.Queue()
        namespace = {"VOICE_TEXT": "saved", "DEFAULT_VOICE_TEXT": "default", "TTS_REQ_Q": tasks,
                     "ensure_tts_worker": lambda: None, "log_file_only": lambda *_args: None}
        config = configparser.ConfigParser(interpolation=None)
        config["ui"] = {"voice_text": "saved"}
        saves = []
        open_voice_text_dialog_runtime(
            self.root, "saved", config=config, safe_save=lambda: saves.append(True),
            set_voice_text=lambda text: namespace.__setitem__("VOICE_TEXT", text),
            generate_voice=lambda **kwargs: generate_alert_voice_namespace_runtime(namespace, **kwargs),
            system_ui=lambda *_args: None, center_window=lambda *_args: None, open_dialog=open_voice_text_dialog,
        )
        window = next(w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel))
        window.withdraw()
        editor = next(w for w in window.winfo_children() if isinstance(w, tk.Text))
        editor.delete("1.0", "end")
        editor.insert("1.0", "draft")
        buttons = {w.cget("text"): w for w in window.winfo_children() if isinstance(w, tk.Button)}
        buttons["试听"].invoke()
        buttons["取消"].invoke()
        previews, plays = [], []
        drain_tts(tasks, target, preview=previews)
        play_alert_runtime(
            voice_enabled=True, tts_file=str(target), get_last_play_time=lambda: -100,
            set_last_play_time=lambda _value: None, monotonic=lambda: 100,
            play_sound=lambda file, flags: plays.append(Path(file).read_bytes()),
            beep=lambda *_args: self.fail("Unexpected beep"), filename_flag=1, async_flag=2, beep_flag=0,
        )
        self.assertEqual(previews, [wave_data("draft")])
        self.assertEqual(plays, [wave_data("saved")])
        self.assertEqual(config.get("ui", "voice_text"), "saved")
        self.assertEqual(namespace["VOICE_TEXT"], "saved")
        self.assertEqual(saves, [])
        self.assertEqual(self.errors, [])

    def test_named_mutex_wait_keeps_tk_alive_and_restores_dialog_grab(self):
        namespace = self.namespace()
        acquired, release = threading.Event(), threading.Event()
        lock_errors = []

        def hold_mutex():
            mutex, result = acquire_named_mutex_lock(config_mutex_name(namespace["CONFIG_FILE"]))
            if not mutex:
                lock_errors.append(result)
                acquired.set()
                return
            acquired.set()
            try:
                release.wait(2)
            finally:
                release_named_mutex_lock(mutex)

        holder = threading.Thread(target=hold_mutex)
        holder.start()
        self.assertTrue(acquired.wait(1))
        if lock_errors:
            holder.join(3)
            self.skipTest("Windows named mutex unavailable")
        dialog = tk.Toplevel(self.root)
        dialog.grab_set()
        beats, applied = [], []
        self.root.after(40, lambda: beats.append(namespace.get("_CONFIG_SAVE_ACTIVE")))
        self.root.after(250, release.set)
        started = time.monotonic()
        try:
            # Some existing callers hold the RLock across their transaction.
            with namespace["CONFIG_LOCK"]:
                result = toggle_voice_broadcast_runtime(
                    True, namespace["config"], lambda: safe_save_config_namespace_runtime(namespace),
                    applied.append, lambda: None, lambda *_args: None,
                )
        finally:
            release.set()
            holder.join(3)
        self.assertFalse(result)
        self.assertEqual(applied, [False])
        self.assertEqual(beats, [True])
        self.assertLess(time.monotonic() - started, 1.5)
        self.assertIs(self.root.grab_current(), dialog)
        self.assertFalse(namespace["_CONFIG_SAVE_ACTIVE"])
        self.assertEqual(self.errors, [])

    def test_in_process_lock_wait_also_keeps_tk_alive(self):
        namespace = self.namespace()
        acquired, release = threading.Event(), threading.Event()

        def hold():
            with namespace["CONFIG_LOCK"]:
                acquired.set()
                release.wait(2)

        holder = threading.Thread(target=hold)
        holder.start()
        self.assertTrue(acquired.wait(1))
        self.root.after(80, release.set)
        started = time.monotonic()
        try:
            self.assertTrue(safe_save_config_namespace_runtime(namespace))
        finally:
            release.set()
            holder.join(3)
        self.assertLess(time.monotonic() - started, 1)
        self.assertEqual(self.errors, [])

    def test_slow_failed_save_preserves_config_and_does_not_apply_toggle(self):
        namespace = self.namespace()
        original = Path(namespace["CONFIG_FILE"]).read_bytes()
        beats, applied = [], []

        def fail_replace(*_args):
            time.sleep(0.2)
            raise PermissionError("synthetic write failure")

        def save(**kwargs):
            return safe_save_config_runtime(**kwargs, replace_file=fail_replace)

        self.root.after(40, lambda: beats.append(True))
        # Recording kwargs would retain the Tk-owned waiter's bound run_io.
        with patch("sms_ui.app_infrastructure_namespace_runtime.safe_save_config_runtime", new=save):
            result = toggle_voice_broadcast_runtime(
                True, namespace["config"], lambda: safe_save_config_namespace_runtime(namespace),
                applied.append, lambda: None, lambda *_args: None,
            )
        self.assertTrue(result)
        self.assertEqual(applied, [])
        self.assertEqual(beats, [True])
        self.assertEqual(Path(namespace["CONFIG_FILE"]).read_bytes(), original)
        self.assertEqual(namespace["config"].get("ui", "voice_enabled"), "1")
        self.assertEqual(self.errors, [])

    def test_waiting_ui_save_does_not_lose_changes_when_previous_writer_commits(self):
        namespace = self.namespace()
        acquired, release = threading.Event(), threading.Event()
        outcomes = []

        def acquire(name, timeout_ms):
            acquired.set()
            release.wait(2)
            return acquire_named_mutex_lock(name, timeout_ms)

        def previous_writer():
            outcomes.append(safe_save_config_runtime(
                config=namespace["config"], config_file=namespace["CONFIG_FILE"],
                config_lock=namespace["CONFIG_LOCK"], acquire_process_lock=acquire,
            ))

        holder = threading.Thread(target=previous_writer)
        holder.start()
        self.assertTrue(acquired.wait(1))
        self.root.after(60, release.set)
        try:
            result = toggle_voice_broadcast_runtime(
                True, namespace["config"], lambda: safe_save_config_namespace_runtime(namespace),
                lambda _enabled: None, lambda: None, lambda *_args: None,
            )
        finally:
            release.set()
            holder.join(3)
        self.assertEqual(outcomes, [True])
        self.assertFalse(result)
        stored = configparser.ConfigParser(interpolation=None)
        stored.read(namespace["CONFIG_FILE"], encoding="utf-8")
        self.assertEqual(stored.get("ui", "voice_enabled"), "0")
        self.assertEqual(namespace["config"].get("ui", "voice_enabled"), "0")
        self.assertEqual(self.errors, [])

    def test_reload_and_exit_wait_for_save_commit(self):
        namespace = self.namespace()
        namespace.update(is_exiting=False, messagebox=None, serial_running=False, TK_SHUTDOWN=threading.Event(),
                         serial_stop_event=threading.Event(), serial_wakeup_event=threading.Event(), TTS_STOP=threading.Event(),
                         safe_set_events=lambda *_a: None, stop_cloud_control=lambda **_k: None,
                         safe_close_serial=lambda: None, stop_tray_icon=lambda **_k: None,
                         flush_log_queue=lambda *_a, **_k: None, FILE_LOG_Q=queue.Queue(), third_push_stop=threading.Event(),
                         THIRD_PUSH_Q=queue.Queue(), file_log_stop=threading.Event())
        observed = []

        def save(**kwargs):
            actual = kwargs["run_io"]

            def slow_io(operation):
                return actual(lambda: (time.sleep(0.15), operation())[1])

            return safe_save_config_runtime(**{**kwargs, "run_io": slow_io})

        def during_save():
            self.assertFalse(reload_shared_ui_config_namespace_runtime(namespace))
            observed.append(cleanup_and_exit_namespace_runtime(
                namespace, cleanup_app_runtime=lambda **_kwargs: observed.append(namespace["_CONFIG_SAVE_ACTIVE"]),
            ))

        self.root.after(30, during_save)
        with patch("sms_ui.app_infrastructure_namespace_runtime.safe_save_config_runtime", new=save):
            self.assertTrue(safe_save_config_namespace_runtime(namespace))
        self.root.update()
        self.assertEqual(observed, ["saving_config", False])
        self.assertEqual(self.errors, [])

    def test_background_rebind_and_ui_save_keep_independent_commit_results(self):
        for rebind_fails in (False, True):
            for ui_fails in (False, True):
                with self.subTest(rebind_fails=rebind_fails, ui_fails=ui_fails):
                    namespace = self.namespace()
                    config = namespace["config"]
                    config["serial"] = {"mode": "Manual", "port": "COM901", "baud": "115200", "unknown": "keep"}
                    self.assertTrue(safe_save_config_namespace_runtime(namespace))
                    entered, release = threading.Event(), threading.Event()
                    results, applied, hints = [], [], []
                    namespace.update(
                        MODE="Manual", PORT="COM901", BAUD=115200,
                        find_luat_best_port=lambda: ("COM902", "synthetic"),
                        list_ports=SimpleNamespace(comports=lambda: []),
                        choose_manual_rebind_candidate=lambda *_a, **_k: SimpleNamespace(
                            found=True, device="COM902", description="synthetic"),
                        safe_save_config=lambda **kwargs: safe_save_config_namespace_runtime(namespace, **kwargs),
                        system_ui=lambda *_a: None, set_status=lambda *_a: None,
                        serial_wakeup_event=threading.Event(),
                        _rebind_hint_notice=SimpleNamespace(reset=lambda: hints.append(True)),
                        manual_rebind_hint=lambda *_a: "synthetic rebind",
                    )

                    def save(**kwargs):
                        background = threading.current_thread() is not threading.main_thread()

                        def replace(source, target):
                            if background:
                                entered.set()
                                if not release.wait(2):
                                    raise TimeoutError("Isolated UI did not release the writer")
                            if (rebind_fails if background else ui_fails):
                                raise PermissionError("synthetic config write failure")
                            Path(source).replace(target)

                        return safe_save_config_runtime(**kwargs, replace_file=replace)

                    def rebind():
                        results.append(try_rebind_manual_port_namespace_runtime(namespace, "synthetic"))

                    thread = threading.Thread(target=rebind)
                    try:
                        with patch("sms_ui.app_infrastructure_namespace_runtime.safe_save_config_runtime", new=save):
                            thread.start()
                            self.assertTrue(entered.wait(1))
                            self.root.after(40, release.set)
                            result = toggle_voice_broadcast_runtime(
                                True, config, namespace["safe_save_config"], applied.append,
                                lambda: None, lambda *_a: None,
                            )
                    finally:
                        release.set()
                        thread.join(3)
                    self.assertFalse(thread.is_alive())
                    disk = configparser.ConfigParser(interpolation=None)
                    disk.read(namespace["CONFIG_FILE"], encoding="utf-8")
                    expected_port = "COM901" if rebind_fails else "COM902"
                    self.assertEqual(results, [not rebind_fails])
                    self.assertEqual(namespace["PORT"], expected_port)
                    self.assertEqual(config.get("serial", "port"), expected_port)
                    self.assertEqual(disk.get("serial", "port"), expected_port)
                    self.assertEqual(disk.get("serial", "unknown"), "keep")
                    self.assertEqual(disk.get("ui", "voice_enabled"), "1" if ui_fails else "0")
                    self.assertEqual(config.get("ui", "voice_enabled"), "1" if ui_fails else "0")
                    self.assertEqual(result, ui_fails)
                    self.assertEqual(applied, [] if ui_fails else [False])
                    self.assertEqual(hints, [] if rebind_fails else [True])
                    self.assertEqual(self.errors, [])

    def test_save_ui_setup_failure_does_not_leave_windows_disabled(self):
        namespace = self.namespace()
        original = Path(namespace["CONFIG_FILE"]).read_bytes()
        def fail_after(*_args):
            raise tk.TclError("synthetic timer failure")

        with patch.object(self.root, "after", new=fail_after):
            self.assertFalse(safe_save_config_namespace_runtime(namespace))
        if self.root.tk.call("tk", "windowingsystem") == "win32":
            self.assertFalse(self.root.attributes("-disabled"))
        self.assertFalse(namespace["_CONFIG_SAVE_ACTIVE"])
        self.assertEqual(Path(namespace["CONFIG_FILE"]).read_bytes(), original)

    def test_visible_progress_animation_stops_when_root_is_destroyed(self):
        ShutdownProgress(self.root)
        self.root.destroy()
        self.root.tk.call("after", 50)
        self.root.tk.call("update")
        self.assertFalse(self.root.tk.call("after", "info"))
        self.assertFalse(self.root.tk.globalgetvar("background_errors"))
        self.assertEqual(self.errors, [])

    def test_visible_progress_animation_stops_when_progress_window_is_destroyed(self):
        view = ShutdownProgress(self.root)
        view.window.destroy()
        self.assertTrue(self.root.winfo_exists())
        self.assertFalse(self.root.tk.call("after", "info"))
        self.root.update()
        self.assertFalse(self.root.tk.globalgetvar("background_errors"))
        self.assertEqual(self.errors, [])

    def test_save_wait_releases_tk_variable_on_ui_thread(self):
        gc.collect()
        enabled = gc.isenabled()
        gc.disable()
        references, finalized = [], []
        ui_thread = threading.get_ident()

        class TrackedVariable(tk.BooleanVar):
            def __init__(variable, *args, **kwargs):
                super().__init__(*args, **kwargs)
                references.append(weakref.ref(variable))

            def __del__(variable):
                finalized.append(threading.get_ident())
                super().__del__()

        try:
            ready = threading.Event()
            with patch("sms_ui.config_save_runtime.tk.BooleanVar", TrackedVariable):
                with UiConfigSave({"root": self.root}) as waiter:
                    self.root.after(30, ready.set)
                    self.assertTrue(waiter.wait_for(ready.is_set))
            self.assertEqual(len(references), 1)
            self.assertIsNone(references[0]())
            self.assertEqual(finalized, [ui_thread])
            self.assertEqual(self.errors, [])
        finally:
            if enabled:
                gc.enable()
            gc.collect()

    def test_save_releases_waiter_without_cyclic_collection(self):
        gc.collect()
        enabled = gc.isenabled()
        gc.disable()
        references = []

        class TrackedWaiter(UiConfigSave):
            def __init__(waiter, namespace):
                super().__init__(namespace)
                references.append(weakref.ref(waiter))

        try:
            for fails in (False, True):
                with self.subTest(fails=fails):
                    references.clear()
                    namespace = self.namespace()

                    def replace(source, target):
                        if fails:
                            raise PermissionError("synthetic failed write")
                        Path(source).replace(target)

                    def save(**kwargs):
                        return safe_save_config_runtime(**kwargs, replace_file=replace)

                    with patch("sms_ui.config_save_runtime.UiConfigSave", TrackedWaiter), \
                         patch("sms_ui.app_infrastructure_namespace_runtime.safe_save_config_runtime", new=save):
                        self.assertEqual(safe_save_config_namespace_runtime(namespace), not fails)
                    self.assertEqual(len(references), 1)
                    self.assertIsNone(references[0]())
            self.assertEqual(self.errors, [])
        finally:
            if enabled:
                gc.enable()
            gc.collect()

    def check_destroy_during_save(self, defer_exit):
        namespace = self.namespace()
        namespace["config"].set("ui", "voice_enabled", "0")
        entered, release = threading.Event(), threading.Event()
        deferred_results, deferred_calls = [], []

        def replace(source, target):
            entered.set()
            if not release.wait(2):
                raise TimeoutError("Isolated UI did not release the write")
            Path(source).replace(target)

        def save(**kwargs):
            return safe_save_config_runtime(**kwargs, replace_file=replace)

        def destroy_during_write():
            if not entered.is_set():
                self.root.after(10, destroy_during_write)
                return
            if defer_exit:
                deferred_results.append(cleanup_and_exit_namespace_runtime(
                    namespace, cleanup_app_runtime=lambda **_kwargs: deferred_calls.append(True),
                ))
            # Let both the poll and delayed progress display become due.
            time.sleep(0.14)
            self.root.destroy()
            release.set()

        self.root.after(10, destroy_during_write)
        try:
            with patch("sms_ui.app_infrastructure_namespace_runtime.safe_save_config_runtime", new=save):
                self.assertTrue(safe_save_config_namespace_runtime(namespace))
        finally:
            release.set()
        self.assertTrue(entered.is_set())
        self.assertIn("voice_enabled = 0", Path(namespace["CONFIG_FILE"]).read_text(encoding="utf-8"))
        self.assertEqual(namespace["config"].get("ui", "voice_enabled"), "0")
        self.assertFalse(namespace["_CONFIG_SAVE_ACTIVE"])
        self.assertNotIn("_CONFIG_SAVE_THREAD", namespace)
        self.assertNotIn("_CONFIG_SAVE_DEFERRED_ACTION", namespace)
        self.assertEqual(deferred_results, ["saving_config"] if defer_exit else [])
        self.assertEqual(deferred_calls, [])
        self.assertEqual(self.errors, [])
        self.assertFalse(self.root.tk.call("after", "info"))
        self.root.tk.call("update")
        self.assertFalse(self.root.tk.globalgetvar("background_errors"))

    def test_destroy_during_disk_save_cleans_poll_and_accounts_for_commit(self):
        self.check_destroy_during_save(defer_exit=False)

    def test_destroy_during_disk_save_discards_deferred_exit_without_raising(self):
        self.check_destroy_during_save(defer_exit=True)

    def test_destroy_during_lock_wait_does_not_start_a_disk_write(self):
        namespace = self.namespace()
        original = Path(namespace["CONFIG_FILE"]).read_bytes()
        namespace["config"].set("ui", "voice_enabled", "0")
        acquired, release = threading.Event(), threading.Event()

        def hold():
            with namespace["CONFIG_LOCK"]:
                acquired.set()
                release.wait(2)

        holder = threading.Thread(target=hold)
        holder.start()
        try:
            self.assertTrue(acquired.wait(1))
            self.root.after(30, self.root.destroy)
            self.assertFalse(safe_save_config_namespace_runtime(namespace))
        finally:
            release.set()
            holder.join(3)
        self.assertEqual(Path(namespace["CONFIG_FILE"]).read_bytes(), original)
        self.assertFalse(namespace["_CONFIG_SAVE_ACTIVE"])
        self.assertNotIn("_CONFIG_SAVE_THREAD", namespace)
        self.assertEqual(self.errors, [])
        self.assertFalse(self.root.tk.call("after", "info"))
        self.root.tk.call("update")
        self.assertFalse(self.root.tk.globalgetvar("background_errors"))


class HttpDeadlineTests(unittest.TestCase):
    def test_normal_redirected_and_chunked_success_responses_are_read_completely(self):
        body = json.dumps({"errcode": 0, "padding": "a" * 12000}).encode()

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *_args):
                pass

            def do_GET(self):
                if self.path == "/redirect":
                    self.send_response(302)
                    self.send_header("Location", "/normal")
                    self.end_headers()
                    return
                self.send_response(200)
                if self.path == "/chunked":
                    self.send_header("Transfer-Encoding", "chunked")
                    self.end_headers()
                    for start in range(0, len(body), 100):
                        chunk = body[start:start + 100]
                        self.wfile.write(f"{len(chunk):x}\r\n".encode() + chunk + b"\r\n")
                    self.wfile.write(b"0\r\n\r\n")
                else:
                    self.send_header("Content-Length", str(len(body)))
                    self.end_headers()
                    self.wfile.write(body)

        server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        server.daemon_threads = True
        worker = threading.Thread(target=lambda: server.serve_forever(poll_interval=0.01))
        worker.start()
        try:
            with patch("urllib.request.getproxies", return_value={}):
                for suffix in ("normal", "redirect", "chunked"):
                    with self.subTest(suffix=suffix):
                        self.assertEqual(http_request(
                            f"http://127.0.0.1:{server.server_port}/{suffix}", method="GET", timeout=1,
                        ), (True, 200, body.decode()))
        finally:
            server.shutdown()
            server.server_close()
            worker.join(2)

    def test_slow_body_headers_and_chunked_response_have_one_deadline(self):
        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *_args):
                pass

            def do_POST(self):
                if self.path == "/headers":
                    data = b"HTTP/1.1 200 OK\r\nX-Slow: " + b"x" * 40 + b"\r\nContent-Length: 2\r\n\r\n{}"
                elif self.path == "/chunked":
                    self.send_response(200)
                    self.send_header("Transfer-Encoding", "chunked")
                    self.end_headers()
                    data = b"1\r\nx\r\n" * 20 + b"0\r\n\r\n"
                else:
                    self.send_response(200)
                    self.send_header("Content-Length", "40")
                    self.end_headers()
                    data = b"x" * 40
                try:
                    for byte in data:
                        self.wfile.write(bytes([byte]))
                        self.wfile.flush()
                        time.sleep(0.025)
                except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
                    pass

        server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        server.daemon_threads = True
        worker = threading.Thread(target=lambda: server.serve_forever(poll_interval=0.01))
        worker.start()
        try:
            with patch("urllib.request.getproxies", return_value={}):
                for suffix in ("body", "headers", "chunked"):
                    with self.subTest(suffix=suffix):
                        started = time.monotonic()
                        ok, code, _body = http_request(f"http://127.0.0.1:{server.server_port}/{suffix}", timeout=0.12)
                        self.assertFalse(ok)
                        self.assertIsNone(code)
                        self.assertLess(time.monotonic() - started, 0.7)
        finally:
            server.shutdown()
            server.server_close()
            worker.join(2)

    def test_https_keeps_certificate_and_hostname_checks(self):
        connection = _HTTPSConnection("example.test", deadline=_Deadline(1))
        self.assertTrue(connection._context.check_hostname)
        self.assertEqual(connection._context.verify_mode, ssl.CERT_REQUIRED)
        connection.close()


class RecordingIoTests(unittest.IsolatedAsyncioTestCase):
    async def test_uploaded_history_scans_once_without_stalling_the_loop(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for index in range(20):
                path = root / f"synthetic-{index}.amr"
                path.write_bytes(b"#!AMR\n" + b"x" * 16)
                path.with_suffix(".amr.json").write_text(json.dumps({
                    "recording_id": f"synthetic-{index}", "imei": "111111111111111", "path": path.name,
                    "size": 22, "started_at": 1, "upload_status": "uploaded",
                }), encoding="utf-8")
            repository = CallRecordingRepository(directory)
            uploader = CloudCallRecordingUploader(repository)
            original = repository._read_metadata
            calls, beats = [], []

            def slow_read(path):
                calls.append(threading.get_ident())
                time.sleep(0.015)
                return original(path)

            async def no_send(*_args):
                self.fail("Uploaded recordings must not be sent again")

            async def heartbeat():
                for _ in range(15):
                    beats.append(time.monotonic())
                    await asyncio.sleep(0.01)

            with patch.object(repository, "_read_metadata", side_effect=slow_read):
                await asyncio.gather(heartbeat(), uploader._drain(
                    object(), no_send, lambda: {"imei": "111111111111111"}, lambda _ws: True, lambda: True,
                ))
            self.assertEqual(len(calls), 20)
            self.assertTrue(all(ident != threading.get_ident() for ident in calls))
            self.assertLess(max(b - a for a, b in zip(beats, beats[1:])), 0.2)
            await uploader.stop()

    async def test_cancelled_io_settles_before_upload_is_released(self):
        entered, release = threading.Event(), threading.Event()
        uploader = CloudCallRecordingUploader(None)
        completed = []

        def write_state():
            entered.set()
            release.wait(2)
            completed.append(True)

        task = asyncio.create_task(uploader._io(write_state))
        self.assertTrue(await asyncio.to_thread(entered.wait, 1))
        try:
            task.cancel()
            await asyncio.sleep(0.02)
            self.assertFalse(task.done())
            task.cancel()
            await asyncio.sleep(0.02)
            self.assertFalse(task.done())
        finally:
            release.set()
        with self.assertRaises(asyncio.CancelledError):
            await asyncio.wait_for(task, 1)
        self.assertEqual(completed, [True])


if __name__ == "__main__":
    unittest.main()
