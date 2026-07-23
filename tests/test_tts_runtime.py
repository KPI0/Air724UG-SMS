import os
import queue
import tempfile
import threading
import unittest
from unittest.mock import patch

from sms_core.tts_runtime import (
    cleanup_tts_alt_files,
    clear_tts_queue,
    enqueue_tts_request,
    ensure_tts_worker_runtime,
    generate_alert_voice_runtime,
    generate_tts_file,
    instance_tts_file_path,
    normalize_voice_text,
    tts_worker_loop,
    tts_file_family,
)


class FakeEngine:
    def __init__(self):
        self.rate = None
        self.saved = None
        self.stopped = False

    def setProperty(self, name, value):
        if name == "rate":
            self.rate = value

    def save_to_file(self, text, path):
        self.saved = (text, path)
        with open(path, "wb") as file:
            file.write(b"wav")

    def runAndWait(self):
        pass

    def stop(self):
        self.stopped = True


class FakeThread:
    def __init__(self, *, alive=False):
        self.alive = alive
        self.started = False

    def is_alive(self):
        return self.alive

    def start(self):
        self.started = True
        self.alive = True


class FakeEvent:
    def __init__(self, set_value=False):
        self.set_value = set_value

    def is_set(self):
        return self.set_value


class TtsRuntimeTests(unittest.TestCase):
    def test_instance_tts_file_path_keeps_first_and_splits_later_instances(self):
        self.assertEqual(instance_tts_file_path("tts", 1), os.path.join("tts", "alert.wav"))
        self.assertEqual(instance_tts_file_path("tts", 3), os.path.join("tts", "alert_3.wav"))

    def test_normalize_voice_text_uses_default_for_blank_values(self):
        self.assertEqual(normalize_voice_text("", "default"), "default")
        self.assertEqual(normalize_voice_text(" hello ", "default"), "hello")

    def test_clear_and_enqueue_tts_request_debounces_queue(self):
        request_queue = queue.Queue()
        request_queue.put(("old", False, False))
        request_queue.put(("older", False, False))

        self.assertEqual(clear_tts_queue(request_queue), 2)
        self.assertEqual(request_queue.unfinished_tasks, 0)
        enqueue_tts_request(request_queue, "new", force=True, play_after=True)

        self.assertEqual(request_queue.get_nowait(), ("new", True, True))
        request_queue.task_done()

    def test_tts_worker_skips_malformed_request_and_balances_queue(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = os.path.join(tmp, "alert.wav")
            with open(target, "wb") as file:
                file.write(b"wav")

            request_queue = queue.Queue()
            request_queue.put(("malformed",))
            request_queue.put(("hello", False, False))
            errors = []

            class StopWhenDrained:
                def is_set(self):
                    return request_queue.empty() and request_queue.unfinished_tasks == 0

            tts_worker_loop(
                StopWhenDrained(),
                request_queue,
                threading.Lock(),
                lambda: target,
                lambda _path: None,
                tmp,
                "default",
                lambda **_kwargs: None,
                errors.append,
                engine_factory=lambda: FakeEngine(),
                poll_timeout=0,
            )

            self.assertEqual(len(errors), 1)
            self.assertEqual(request_queue.unfinished_tasks, 0)

    def test_tts_worker_survives_play_callback_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = os.path.join(tmp, "alert.wav")
            with open(target, "wb") as file:
                file.write(b"wav")

            request_queue = queue.Queue()
            request_queue.put(("hello", False, True))
            errors = []

            class StopWhenDrained:
                def is_set(self):
                    return request_queue.empty() and request_queue.unfinished_tasks == 0

            tts_worker_loop(
                StopWhenDrained(),
                request_queue,
                threading.Lock(),
                lambda: target,
                lambda _path: None,
                tmp,
                "default",
                lambda **_kwargs: (_ for _ in ()).throw(RuntimeError("play failed")),
                errors.append,
                engine_factory=lambda: FakeEngine(),
                poll_timeout=0,
            )

            self.assertEqual(len(errors), 1)
            self.assertIn("play failed", str(errors[0]))
            self.assertEqual(request_queue.unfinished_tasks, 0)

    def test_cleanup_tts_alt_files_keeps_current_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            current = os.path.join(tmp, "alert_alt_keep.wav")
            stale = os.path.join(tmp, "alert_alt_old.wav")
            other = os.path.join(tmp, "note.wav")
            for path in (current, stale, other):
                with open(path, "wb") as file:
                    file.write(b"x")

            self.assertEqual(cleanup_tts_alt_files(tmp, current), 1)
            self.assertTrue(os.path.exists(current))
            self.assertFalse(os.path.exists(stale))
            self.assertTrue(os.path.exists(other))

    def test_cleanup_tts_alt_files_does_not_delete_other_instance_family(self):
        with tempfile.TemporaryDirectory() as tmp:
            own_stale = os.path.join(tmp, "alert_2_alt_old.wav")
            other_active = os.path.join(tmp, "alert_3_alt_active.wav")
            for path in (own_stale, other_active):
                with open(path, "wb") as file:
                    file.write(b"x")

            self.assertEqual(
                cleanup_tts_alt_files(tmp, os.path.join(tmp, "alert_2.wav")),
                1,
            )
            self.assertFalse(os.path.exists(own_stale))
            self.assertTrue(os.path.exists(other_active))
            self.assertEqual(tts_file_family(other_active), "alert_3")

    def test_generate_tts_file_replaces_target(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = os.path.join(tmp, "alert.wav")
            engine = FakeEngine()

            new_path = generate_tts_file(
                "hello",
                target,
                tmp,
                threading.Lock(),
                engine_factory=lambda: engine,
            )

            self.assertEqual(new_path, target)
            self.assertEqual(engine.rate, 150)
            self.assertTrue(engine.stopped)
            with open(target, "rb") as file:
                self.assertEqual(file.read(), b"wav")

    def test_generate_tts_file_uses_unique_temp_paths_across_instances(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = os.path.join(tmp, "alert.wav")
            engines = [FakeEngine(), FakeEngine()]

            generate_tts_file(
                "one",
                target,
                tmp,
                threading.Lock(),
                engine_factory=lambda: engines[0],
                uuid_func=lambda: "instance_one",
            )
            generate_tts_file(
                "two",
                target,
                tmp,
                threading.Lock(),
                engine_factory=lambda: engines[1],
                uuid_func=lambda: "instance_two",
            )

            self.assertTrue(engines[0].saved[1].endswith(".instance_one.tmp.wav"))
            self.assertTrue(engines[1].saved[1].endswith(".instance_two.tmp.wav"))
            self.assertNotEqual(engines[0].saved[1], engines[1].saved[1])

    def test_generate_tts_file_uses_instance_family_for_permission_fallback(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = os.path.join(tmp, "alert_2.wav")
            engine = FakeEngine()
            real_replace = os.replace
            calls = []

            def replace_with_first_permission_error(src, dst):
                calls.append((src, dst))
                if len(calls) == 1:
                    raise PermissionError("busy")
                return real_replace(src, dst)

            with patch("sms_core.tts_runtime.os.replace", side_effect=replace_with_first_permission_error):
                result = generate_tts_file(
                    "hello",
                    target,
                    tmp,
                    threading.Lock(),
                    engine_factory=lambda: engine,
                    uuid_func=lambda: "token",
                )

            self.assertEqual(result, os.path.join(tmp, "alert_2_alt_token.wav"))
            self.assertTrue(os.path.exists(result))

    def test_ensure_tts_worker_runtime_starts_thread(self):
        state = {"thread": None}

        result = ensure_tts_worker_runtime(
            get_thread=lambda: state["thread"],
            set_thread=lambda thread: state.__setitem__("thread", thread),
            stop_event=FakeEvent(False),
            worker_target=lambda: None,
            thread_factory=lambda **_kwargs: FakeThread(),
            log_error=lambda exc: None,
        )

        self.assertEqual(result, "started")
        self.assertTrue(state["thread"].started)

    def test_ensure_tts_worker_runtime_skips_running_or_stopped(self):
        self.assertEqual(
            ensure_tts_worker_runtime(
                get_thread=lambda: FakeThread(alive=True),
                set_thread=lambda thread: None,
                stop_event=FakeEvent(False),
                worker_target=lambda: None,
                thread_factory=lambda **_kwargs: FakeThread(),
                log_error=lambda exc: None,
            ),
            "already_running",
        )
        self.assertEqual(
            ensure_tts_worker_runtime(
                get_thread=lambda: None,
                set_thread=lambda thread: None,
                stop_event=FakeEvent(True),
                worker_target=lambda: None,
                thread_factory=lambda **_kwargs: FakeThread(),
                log_error=lambda exc: None,
            ),
            "stopped",
        )

    def test_ensure_tts_worker_runtime_clears_published_thread_when_start_fails(self):
        state = {"thread": None}
        errors = []

        class BrokenThread(FakeThread):
            def start(self):
                raise RuntimeError("start failed")

        result = ensure_tts_worker_runtime(
            get_thread=lambda: state["thread"],
            set_thread=lambda thread: state.__setitem__("thread", thread),
            stop_event=FakeEvent(False),
            worker_target=lambda: None,
            thread_factory=lambda **_kwargs: BrokenThread(),
            log_error=errors.append,
        )

        self.assertEqual(result, "error")
        self.assertIsNone(state["thread"])
        self.assertEqual(len(errors), 1)
        self.assertEqual(str(errors[0]), "start failed")

    def test_generate_alert_voice_runtime_enqueues_normalized_request(self):
        request_queue = queue.Queue()
        calls = []

        result = generate_alert_voice_runtime(
            force=True,
            text=None,
            play_after=True,
            get_voice_text=lambda: " hello ",
            default_text="default",
            ensure_worker=lambda: calls.append("worker"),
            request_queue=request_queue,
            log_queue_full=lambda: calls.append("full"),
        )

        self.assertEqual(result, "queued")
        self.assertEqual(calls, ["worker"])
        self.assertEqual(request_queue.get_nowait(), ("hello", True, True))

    def test_generate_alert_voice_runtime_logs_full_queue(self):
        request_queue = queue.Queue(maxsize=1)
        request_queue.put_nowait(("existing", False, False))
        calls = []

        result = generate_alert_voice_runtime(
            get_voice_text=lambda: "hello",
            default_text="default",
            ensure_worker=lambda: None,
            request_queue=request_queue,
            log_queue_full=lambda: calls.append("full"),
            enqueue_request=lambda *_args, **_kwargs: (_ for _ in ()).throw(queue.Full()),
        )

        self.assertEqual(result, "full")
        self.assertEqual(calls, ["full"])



class GenerateTtsFileResilienceTests(unittest.TestCase):
    def test_stop_called_even_when_run_and_wait_raises(self):
        class BoomEngine(FakeEngine):
            def runAndWait(self):
                raise RuntimeError("sapi boom")

        engine = BoomEngine()
        with tempfile.TemporaryDirectory() as tmp:
            target = os.path.join(tmp, "alert.wav")
            with self.assertRaises(RuntimeError):
                generate_tts_file(
                    "hello",
                    target,
                    tmp,
                    threading.Lock(),
                    engine_factory=lambda: engine,
                )
        # stop() must still run so the SAPI engine is released.
        self.assertTrue(engine.stopped)

    def test_bare_filename_without_directory_does_not_crash(self):
        cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp:
            os.chdir(tmp)
            try:
                engine = FakeEngine()
                new_path = generate_tts_file(
                    "hello",
                    "alert.wav",  # no directory component
                    tmp,
                    threading.Lock(),
                    engine_factory=lambda: engine,
                )
                self.assertEqual(new_path, "alert.wav")
                self.assertTrue(os.path.exists(os.path.join(tmp, "alert.wav")))
            finally:
                os.chdir(cwd)


if __name__ == "__main__":
    unittest.main()
