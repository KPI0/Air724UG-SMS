import unittest

from sms_ui.audio_namespace_runtime import (
    ensure_tts_worker_namespace_runtime,
    generate_alert_voice_namespace_runtime,
    play_alert_namespace_runtime,
    set_tts_file_namespace_runtime,
    tts_worker_namespace_runtime,
)


class FakeWinsound:
    SND_FILENAME = 1
    SND_ASYNC = 2
    MB_ICONASTERISK = 4

    def __init__(self, calls):
        self.calls = calls

    def PlaySound(self, *args):
        self.calls.append(("play", args))

    def MessageBeep(self, *args):
        self.calls.append(("beep", args))


class FakeThreading:
    class Thread:
        pass


class FakeTime:
    @staticmethod
    def monotonic():
        return 10.0


class AudioNamespaceRuntimeTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        return {
            "calls": calls,
            "TTS_STOP": "stop",
            "TTS_REQ_Q": "queue",
            "TTS_LOCK": "lock",
            "TTS_FILE": "alert.wav",
            "TTS_DIR": "tts",
            "TTS_THREAD": "thread",
            "DEFAULT_VOICE_TEXT": "default",
            "VOICE_TEXT": "voice",
            "VOICE_ENABLED": True,
            "_last_play_time": 0.0,
            "_tts_worker": lambda: calls.append(("worker",)),
            "_set_tts_file": lambda path: calls.append(("set_file", path)),
            "play_alert": lambda **kwargs: calls.append(("play_alert", kwargs)),
            "ensure_tts_worker": lambda: calls.append(("ensure",)),
            "log_file_only": lambda message: calls.append(("log", message)),
            "winsound": FakeWinsound(calls),
            "threading": FakeThreading,
            "time": FakeTime,
        }

    def test_tts_worker_namespace_runtime_forwards_callbacks(self):
        namespace = self.make_namespace()
        calls = []

        result = tts_worker_namespace_runtime(
            namespace,
            worker_loop=lambda *args, **kwargs: calls.append((args, kwargs)) or "worked",
        )

        self.assertEqual(result, "worked")
        args, kwargs = calls[0]
        self.assertEqual(args[:4], ("stop", "queue", "lock", args[3]))
        self.assertEqual(args[3](), "alert.wav")
        args[4]("next.wav")
        args[7](force=True)
        args[8](RuntimeError("boom"))
        kwargs["fallback_beep"]()
        self.assertIn(("set_file", "next.wav"), namespace["calls"])
        self.assertIn(("play_alert", {"force": True}), namespace["calls"])
        self.assertIn(("log", "TTS 生成失败，使用系统声音兜底：boom"), namespace["calls"])
        self.assertIn(("beep", (4,)), namespace["calls"])

    def test_set_tts_file_namespace_runtime_updates_state(self):
        namespace = self.make_namespace()

        set_tts_file_namespace_runtime(namespace, "next.wav")

        self.assertEqual(namespace["TTS_FILE"], "next.wav")

    def test_ensure_tts_worker_namespace_runtime_forwards_thread_state(self):
        namespace = self.make_namespace()
        calls = []

        result = ensure_tts_worker_namespace_runtime(
            namespace,
            ensure_runtime=lambda **kwargs: calls.append(kwargs) or "started",
        )

        self.assertEqual(result, "started")
        forwarded = calls[0]
        self.assertEqual(forwarded["get_thread"](), "thread")
        forwarded["set_thread"]("next_thread")
        self.assertEqual(namespace["TTS_THREAD"], "next_thread")
        self.assertEqual(forwarded["stop_event"], "stop")
        self.assertIs(forwarded["thread_factory"], FakeThreading.Thread)
        forwarded["log_error"](RuntimeError("bad"))
        self.assertIn(("log", "TTS 线程启动失败：bad"), namespace["calls"])

    def test_generate_alert_voice_namespace_runtime_forwards_current_voice_text(self):
        namespace = self.make_namespace()
        calls = []

        result = generate_alert_voice_namespace_runtime(
            namespace,
            force=True,
            text="manual",
            play_after=True,
            generate_runtime=lambda **kwargs: calls.append(kwargs) or "queued",
        )

        self.assertEqual(result, "queued")
        forwarded = calls[0]
        self.assertTrue(forwarded["force"])
        self.assertEqual(forwarded["text"], "manual")
        self.assertEqual(forwarded["get_voice_text"](), "voice")
        self.assertEqual(forwarded["default_text"], "default")
        forwarded["ensure_worker"]()
        forwarded["log_queue_full"]()
        self.assertIn(("ensure",), namespace["calls"])
        self.assertIn(("log", "⚠️ TTS 请求队列已满，已丢弃一次生成请求"), namespace["calls"])

    def test_play_alert_namespace_runtime_updates_last_play_time(self):
        namespace = self.make_namespace()
        calls = []

        result = play_alert_namespace_runtime(
            namespace,
            force=True,
            play_runtime=lambda **kwargs: calls.append(kwargs) or (
                kwargs["set_last_play_time"](kwargs["monotonic"]()) or "played"
            ),
        )

        self.assertEqual(result, "played")
        forwarded = calls[0]
        self.assertTrue(forwarded["force"])
        self.assertTrue(forwarded["voice_enabled"])
        self.assertEqual(forwarded["tts_file"], "alert.wav")
        self.assertEqual(forwarded["filename_flag"] | forwarded["async_flag"], 3)
        self.assertEqual(forwarded["beep_flag"], 4)
        self.assertEqual(namespace["_last_play_time"], 10.0)


if __name__ == "__main__":
    unittest.main()
