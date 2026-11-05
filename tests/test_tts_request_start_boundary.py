import configparser
import os
import queue
import tempfile
import threading
import unittest
from types import SimpleNamespace

from sms_core.config_runtime import safe_save_config_runtime
from sms_ui.audio_namespace_bindings import install_audio_namespace_bindings
from sms_ui.settings_namespace_bindings import install_settings_namespace_bindings


class SyntheticThread:
    def __init__(self, *, reject=False, **_options):
        self.reject = reject
        self.alive = False

    def start(self):
        if self.reject:
            raise RuntimeError("synthetic thread start rejected")
        self.alive = True

    def is_alive(self):
        return self.alive


class TtsRequestStartBoundaryTests(unittest.TestCase):
    def namespace(self, *, rejected=False, stopped=False, alive=False, config_file=None):
        config = configparser.ConfigParser(interpolation=None)
        config["ui"] = {"voice_text": "old alert"}
        stop_event = threading.Event()
        if stopped:
            stop_event.set()
        thread = SyntheticThread()
        thread.alive = alive
        callbacks = {}
        messages = []
        logs = []
        namespace = {
            "TTS_THREAD": thread if alive else None,
            "TTS_STOP": stop_event, "TTS_REQ_Q": queue.Queue(maxsize=2),
            "TTS_LOCK": threading.Lock(), "TTS_FILE": "unused.wav", "TTS_DIR": "unused",
            "VOICE_TEXT": "old alert", "DEFAULT_VOICE_TEXT": "default alert",
            "VOICE_ENABLED": True, "root": object(), "config": config,
            "threading": SimpleNamespace(Thread=lambda **options: SyntheticThread(reject=rejected, **options)),
            "log_file_only": logs.append,
            "system_ui": lambda message, tag="normal": messages.append(message),
            "center_window": lambda *_args: None,
            "register_config_sync_refresher": lambda *_args: None,
            "_ui_open_voice_text_dialog": lambda parent, text, preview, save, center, **options:
                callbacks.update(preview=preview, save=save),
        }
        namespace["safe_save_config"] = lambda: safe_save_config_runtime(
            config=config, config_file=config_file, config_lock=threading.RLock(), log_error=logs.append,
        ) if config_file is not None else True
        install_audio_namespace_bindings(namespace)
        install_settings_namespace_bindings(namespace)
        namespace["open_voice_text_dialog"]()
        return namespace, callbacks, messages

    def test_rejected_start_does_not_discard_or_add_pending_saved_alert(self):
        namespace, _, _ = self.namespace(rejected=True)
        queue_ = namespace["TTS_REQ_Q"]
        original = ("pending saved alert", True, False)
        queue_.put_nowait(original)

        result = namespace["generate_alert_voice"](force=True, text="replacement")

        self.assertEqual(result, "error")
        self.assertIsNone(namespace["TTS_THREAD"])
        self.assertEqual(queue_.get_nowait(), original)
        queue_.task_done()
        self.assertTrue(queue_.empty())
        self.assertEqual(queue_.unfinished_tasks, 0)

    def test_stopped_alive_worker_does_not_accept_another_request(self):
        namespace, _, _ = self.namespace(stopped=True, alive=True)

        result = namespace["generate_alert_voice"](force=True, text="after stop")

        self.assertEqual(result, "stopped")
        self.assertTrue(namespace["TTS_THREAD"].is_alive())
        self.assertTrue(namespace["TTS_REQ_Q"].empty())
        self.assertEqual(namespace["TTS_REQ_Q"].unfinished_tasks, 0)

    def test_preview_reports_rejected_start_without_queuing(self):
        namespace, callbacks, messages = self.namespace(rejected=True)

        callbacks["preview"]("synthetic preview")

        self.assertTrue(namespace["TTS_REQ_Q"].empty())
        self.assertEqual(namespace["TTS_REQ_Q"].unfinished_tasks, 0)
        self.assertEqual(messages, ["❌ 无法启动语音生成，请稍后重试"])

    def test_preview_reports_stopped_worker_without_queuing(self):
        namespace, callbacks, messages = self.namespace(stopped=True, alive=True)

        callbacks["preview"]("synthetic preview")

        self.assertTrue(namespace["TTS_REQ_Q"].empty())
        self.assertEqual(messages, ["❌ 语音生成已停止，请重启软件后重试"])

    def test_saved_text_remains_committed_when_generation_cannot_start(self):
        with tempfile.TemporaryDirectory() as temporary:
            config_file = os.path.join(temporary, "config.ini")
            namespace, callbacks, messages = self.namespace(rejected=True, config_file=config_file)

            saved = callbacks["save"]("new saved alert")

            self.assertTrue(saved)
            self.assertEqual(namespace["VOICE_TEXT"], "new saved alert")
            stored = configparser.ConfigParser(interpolation=None)
            self.assertEqual(stored.read(config_file, encoding="utf-8"), [config_file])
            self.assertEqual(stored.get("ui", "voice_text"), "new saved alert")
            self.assertTrue(namespace["TTS_REQ_Q"].empty())
            self.assertEqual(messages, [
                "🔊 已更新语音播报内容：new saved alert",
                "⚠️ 内容已保存，但无法启动语音生成，请稍后重试",
            ])

    def test_full_preview_queue_preserves_requests_and_reports_failure(self):
        namespace, callbacks, messages = self.namespace(alive=True)
        queue_ = namespace["TTS_REQ_Q"]
        original = [("saved alert", True, False), ("earlier preview", True, True, True)]
        for task in original:
            queue_.put_nowait(task)

        callbacks["preview"]("new preview")

        self.assertEqual(messages, ["❌ 语音生成队列已满，请稍后重试"])
        self.assertEqual([queue_.get_nowait() for _ in original], original)
        for _ in original:
            queue_.task_done()
        self.assertEqual(queue_.unfinished_tasks, 0)

    def test_retry_after_start_failure_queues_only_new_preview(self):
        namespace, callbacks, messages = self.namespace(rejected=True)
        callbacks["preview"]("failed preview")
        namespace["threading"] = SimpleNamespace(Thread=SyntheticThread)

        callbacks["preview"]("new preview")

        self.assertTrue(namespace["TTS_THREAD"].is_alive())
        queue_ = namespace["TTS_REQ_Q"]
        self.assertEqual(queue_.get_nowait(), ("new preview", True, True, True))
        queue_.task_done()
        self.assertTrue(queue_.empty())
        self.assertEqual(messages, ["❌ 无法启动语音生成，请稍后重试"])


if __name__ == "__main__":
    unittest.main()
