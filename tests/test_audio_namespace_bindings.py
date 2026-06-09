import unittest
from unittest.mock import patch

import sms_ui.audio_namespace_bindings as bindings


class AudioNamespaceBindingsTests(unittest.TestCase):
    def test_install_registers_expected_names(self):
        namespace = {}

        result = bindings.install_audio_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        for name in (
            "_tts_worker",
            "_set_tts_file",
            "ensure_tts_worker",
            "generate_alert_voice",
            "play_alert",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_bindings_forward_namespace_and_arguments(self):
        namespace = {}
        bindings.install_audio_namespace_bindings(namespace)

        with patch.object(bindings, "tts_worker_namespace_runtime", return_value="worked") as worker, \
                patch.object(bindings, "set_tts_file_namespace_runtime", return_value="set") as set_file, \
                patch.object(bindings, "ensure_tts_worker_namespace_runtime", return_value="ensured") as ensure, \
                patch.object(bindings, "generate_alert_voice_namespace_runtime", return_value="generated") as generate, \
                patch.object(bindings, "play_alert_namespace_runtime", return_value="played") as play:
            self.assertEqual(namespace["_tts_worker"](), "worked")
            self.assertEqual(namespace["_set_tts_file"]("next.wav"), "set")
            self.assertEqual(namespace["ensure_tts_worker"](), "ensured")
            self.assertEqual(
                namespace["generate_alert_voice"](force=True, text="hello", play_after=True),
                "generated",
            )
            self.assertEqual(namespace["play_alert"](force=True), "played")

        worker.assert_called_once_with(namespace)
        set_file.assert_called_once_with(namespace, "next.wav")
        ensure.assert_called_once_with(namespace)
        self.assertEqual(generate.call_args.args, (namespace,))
        self.assertTrue(generate.call_args.kwargs["force"])
        self.assertEqual(generate.call_args.kwargs["text"], "hello")
        self.assertTrue(generate.call_args.kwargs["play_after"])
        play.assert_called_once_with(namespace, force=True)


if __name__ == "__main__":
    unittest.main()
