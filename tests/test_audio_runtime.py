import unittest

from sms_ui.audio_runtime import play_alert_runtime


class AudioRuntimeTests(unittest.TestCase):
    def test_play_alert_runtime_skips_when_disabled(self):
        result = play_alert_runtime(
            voice_enabled=False,
            tts_file="alert.wav",
            get_last_play_time=lambda: 0.0,
            set_last_play_time=lambda value: None,
            play_sound=lambda *_args: None,
            beep=lambda *_args: None,
            filename_flag=1,
            async_flag=2,
            beep_flag=4,
        )

        self.assertEqual(result, "disabled")

    def test_play_alert_runtime_skips_during_cooldown(self):
        result = play_alert_runtime(
            voice_enabled=True,
            tts_file="alert.wav",
            get_last_play_time=lambda: 9.0,
            set_last_play_time=lambda value: None,
            play_sound=lambda *_args: None,
            beep=lambda *_args: None,
            filename_flag=1,
            async_flag=2,
            beep_flag=4,
            monotonic=lambda: 10.0,
        )

        self.assertEqual(result, "cooldown")

    def test_play_alert_runtime_plays_existing_file(self):
        calls = []
        last_time = []

        result = play_alert_runtime(
            voice_enabled=True,
            tts_file="alert.wav",
            get_last_play_time=lambda: 0.0,
            set_last_play_time=last_time.append,
            play_sound=lambda *args: calls.append(("play", args)),
            beep=lambda *args: calls.append(("beep", args)),
            filename_flag=1,
            async_flag=2,
            beep_flag=4,
            monotonic=lambda: 10.0,
            path_exists=lambda path: True,
        )

        self.assertEqual(result, "played")
        self.assertEqual(last_time, [10.0])
        self.assertEqual(calls, [("play", ("alert.wav", 3))])

    def test_play_alert_runtime_beeps_when_file_missing(self):
        calls = []

        result = play_alert_runtime(
            voice_enabled=True,
            tts_file="alert.wav",
            get_last_play_time=lambda: 0.0,
            set_last_play_time=lambda value: None,
            play_sound=lambda *_args: calls.append("play"),
            beep=lambda *args: calls.append(("beep", args)),
            filename_flag=1,
            async_flag=2,
            beep_flag=4,
            monotonic=lambda: 10.0,
            path_exists=lambda path: False,
        )

        self.assertEqual(result, "beep")
        self.assertEqual(calls, [("beep", (4,))])

    def test_play_alert_runtime_fallback_beeps_on_play_error(self):
        calls = []

        result = play_alert_runtime(
            voice_enabled=True,
            tts_file="alert.wav",
            get_last_play_time=lambda: 0.0,
            set_last_play_time=lambda value: None,
            play_sound=lambda *_args: (_ for _ in ()).throw(RuntimeError("boom")),
            beep=lambda *args: calls.append(("beep", args)),
            filename_flag=1,
            async_flag=2,
            beep_flag=4,
            monotonic=lambda: 10.0,
            path_exists=lambda path: True,
        )

        self.assertEqual(result, "fallback_beep")
        self.assertEqual(calls, [("beep", (4,))])


if __name__ == "__main__":
    unittest.main()
