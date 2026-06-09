import os
import time


def play_alert_runtime(
    *,
    force=False,
    voice_enabled,
    tts_file,
    get_last_play_time,
    set_last_play_time,
    play_sound,
    beep,
    filename_flag,
    async_flag,
    beep_flag,
    cooldown_seconds=3.0,
    monotonic=time.monotonic,
    path_exists=os.path.exists,
):
    if (not force) and (not voice_enabled):
        return "disabled"

    now = monotonic()
    if (not force) and (now - get_last_play_time() < cooldown_seconds):
        return "cooldown"
    set_last_play_time(now)

    try:
        if path_exists(tts_file):
            play_sound(tts_file, filename_flag | async_flag)
            return "played"
        beep(beep_flag)
        return "beep"
    except Exception:
        beep(beep_flag)
        return "fallback_beep"
