import threading
import time

from sms_core.tts_runtime import (
    ensure_tts_worker_runtime,
    generate_alert_voice_runtime,
    tts_worker_loop,
)
from sms_ui.audio_runtime import play_alert_runtime


def tts_worker_namespace_runtime(namespace, *, worker_loop=tts_worker_loop):
    return worker_loop(
        namespace["TTS_STOP"],
        namespace["TTS_REQ_Q"],
        namespace["TTS_LOCK"],
        lambda: namespace["TTS_FILE"],
        lambda path: namespace["_set_tts_file"](path),
        namespace["TTS_DIR"],
        namespace["DEFAULT_VOICE_TEXT"],
        namespace["play_alert"],
        lambda exc: namespace["log_file_only"](f"TTS 生成失败，使用系统声音兜底：{exc}"),
        fallback_beep=lambda: namespace["winsound"].MessageBeep(namespace["winsound"].MB_ICONASTERISK),
    )


def set_tts_file_namespace_runtime(namespace, path):
    namespace.__setitem__("TTS_FILE", path)


def ensure_tts_worker_namespace_runtime(namespace, *, ensure_runtime=ensure_tts_worker_runtime):
    return ensure_runtime(
        get_thread=lambda: namespace["TTS_THREAD"],
        set_thread=lambda thread: namespace.__setitem__("TTS_THREAD", thread),
        stop_event=namespace["TTS_STOP"],
        worker_target=namespace["_tts_worker"],
        thread_factory=namespace.get("threading", threading).Thread,
        log_error=lambda exc: namespace["log_file_only"](f"TTS 线程启动失败：{exc}"),
    )


def generate_alert_voice_namespace_runtime(
    namespace,
    *,
    force=False,
    text=None,
    play_after=False,
    generate_runtime=generate_alert_voice_runtime,
):
    return generate_runtime(
        force=force,
        text=text,
        play_after=play_after,
        get_voice_text=lambda: namespace["VOICE_TEXT"],
        default_text=namespace["DEFAULT_VOICE_TEXT"],
        ensure_worker=namespace["ensure_tts_worker"],
        request_queue=namespace["TTS_REQ_Q"],
        log_queue_full=lambda: namespace["log_file_only"]("⚠️ TTS 请求队列已满，已丢弃一次生成请求"),
    )


def play_alert_namespace_runtime(namespace, *, force=False, play_runtime=play_alert_runtime):
    return play_runtime(
        force=force,
        voice_enabled=namespace["VOICE_ENABLED"],
        tts_file=namespace["TTS_FILE"],
        get_last_play_time=lambda: namespace["_last_play_time"],
        set_last_play_time=lambda value: namespace.__setitem__("_last_play_time", value),
        play_sound=namespace["winsound"].PlaySound,
        beep=namespace["winsound"].MessageBeep,
        filename_flag=namespace["winsound"].SND_FILENAME,
        async_flag=namespace["winsound"].SND_ASYNC,
        beep_flag=namespace["winsound"].MB_ICONASTERISK,
        monotonic=namespace.get("time", time).monotonic,
    )
