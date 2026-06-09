from sms_ui.audio_namespace_runtime import (
    ensure_tts_worker_namespace_runtime,
    generate_alert_voice_namespace_runtime,
    play_alert_namespace_runtime,
    set_tts_file_namespace_runtime,
    tts_worker_namespace_runtime,
)


def install_audio_namespace_bindings(namespace):
    def tts_worker():
        return tts_worker_namespace_runtime(namespace)

    def set_tts_file(path):
        return set_tts_file_namespace_runtime(namespace, path)

    def ensure_tts_worker():
        return ensure_tts_worker_namespace_runtime(namespace)

    def generate_alert_voice(force=False, text=None, play_after=False):
        return generate_alert_voice_namespace_runtime(
            namespace,
            force=force,
            text=text,
            play_after=play_after,
        )

    def play_alert(force=False):
        return play_alert_namespace_runtime(namespace, force=force)

    namespace.update({
        "_tts_worker": tts_worker,
        "_set_tts_file": set_tts_file,
        "ensure_tts_worker": ensure_tts_worker,
        "generate_alert_voice": generate_alert_voice,
        "play_alert": play_alert,
    })
    return namespace
