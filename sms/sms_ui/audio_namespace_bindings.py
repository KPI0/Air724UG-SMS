from sms_core.namespace_binding import make_namespace_runtime_binder
from sms_ui.audio_namespace_runtime import (
    ensure_tts_worker_namespace_runtime,
    generate_alert_voice_namespace_runtime,
    play_alert_namespace_runtime,
    set_tts_file_namespace_runtime,
    tts_worker_namespace_runtime,
)


def install_audio_namespace_bindings(namespace):
    bind = make_namespace_runtime_binder(namespace, globals())

    namespace.update({
        "_tts_worker": bind("tts_worker_namespace_runtime"),
        "_set_tts_file": bind("set_tts_file_namespace_runtime"),
        "ensure_tts_worker": bind("ensure_tts_worker_namespace_runtime"),
        "generate_alert_voice": bind(
            "generate_alert_voice_namespace_runtime",
            positional_keywords=("force", "text", "play_after"),
        ),
        "play_alert": bind(
            "play_alert_namespace_runtime",
            positional_keywords=("force",),
        ),
    })
    return namespace
