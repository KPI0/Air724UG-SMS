import os
import queue
import uuid

from sms_core.threading_runtime import start_daemon_thread


def ensure_tts_worker_runtime(
    *,
    get_thread,
    set_thread,
    stop_event,
    worker_target,
    thread_factory,
    log_error,
):
    try:
        thread = get_thread()
        if thread is not None and thread.is_alive():
            return "already_running"
        if stop_event.is_set():
            return "stopped"
        thread = start_daemon_thread(
            "tts_worker",
            worker_target,
            log_error=log_error,
            before_start=set_thread,
            thread_factory=thread_factory,
        )
        return "started"
    except Exception as exc:
        log_error(exc)
        return "error"


def generate_alert_voice_runtime(
    *,
    force=False,
    text=None,
    play_after=False,
    get_voice_text,
    default_text,
    ensure_worker,
    request_queue,
    log_queue_full,
    normalize_text=None,
    enqueue_request=None,
):
    normalize_text = normalize_text or normalize_voice_text
    enqueue_request = enqueue_request or enqueue_tts_request
    try:
        text_snapshot = normalize_text(get_voice_text() if text is None else text, default_text)
    except Exception:
        text_snapshot = default_text

    try:
        ensure_worker()
        enqueue_request(request_queue, text_snapshot, force=force, play_after=play_after)
        return "queued"
    except queue.Full:
        log_queue_full()
        return "full"


def cleanup_tts_alt_files(tts_dir, current_file):
    """Remove stale alternate wav files left after PermissionError fallbacks."""
    removed = 0
    try:
        current = os.path.abspath(current_file)
        for name in os.listdir(tts_dir):
            if not (name.startswith("alert_alt_") and name.endswith(".wav")):
                continue
            path = os.path.abspath(os.path.join(tts_dir, name))
            if path == current:
                continue
            try:
                os.remove(path)
                removed += 1
            except Exception:
                pass
    except Exception:
        pass
    return removed


def normalize_voice_text(text, default_text):
    return (str(text or "").strip() or str(default_text or "").strip())


def clear_tts_queue(request_queue):
    cleared = 0
    while True:
        try:
            request_queue.get_nowait()
            request_queue.task_done()
            cleared += 1
        except queue.Empty:
            break
        except Exception:
            break
    return cleared


def enqueue_tts_request(request_queue, text, force=False, play_after=False, clear_existing=True):
    if clear_existing:
        clear_tts_queue(request_queue)
    request_queue.put_nowait((text, bool(force), bool(play_after)))
    return True


def _unpack_tts_task(task, default_text):
    if len(task) == 3:
        text, force, play_after = task
    else:
        text, force = task
        play_after = False
    return normalize_voice_text(text, default_text), bool(force), bool(play_after)


def generate_tts_file(
    text,
    tts_file,
    tts_dir,
    tts_lock,
    engine_factory,
    cleanup_func=cleanup_tts_alt_files,
    uuid_func=None,
):
    with tts_lock:
        target_dir = os.path.dirname(tts_file)
        if target_dir:
            os.makedirs(target_dir, exist_ok=True)
        cleanup_func(tts_dir, tts_file)
        tmp_path = tts_file + ".tmp.wav"

        engine = engine_factory()
        try:
            engine.setProperty("rate", 150)
            engine.save_to_file(text, tmp_path)
            engine.runAndWait()
        finally:
            # stop() must run even if runAndWait() raised, otherwise the SAPI
            # engine can be left in a busy state and leak COM resources.
            try:
                engine.stop()
            except Exception:
                pass
            del engine

        if not os.path.exists(tmp_path):
            return tts_file

        try:
            os.replace(tmp_path, tts_file)
            return tts_file
        except PermissionError:
            token = (uuid_func or (lambda: uuid.uuid4().hex[:8]))()
            fallback_file = os.path.join(tts_dir, f"alert_alt_{token}.wav")
            os.replace(tmp_path, fallback_file)
            return fallback_file


def tts_worker_loop(
    stop_event,
    request_queue,
    tts_lock,
    get_tts_file,
    set_tts_file,
    tts_dir,
    default_text,
    play_after_callback,
    error_callback,
    fallback_beep=None,
    engine_factory=None,
    poll_timeout=0.5,
):
    if engine_factory is None:
        import pyttsx3

        engine_factory = pyttsx3.init

    while not stop_event.is_set():
        try:
            task = request_queue.get(timeout=poll_timeout)
        except queue.Empty:
            continue

        text, force, play_after = _unpack_tts_task(task, default_text)
        current_file = get_tts_file()

        if (not force) and os.path.exists(current_file):
            try:
                request_queue.task_done()
            except Exception:
                pass
            if play_after:
                play_after_callback(force=True)
            continue

        try:
            new_file = generate_tts_file(
                text,
                current_file,
                tts_dir,
                tts_lock,
                engine_factory,
            )
            if new_file != current_file:
                set_tts_file(new_file)
        except Exception as exc:
            try:
                tmp_path = get_tts_file() + ".tmp.wav"
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
            error_callback(exc)
            if play_after and fallback_beep is not None:
                try:
                    fallback_beep()
                except Exception:
                    pass
                play_after = False
        finally:
            try:
                request_queue.task_done()
            except Exception:
                pass

        if play_after:
            play_after_callback(force=True)
