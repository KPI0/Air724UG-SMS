import os
import queue
import re
import uuid

from sms_core.threading_runtime import start_daemon_thread, task_done_safely


def instance_tts_file_path(tts_dir, instance_number=1):
    try:
        number = max(1, int(instance_number or 1))
    except (TypeError, ValueError):
        number = 1
    filename = "alert.wav" if number == 1 else f"alert_{number}.wav"
    return os.path.join(tts_dir, filename)


def ensure_tts_worker_runtime(
    *,
    get_thread,
    set_thread,
    stop_event,
    worker_target,
    thread_factory,
    log_error,
):
    published_thread = None

    def publish_thread(thread):
        nonlocal published_thread
        published_thread = thread
        set_thread(thread)

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
            before_start=publish_thread,
            thread_factory=thread_factory,
        )
        return "started"
    except Exception as exc:
        if published_thread is not None:
            try:
                if get_thread() is published_thread:
                    set_thread(None)
            except Exception:
                pass
        try:
            log_error(exc)
        except Exception:
            pass
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
        family = tts_file_family(current_file)
        alt_prefix = family + "_alt_"
        for name in os.listdir(tts_dir):
            if not (name.startswith(alt_prefix) and name.endswith(".wav")):
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


def tts_file_family(tts_file):
    name = os.path.basename(str(tts_file or ""))
    stem = name[:-4] if name.lower().endswith(".wav") else name
    if "_alt_" in stem:
        stem = stem.split("_alt_", 1)[0]
    if re.fullmatch(r"alert(?:_\d+)?", stem):
        return stem
    return "alert"


def normalize_voice_text(text, default_text):
    return (str(text or "").strip() or str(default_text or "").strip())


def clear_tts_queue(request_queue):
    cleared = 0
    while True:
        try:
            request_queue.get_nowait()
            task_done_safely(request_queue)
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


def _safe_tts_error(error_callback, exc):
    try:
        error_callback(exc)
    except Exception:
        pass


def _safe_play_after(play_after_callback, error_callback):
    try:
        play_after_callback(force=True)
        return True
    except Exception as exc:
        _safe_tts_error(error_callback, exc)
        return False


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
        token = (uuid_func or (lambda: uuid.uuid4().hex[:12]))()
        tmp_path = f"{tts_file}.{token}.tmp.wav"

        engine = engine_factory()
        try:
            engine.setProperty("rate", 150)
            engine.save_to_file(text, tmp_path)
            engine.runAndWait()
        except Exception:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
            raise
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
            fallback_file = os.path.join(
                tts_dir,
                f"{tts_file_family(tts_file)}_alt_{token}.wav",
            )
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

        try:
            if stop_event.is_set():
                break

            try:
                text, force, play_after = _unpack_tts_task(task, default_text)
                current_file = get_tts_file()
            except Exception as exc:
                if not stop_event.is_set():
                    _safe_tts_error(error_callback, exc)
                continue

            if (not force) and os.path.exists(current_file):
                if play_after:
                    _safe_play_after(play_after_callback, error_callback)
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
                if not stop_event.is_set():
                    _safe_tts_error(error_callback, exc)
                if not stop_event.is_set() and play_after and fallback_beep is not None:
                    try:
                        fallback_beep()
                    except Exception:
                        pass
                    play_after = False

            if play_after and not stop_event.is_set():
                _safe_play_after(play_after_callback, error_callback)
        except Exception as exc:
            # A malformed request or callback failure must not terminate the
            # only TTS worker; report it and continue with the next request.
            if not stop_event.is_set():
                _safe_tts_error(error_callback, exc)
        finally:
            task_done_safely(request_queue)
