import queue
import threading


def drain_available_log_lines(log_queue, first_item):
    path, line = first_item
    batches = {path: [line]}
    while True:
        try:
            next_path, next_line = log_queue.get_nowait()
        except queue.Empty:
            break
        batches.setdefault(next_path, []).append(next_line)
    return batches


def write_log_batches(batches, encoding="utf-8", open_file=open):
    written = 0
    for path, lines in batches.items():
        try:
            with open_file(path, "a", encoding=encoding) as file:
                file.writelines(lines)
            written += len(lines)
        except Exception:
            pass
    return written


def run_file_log_worker(
    *,
    log_queue,
    stop_event,
    poll_timeout=0.5,
    drain_batches=drain_available_log_lines,
    write_batches=write_log_batches,
):
    while not stop_event.is_set():
        try:
            first_item = log_queue.get(timeout=poll_timeout)
        except queue.Empty:
            continue

        batches = drain_batches(log_queue, first_item)
        write_batches(batches)


def start_file_log_worker(
    *,
    log_queue,
    stop_event,
    thread_factory=threading.Thread,
):
    thread = thread_factory(
        target=lambda: run_file_log_worker(log_queue=log_queue, stop_event=stop_event),
        daemon=True,
    )
    thread.start()
    return thread
