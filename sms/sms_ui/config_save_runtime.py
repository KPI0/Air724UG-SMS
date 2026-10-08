"""Keep synchronous settings transactions compatible without blocking Tk I/O."""
from contextlib import contextmanager
import queue
import threading
import time
import tkinter as tk

from sms_core.threading_runtime import start_daemon_thread
from sms_ui.shutdown_progress import ShutdownProgress


def run_config_transaction_for_ui(namespace, operation, *, queue_timeout=10.0):
    """Serialize background config commits with UI edits; keep file I/O off Tk."""
    def stopping():
        return bool(namespace.get("is_exiting")) or any(
            event is not None and event.is_set()
            for event in (namespace.get("TK_SHUTDOWN"), namespace.get("serial_stop_event"))
        )

    def allowed():
        return not stopping() and not namespace.get("_CONFIG_SAVE_ACTIVE")

    if not allowed():
        return False
    if threading.current_thread() is threading.main_thread():
        return operation()
    if not isinstance(namespace.get("root"), tk.Misc):
        return operation()

    done = threading.Event()
    state_lock = threading.Lock()
    started = cancelled = False
    result = False

    def commit():
        nonlocal started, operation, result
        with state_lock:
            if cancelled:
                return
            started = True
        try:
            if allowed():
                result = operation()
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if log_error:
                try:
                    log_error(f"配置事务执行失败：{type(exc).__name__}")
                except Exception:
                    pass
        finally:
            operation = None
            done.set()

    posted = namespace["ui_post"](commit)
    if posted is False:
        return False
    deadline = time.monotonic() + max(0.0, queue_timeout)
    while not done.wait(0.05):
        if stopping() or time.monotonic() >= deadline:
            with state_lock:
                if not started:
                    cancelled = True
                    operation = None
                    return False
            # An accepted write must settle even if shutdown starts meanwhile.
    return result


class UiConfigSave:
    def __init__(self, namespace):
        self.namespace = namespace
        self.root = namespace["root"]
        self.previous_grab = self.root.grab_current()
        self.parent = self.previous_grab or self.root
        self.disabled = []
        self.progress = None
        self.show_timer = None

    def _log(self, message):
        callback = self.namespace.get("log_file_only")
        if callback:
            try:
                callback(message)
            except Exception:
                pass

    def _disable_window(self, window):
        # Prevent another settings transaction during the nested Tk wait.
        try:
            previous = window.attributes("-disabled")
            window.attributes("-disabled", True)
            self.disabled.append((window, previous))
        except tk.TclError:
            pass
        for child in window.winfo_children():
            if isinstance(child, tk.Toplevel):
                self._disable_window(child)

    def __enter__(self):
        try:
            self._disable_window(self.root)
            self.show_timer = self.root.after(120, self._show_progress)
        except Exception:
            self.__exit__()
            raise
        return self

    def _show_progress(self):
        self.show_timer = None
        try:
            parent = self.parent if self.parent.winfo_exists() else self.root
            self.progress = ShutdownProgress(
                parent, title="正在保存设置", message="正在保存设置，请稍候…"
            )
        except tk.TclError as exc:
            self._log(f"显示配置保存进度失败：{exc}")

    def wait_for(self, ready):
        if ready():
            return True
        done = tk.BooleanVar(master=self.root, value=False)
        timer = None
        destroyed = False

        def on_destroy(event):
            nonlocal destroyed, timer
            if event.widget is self.root:
                destroyed = True
                # Windows can dispatch a due timer while destroying the root.
                # Cancel before Tk removes the registered callback command.
                if timer:
                    self.root.after_cancel(timer)
                    timer = None
                if self.show_timer:
                    self.root.after_cancel(self.show_timer)
                    self.show_timer = None
                done.set(True)

        def poll():
            nonlocal timer
            timer = None
            if ready():
                done.set(True)
            else:
                timer = self.root.after(20, poll)

        binding = self.root.bind("<Destroy>", on_destroy, add="+")
        try:
            timer = self.root.after(20, poll)
            self.root.wait_variable(done)
            return not destroyed
        finally:
            try:
                # Tcl timers survive destruction of the Tk root and must still
                # be cancelled once the nested wait unwinds.
                if timer:
                    self.root.after_cancel(timer)
                if not destroyed:
                    self.root.unbind("<Destroy>", binding)
            finally:
                # Break the recursive poll closure and release its Tk variable
                # here, before a worker's garbage collection can finalize it.
                poll = None
                done = None

    @contextmanager
    def lock(self, config_lock):
        if not self.wait_for(lambda: config_lock.acquire(blocking=False)):
            raise RuntimeError("窗口已关闭，未开始保存配置")
        try:
            yield
        finally:
            config_lock.release()

    def run_io(self, operation):
        result = queue.Queue(maxsize=1)

        def work():
            try:
                result.put((True, operation()))
            except BaseException as exc:
                result.put((False, exc))

        worker = start_daemon_thread(
            "config_save", work,
            before_start=lambda thread: self.namespace.__setitem__("_CONFIG_SAVE_THREAD", thread),
        )
        try:
            self.wait_for(lambda: not result.empty())
        finally:
            # Even if Tk is destroyed, a completed disk write must be accounted
            # for before a caller can roll back memory or a second save starts.
            worker.join()
        success, value = result.get_nowait()
        if not success:
            try:
                raise value
            finally:
                # The traceback includes this frame; do not retain the
                # exception and its Tk-owned waiter through that cycle.
                value = None
        return value

    def __exit__(self, *_args):
        try:
            if self.show_timer:
                self.root.after_cancel(self.show_timer)
            if self.progress:
                self.progress.close()
        except tk.TclError:
            pass
        for window, previous in self.disabled:
            try:
                window.attributes("-disabled", previous)
            except tk.TclError:
                pass
        try:
            if self.previous_grab and self.previous_grab.winfo_exists():
                self.previous_grab.grab_set()
        except tk.TclError:
            pass
        self.namespace.pop("_CONFIG_SAVE_THREAD", None)


def save_config_for_ui(namespace, save_config, options):
    root = namespace.get("root")
    if not isinstance(root, tk.Misc) or threading.current_thread() is not threading.main_thread():
        return save_config(**options)
    if namespace.get("_CONFIG_SAVE_ACTIVE"):
        return False
    namespace["_CONFIG_SAVE_ACTIVE"] = True
    try:
        with UiConfigSave(namespace) as waiter:
            return save_config(**{
                **options, "config_lock": waiter.lock(options["config_lock"]), "run_io": waiter.run_io,
            })
    except Exception as exc:
        log_error = options.get("log_error")
        if log_error:
            try:
                log_error(f"配置保存失败：{exc}")
            except Exception:
                pass
        return False
    finally:
        namespace["_CONFIG_SAVE_ACTIVE"] = False
        deferred = namespace.pop("_CONFIG_SAVE_DEFERRED_ACTION", None)
        if deferred:
            try:
                if root.winfo_exists():
                    root.after_idle(deferred)
            except tk.TclError:
                # The write may have settled after the root was destroyed.
                pass
