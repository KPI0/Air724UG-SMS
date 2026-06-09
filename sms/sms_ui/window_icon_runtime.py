def install_window_icon_runtime(
    root,
    tk_module,
    messagebox_module,
    *,
    icon_path,
    path_exists,
    log_error,
):
    try:
        try:
            root.iconbitmap(icon_path)
        except Exception as exc:
            log_error(f"icon.ico 加载失败：{exc}")

        def apply_window_icon(window):
            try:
                if icon_path and path_exists(icon_path):
                    window.iconbitmap(icon_path)
            except Exception:
                pass

        original_toplevel = tk_module.Toplevel

        def patched_toplevel(*args, **kwargs):
            window = original_toplevel(*args, **kwargs)
            try:
                window.after(0, lambda w=window: apply_window_icon(w))
            except Exception:
                apply_window_icon(window)
            return window

        tk_module.Toplevel = patched_toplevel

        def wrap_messagebox(fn):
            def inner(title, message, **options):
                if "parent" not in options:
                    options["parent"] = root
                return fn(title, message, **options)
            return inner

        messagebox_module.showinfo = wrap_messagebox(messagebox_module.showinfo)
        messagebox_module.showwarning = wrap_messagebox(messagebox_module.showwarning)
        messagebox_module.showerror = wrap_messagebox(messagebox_module.showerror)
        messagebox_module.askyesno = wrap_messagebox(messagebox_module.askyesno)
        return True
    except Exception as exc:
        log_error(f"弹窗图标补丁加载失败：{exc}")
        return False
