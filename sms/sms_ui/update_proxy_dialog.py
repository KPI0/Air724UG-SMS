import tkinter as tk
from tkinter import messagebox


DEFAULT_API_PROXY_BASE = "https://github-api.daybyday.top/"
DEFAULT_PROXY_BASE = "https://gh-proxy.com/"


def open_update_proxy_dialog(
    parent,
    api_proxy_base,
    proxy_base,
    on_save,
    on_test,
    center_window,
):
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("检查更新代理设置")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(fill=tk.BOTH, expand=True)

    proxy_var = tk.StringVar(value=proxy_base)
    api_var = tk.StringVar(value=api_proxy_base)

    tk.Label(frame, text="API 代理前缀 api_proxy_base：").grid(row=0, column=0, sticky="w")
    api_entry = tk.Entry(frame, textvariable=api_var, width=44)
    api_entry.grid(row=1, column=0, pady=(4, 10), sticky="w")

    tk.Label(frame, text="下载代理前缀 proxy_base：").grid(row=2, column=0, sticky="w")
    tk.Entry(frame, textvariable=proxy_var, width=44).grid(row=3, column=0, pady=(4, 10), sticky="w")

    def save():
        try:
            result = on_save(api_var.get(), proxy_var.get(), win)
            if result is False:
                raise RuntimeError("配置保存失败")
            messagebox.showinfo("完成", "代理设置已保存", parent=win)
        except Exception as exc:
            messagebox.showerror("保存失败", f"代理设置保存失败：{exc}", parent=win)

    def test_connection():
        try:
            btn_test.config(state="disabled", text="测试中…")
        except Exception:
            pass
        on_test(
            api_var.get(),
            proxy_var.get(),
            lambda text: _show_test_result(win, btn_test, text),
            lambda text: _show_test_error(win, btn_test, text),
        )

    def reset_default():
        api_var.set(DEFAULT_API_PROXY_BASE)
        proxy_var.set(DEFAULT_PROXY_BASE)

    btns = tk.Frame(frame)
    btns.grid(row=4, column=0, sticky="e", pady=(6, 0))
    btn_test = tk.Button(btns, text="测试连接", width=10, command=test_connection)
    btn_test.pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="恢复默认", width=10, command=reset_default).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="保存", width=10, command=save).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="取消", width=10, command=win.destroy).pack(side=tk.LEFT)

    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    api_entry.focus_set()
    win.bind("<Return>", lambda _e: save())
    win.bind("<Escape>", lambda _e: win.destroy())


def _widget_exists(widget):
    try:
        return bool(widget.winfo_exists())
    except Exception:
        return False


def _restore_test_button(btn_test):
    if not _widget_exists(btn_test):
        return False
    try:
        btn_test.config(state="normal", text="测试连接")
        return True
    except Exception:
        return False


def _show_test_result(win, btn_test, text):
    if not _widget_exists(win) or not _restore_test_button(btn_test):
        return False
    try:
        messagebox.showinfo("测试结果", text, parent=win)
        return True
    except Exception:
        return False


def _show_test_error(win, btn_test, text):
    if not _widget_exists(win) or not _restore_test_button(btn_test):
        return False
    try:
        messagebox.showerror("测试失败", text, parent=win)
        return True
    except Exception:
        return False
