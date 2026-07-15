import tkinter as tk
from tkinter import messagebox, ttk


def open_keywords_setting_dialog(
    parent,
    keywords,
    log_unmatched,
    on_keywords_changed,
    on_log_unmatched_changed,
    center_window,
):
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("短信关键词设置")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    bottom_line = tk.Frame(win, height=1, bg="#d4d4d4")
    bottom_line.pack(side="bottom", fill="x")
    bottom_line.pack_propagate(False)

    frame = tk.Frame(win, padx=12, pady=10)
    frame.pack(side="top", fill=tk.BOTH, expand=True)

    tk.Label(frame, text="关键词列表：").grid(row=0, column=0, sticky="w")

    listbox = tk.Listbox(frame, height=8, width=20)
    listbox.grid(row=1, column=0, rowspan=4, sticky="nsew", pady=(6, 0))

    right = tk.Frame(frame)
    right.grid(row=1, column=1, sticky="n", padx=(12, 0), pady=(6, 0))

    tk.Label(right, text="关键词：").pack(anchor="w")
    entry_var = tk.StringVar()
    entry = tk.Entry(right, textvariable=entry_var, width=22)
    entry.pack(anchor="w", pady=(4, 10))

    def refresh_list(select_index=None):
        listbox.delete(0, tk.END)
        for keyword in keywords:
            listbox.insert(tk.END, keyword)
        if select_index is not None and 0 <= select_index < len(keywords):
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(select_index)
            listbox.see(select_index)

    def get_entry_value():
        return entry_var.get().strip()

    def on_select(_evt=None):
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        try:
            entry_var.set(keywords[idx])
        except Exception:
            pass

    def add_keyword():
        value = get_entry_value()
        if not value:
            messagebox.showerror("错误", "关键词不能为空", parent=win)
            return
        if value in keywords:
            messagebox.showwarning("提示", "该关键词已存在", parent=win)
            return

        keywords.append(value)
        if on_keywords_changed("add", value=value) is False:
            keywords.pop()
            refresh_list()
            return
        refresh_list(select_index=len(keywords) - 1)

    def delete_keyword():
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请选择要删除的关键词", parent=win)
            return

        idx = sel[0]
        if idx < 0 or idx >= len(keywords):
            return

        old_value = keywords.pop(idx)
        if on_keywords_changed("delete", value=old_value) is False:
            keywords.insert(idx, old_value)
            refresh_list(select_index=idx)
            return
        entry_var.set("")
        refresh_list(select_index=min(idx, len(keywords) - 1))

    def edit_keyword():
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请选择要修改的关键词", parent=win)
            return

        idx = sel[0]
        value = get_entry_value()
        if not value:
            messagebox.showerror("错误", "关键词不能为空", parent=win)
            return
        if value in keywords and keywords[idx] != value:
            messagebox.showwarning("提示", "该关键词已存在", parent=win)
            return

        old_value = keywords[idx]
        keywords[idx] = value
        if on_keywords_changed("edit", value=value, old_value=old_value) is False:
            keywords[idx] = old_value
            refresh_list(select_index=idx)
            return
        refresh_list(select_index=idx)

    tk.Button(right, text="增加", width=10, command=add_keyword).pack(anchor="w", pady=(0, 6))
    tk.Button(right, text="删除", width=10, command=delete_keyword).pack(anchor="w", pady=(0, 6))
    tk.Button(right, text="修改", width=10, command=edit_keyword).pack(anchor="w")

    listbox.bind("<<ListboxSelect>>", on_select)

    tip = tk.Label(
        frame,
        text="💡 提示：关键词为空时，不做过滤全部短信都会显示。",
        fg="gray",
        font=("微软雅黑", 9),
        anchor="w",
    )
    tip.grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 6))

    log_unmatched_var = tk.BooleanVar(value=log_unmatched)
    last_log_unmatched = bool(log_unmatched)

    def toggle_log_unmatched():
        nonlocal last_log_unmatched
        enabled = bool(log_unmatched_var.get())
        if on_log_unmatched_changed(enabled) is False:
            log_unmatched_var.set(last_log_unmatched)
            return
        last_log_unmatched = enabled

    chk_unmatched = ttk.Checkbutton(
        frame,
        text="将未匹配关键词的短信也写入到 sms_COM 日志文件中",
        variable=log_unmatched_var,
        command=toggle_log_unmatched,
    )
    chk_unmatched.grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 6))

    bottom = tk.Frame(frame)
    bottom.grid(row=7, column=0, columnspan=2, sticky="e", pady=(0, 10))
    tk.Button(bottom, text="关闭", width=10, command=win.destroy).pack()

    frame.grid_columnconfigure(0, weight=1)

    refresh_list()
    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    entry.focus_set()

    win.bind("<Return>", lambda _e: edit_keyword())
    listbox.bind("<Delete>", lambda _e: delete_keyword())
    win.bind("<Escape>", lambda _e: win.destroy())
