import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.phone_numbers import is_valid_call_filter_number


def open_call_filter_setting_dialog(
    parent,
    mode,
    whitelist,
    blacklist,
    on_mode_changed,
    on_list_changed,
    center_window,
    *,
    register_external_refresh=None,
    get_mode=None,
):
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("来电防骚扰设置")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = tk.Frame(win, padx=15, pady=15)
    frm.pack(fill=tk.BOTH, expand=True)

    mode_frm = tk.LabelFrame(frm, text="过滤模式 (即时生效)", padx=10, pady=8)
    mode_frm.pack(fill="x", pady=(0, 15))

    mode_var = tk.StringVar(value=mode)
    current_mode = mode

    def change_mode():
        nonlocal current_mode
        next_mode = mode_var.get()
        if on_mode_changed(next_mode) is False:
            mode_var.set(current_mode)
            return
        current_mode = next_mode

    tk.Radiobutton(
        mode_frm,
        text="关闭过滤 (允许所有)",
        variable=mode_var,
        value="Disabled",
        command=change_mode,
    ).pack(side="left", padx=5)
    tk.Radiobutton(
        mode_frm,
        text="白名单 (仅限名单内)",
        variable=mode_var,
        value="Whitelist",
        command=change_mode,
    ).pack(side="left", padx=5)
    tk.Radiobutton(
        mode_frm,
        text="黑名单 (拦截名单内)",
        variable=mode_var,
        value="Blacklist",
        command=change_mode,
    ).pack(side="left", padx=5)

    notebook = ttk.Notebook(frm)
    notebook.pack(fill="both", expand=True)

    tab_white = tk.Frame(notebook, padx=10, pady=10)
    tab_black = tk.Frame(notebook, padx=10, pady=10)
    notebook.add(tab_white, text="白名单管理")
    notebook.add(tab_black, text="黑名单管理")

    def build_list_tab(parent_tab, target_list, list_kind, _list_name):
        listbox = tk.Listbox(parent_tab, height=10)
        listbox.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right_frm = tk.Frame(parent_tab)
        right_frm.pack(side="right", fill="y")

        tk.Label(right_frm, text="手机/电话号码：").pack(anchor="w")
        tk.Label(
            right_frm,
            text="💡 提示：需包含国际前缀\n请与模块日志上报完全一致\n(如: +8618888888...)。",
            fg="gray",
            justify="left",
            font=("微软雅黑", 8),
        ).pack(anchor="w", pady=(0, 5))

        entry_var = tk.StringVar()
        entry = ttk.Entry(right_frm, textvariable=entry_var, width=18)
        entry.pack(anchor="w", pady=(0, 10))

        def refresh(select_index=None):
            listbox.delete(0, tk.END)
            for number in target_list:
                listbox.insert(tk.END, number)
            if select_index is not None and 0 <= select_index < len(target_list):
                listbox.selection_set(select_index)
                listbox.see(select_index)

        def on_select(_e=None):
            sel = listbox.curselection()
            if sel:
                entry_var.set(target_list[sel[0]])

        def add_number():
            value = entry_var.get().strip()
            if not value:
                return
            if not is_valid_call_filter_number(value):
                messagebox.showerror(
                    "错误",
                    "号码格式无效，只允许可选的 + 前缀和数字",
                    parent=win,
                )
                return
            if value in target_list:
                messagebox.showwarning("提示", "该号码已在名单中", parent=win)
                return

            target_list.append(value)
            if on_list_changed(list_kind, action="add", value=value) is False:
                target_list.pop()
                refresh()
                return
            refresh()
            entry_var.set("")

        def delete_number():
            sel = listbox.curselection()
            if not sel:
                return

            idx = sel[0]
            value = target_list.pop(idx)
            if on_list_changed(list_kind, action="delete", value=value) is False:
                target_list.insert(idx, value)
                refresh(select_index=idx)
                return
            entry_var.set("")
            refresh(select_index=min(idx, len(target_list) - 1))

        def edit_number():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning(
                    "提示",
                    "请先在左侧列表中选择要修改的号码",
                    parent=win,
                )
                return

            idx = sel[0]
            old_value = target_list[idx]
            new_value = entry_var.get().strip()

            if not new_value:
                messagebox.showerror("错误", "号码不能为空", parent=win)
                return
            if not is_valid_call_filter_number(new_value):
                messagebox.showerror(
                    "错误",
                    "号码格式无效，只允许可选的 + 前缀和数字",
                    parent=win,
                )
                return
            if new_value == old_value:
                return
            if new_value in target_list:
                messagebox.showwarning("提示", "该号码已在名单中", parent=win)
                return

            target_list[idx] = new_value
            if on_list_changed(list_kind, action="edit", value=new_value, old_value=old_value) is False:
                target_list[idx] = old_value
                refresh(select_index=idx)
                return
            refresh(select_index=idx)

        listbox.bind("<<ListboxSelect>>", on_select)

        ttk.Button(right_frm, text="增加", command=add_number).pack(fill="x", pady=4)
        ttk.Button(right_frm, text="删除", command=delete_number).pack(fill="x", pady=4)
        ttk.Button(right_frm, text="修改", command=edit_number).pack(fill="x", pady=4)
        refresh()

        def refresh_external():
            entry_var.set("")
            refresh()

        return refresh_external

    refresh_whitelist = build_list_tab(tab_white, whitelist, "whitelist", "白名单")
    refresh_blacklist = build_list_tab(tab_black, blacklist, "blacklist", "黑名单")

    unregister_external_refresh = None

    def refresh_external_config():
        nonlocal current_mode
        current_mode = get_mode() if get_mode is not None else current_mode
        mode_var.set(current_mode)
        refresh_whitelist()
        refresh_blacklist()

    if register_external_refresh is not None:
        unregister_external_refresh = register_external_refresh(refresh_external_config)

    def unregister_on_destroy(event):
        if event.widget is not win or unregister_external_refresh is None:
            return
        unregister_external_refresh()

    win.bind("<Destroy>", unregister_on_destroy, add="+")

    tk.Button(frm, text="关闭窗口", width=12, command=win.destroy).pack(anchor="e", pady=(12, 0))

    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    win.bind("<Escape>", lambda _e: win.destroy())
