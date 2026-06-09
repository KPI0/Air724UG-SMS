def third_push_result_parent(root, current_window):
    try:
        if current_window is not None and current_window.winfo_exists():
            return current_window
    except Exception:
        pass
    return root


def third_push_result_dialog_plan(ok_channels, fail_infos):
    ok_channels = list(ok_channels or [])
    fail_infos = list(fail_infos or [])

    if ok_channels and fail_infos:
        return (
            "warning",
            "测试部分成功",
            "部分通道推送成功：\n" + "、".join(ok_channels) + "\n\n"
            "以下通道推送失败：\n" + "\n".join(fail_infos),
        )
    if fail_infos:
        return (
            "error",
            "测试推送失败",
            "三方推送测试失败：\n" + "\n".join(fail_infos),
        )
    if ok_channels:
        return (
            "info",
            "测试推送成功",
            "三方推送测试成功：\n" + "、".join(ok_channels),
        )
    return ("warning", "测试推送失败", "没有可用的通知通道。")


def show_third_push_test_result_runtime(
    *,
    root,
    current_window,
    messagebox,
    ui_post,
    ok_channels,
    fail_infos,
    parent_resolver=third_push_result_parent,
    dialog_planner=third_push_result_dialog_plan,
):
    def show_dialog():
        try:
            parent = parent_resolver(root, current_window)
            kind, title, message = dialog_planner(ok_channels, fail_infos)
            if kind == "error":
                messagebox.showerror(title, message, parent=parent)
            elif kind == "info":
                messagebox.showinfo(title, message, parent=parent)
            else:
                messagebox.showwarning(title, message, parent=parent)
        except Exception:
            pass

    ui_post(show_dialog)
