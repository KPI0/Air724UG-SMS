import tkinter as tk
from tkinter import ttk


class SerialDebugFinder:
    def __init__(self, parent, text_widget, center_window):
        self.parent = parent
        self.text = text_widget
        self.center_window = center_window
        self.win = None
        self.var = tk.StringVar(value="")
        self.last_index = "1.0"
        self.trace_id = None

        self.text.tag_config("find_hit", background="yellow")
        self.text.tag_config("find_cur", background="#ff9f1a")
        self.text.tag_raise("find_cur")

    @property
    def term(self) -> str:
        return self.var.get().strip()

    def clear(self):
        try:
            self.text.tag_remove("find_hit", "1.0", "end")
            self.text.tag_remove("find_cur", "1.0", "end")
        except Exception:
            pass

    def find_all(self, term: str):
        self.clear()
        if not term:
            return
        start = "1.0"
        while True:
            pos = self.text.search(term, start, stopindex="end", nocase=True)
            if not pos:
                break
            endpos = f"{pos}+{len(term)}c"
            self.text.tag_add("find_hit", pos, endpos)
            start = endpos

    def highlight_range(self, start_idx: str, end_idx: str):
        term = self.term
        if not term:
            return
        start = start_idx
        while True:
            pos = self.text.search(term, start, stopindex=end_idx, nocase=True)
            if not pos:
                break
            endpos = f"{pos}+{len(term)}c"
            self.text.tag_add("find_hit", pos, endpos)
            start = endpos

    def find_next(self, _event=None):
        term = self.term
        if not term:
            return "break"

        pos = self.text.search(term, self.last_index, stopindex="end", nocase=True)
        if not pos:
            pos = self.text.search(term, "1.0", stopindex="end", nocase=True)
            if not pos:
                return "break"

        endpos = f"{pos}+{len(term)}c"
        self.text.see(pos)
        self.text.mark_set("insert", endpos)
        self.text.tag_remove("find_cur", "1.0", "end")
        self.text.tag_add("find_cur", pos, endpos)
        self.text.tag_raise("find_cur", "find_hit")
        self.last_index = endpos
        return "break"

    def find_prev(self, _event=None):
        term = self.term
        if not term:
            return "break"

        try:
            ranges = self.text.tag_ranges("find_cur")
            cur_start = ranges[0] if ranges else self.text.index("insert")
        except Exception:
            cur_start = self.text.index("insert")

        start = self.text.index(f"{cur_start}-1c")
        pos = self.text.search(term, start, stopindex="1.0", nocase=True, backwards=True)
        if not pos:
            pos = self.text.search(term, "end-1c", stopindex="1.0", nocase=True, backwards=True)
            if not pos:
                return "break"

        endpos = f"{pos}+{len(term)}c"
        self.text.see(pos)
        self.text.mark_set("insert", endpos)
        self.text.tag_remove("find_cur", "1.0", "end")
        self.text.tag_add("find_cur", pos, endpos)
        self.text.tag_raise("find_cur", "find_hit")
        self.last_index = endpos
        return "break"

    def open(self):
        if self.win is not None and self.win.winfo_exists():
            self.win.deiconify()
            self.win.lift()
            return

        self.win = tk.Toplevel(self.parent)
        self.win.withdraw()
        alpha_hidden = False
        try:
            self.win.attributes("-alpha", 0.0)
            alpha_hidden = True
        except Exception:
            pass
        self.win.title("查找 (Ctrl+F)")
        self.win.resizable(False, False)
        self.win.transient(self.parent)

        frame = ttk.Frame(self.win, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="查找：").grid(row=0, column=0, sticky="w")
        entry = ttk.Entry(frame, textvariable=self.var, width=28)
        entry.grid(row=0, column=1, padx=(6, 6))
        ttk.Button(frame, text="上一个", command=self.find_prev).grid(row=0, column=2, padx=(0, 6))
        ttk.Button(frame, text="下一个", command=self.find_next).grid(row=0, column=3)

        def on_change(*_):
            self.last_index = "1.0"
            self.find_all(self.term)

        if self.trace_id is None:
            self.trace_id = self.var.trace_add("write", on_change)

        entry.bind("<Return>", self.find_next)
        entry.bind("<Shift-Return>", self.find_prev)
        entry.bind("<Escape>", self.close)
        self.win.bind("<Escape>", self.close)
        self.win.protocol("WM_DELETE_WINDOW", self.close)

        self.win.update_idletasks()
        self.center_window(self.win, self.parent)
        self.win.deiconify()
        if alpha_hidden:
            # Windows may apply its default placement when a Toplevel is mapped
            # for the first time.  Keep that first map invisible, then center
            # again using the mapped size before revealing the dialog.
            self.win.update_idletasks()
            self.center_window(self.win, self.parent)
            self.win.attributes("-alpha", 1.0)
        self.win.lift()
        entry.focus_set()

    def close(self, _event=None):
        self.clear()
        self.last_index = "1.0"

        if self.trace_id is not None:
            try:
                self.var.trace_remove("write", self.trace_id)
            except Exception:
                pass
            self.trace_id = None

        try:
            self.var.set("")
        except Exception:
            pass

        try:
            if self.win is not None and self.win.winfo_exists():
                self.win.destroy()
        except Exception:
            pass
        self.win = None

        try:
            self.text.focus_set()
        except Exception:
            pass
        return "break"
