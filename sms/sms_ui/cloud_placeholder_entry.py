import secrets
import string
import tkinter as tk
from tkinter import ttk


PLACEHOLDER_STYLE = "CloudPlaceholder.TEntry"


def configure_placeholder_style(master):
    try:
        style = ttk.Style(master)
        style.configure(PLACEHOLDER_STYLE, foreground="#777777")
    except Exception:
        pass


class CloudPlaceholderEntry:
    def __init__(
        self,
        parent,
        frame,
        row,
        label_text,
        variable,
        placeholder,
        visible=False,
        help_command=None,
        random_secret=False,
    ):
        self.parent = parent
        self.variable = variable
        self.placeholder = placeholder
        self.placeholder_active = False
        self.visible_var = tk.BooleanVar(parent, value=bool(visible))
        self.random_secret = bool(random_secret)

        ttk.Label(frame, text=label_text).grid(row=row, column=0, sticky="w", pady=(0, 8))
        self.entry = ttk.Entry(
            frame,
            textvariable=variable,
            show="" if self.visible_var.get() else "*",
        )
        self.entry.grid(row=row, column=1, sticky="ew", pady=(0, 8))
        self.entry.bind("<FocusIn>", self.clear_placeholder)
        self.entry.bind("<FocusOut>", self.restore_placeholder)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=2, sticky="e", padx=(8, 0), pady=(0, 8))

        if help_command is not None:
            ttk.Button(button_frame, text="?", width=3, command=help_command).pack(side="left", padx=(0, 4))
        if random_secret:
            ttk.Button(button_frame, text="🎲", width=3, command=self.generate_secret).pack(side="left", padx=(0, 4))

        self.eye_button = ttk.Button(button_frame, text=self._eye_text(), width=3, command=self.toggle_visible)
        self.eye_button.pack(side="left")
        self.set_placeholder()

    def _eye_text(self):
        return "🙈" if self.visible_var.get() else "👁"

    def _entry_show(self):
        return "" if self.visible_var.get() or self.placeholder_active else "*"

    def get(self):
        value = self.variable.get().strip()
        if self.placeholder_active and value == self.placeholder:
            return ""
        return value

    def set_value(self, value):
        self.placeholder_active = False
        self.variable.set(str(value or ""))
        try:
            self.entry.config(style="TEntry", show=self._entry_show())
        except Exception:
            pass

    def set_placeholder(self):
        if self.variable.get().strip():
            return
        self.placeholder_active = True
        self.variable.set(self.placeholder)
        try:
            self.entry.config(style=PLACEHOLDER_STYLE, show="")
        except Exception:
            pass

    def clear_placeholder(self, _event=None):
        if not self.placeholder_active:
            return
        self.placeholder_active = False
        self.variable.set("")
        try:
            self.entry.config(style="TEntry", show=self._entry_show())
        except Exception:
            pass

    def restore_placeholder(self, _event=None):
        if self.variable.get().strip():
            try:
                self.entry.config(style="TEntry", show=self._entry_show())
            except Exception:
                pass
            return
        self.set_placeholder()

    def toggle_visible(self):
        self.visible_var.set(not self.visible_var.get())
        try:
            self.entry.config(show=self._entry_show())
            self.eye_button.config(text=self._eye_text())
        except Exception:
            pass

    def generate_secret(self):
        chars = string.ascii_letters + string.digits
        self.placeholder_active = False
        self.variable.set("".join(secrets.choice(chars) for _ in range(16)))
        self.visible_var.set(True)
        try:
            self.entry.config(style="TEntry", show="")
            self.eye_button.config(text=self._eye_text())
        except Exception:
            pass
