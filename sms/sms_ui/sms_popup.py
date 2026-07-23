import tkinter as tk
import unicodedata
from tkinter import ttk

from PIL import Image, ImageDraw, ImageTk


MIN_VISIBLE_MESSAGE_LINES = 2
MAX_VISIBLE_MESSAGE_LINES = 10


def estimate_message_lines(message, units_per_line=42):
    text = str(message or "")
    source_lines = text.splitlines() or [""]
    total_lines = 0
    for source_line in source_lines:
        units = sum(
            2 if unicodedata.east_asian_width(char) in ("W", "F", "A") else 1
            for char in source_line
        )
        total_lines += max(1, (units + units_per_line - 1) // units_per_line)
    return total_lines


def message_viewport(display_lines):
    try:
        lines = max(1, int(display_lines))
    except (TypeError, ValueError):
        lines = 1
    visible_lines = max(
        MIN_VISIBLE_MESSAGE_LINES,
        min(lines, MAX_VISIBLE_MESSAGE_LINES),
    )
    return visible_lines, lines > MAX_VISIBLE_MESSAGE_LINES


def create_information_icon(size=42, scale=4):
    """Create a smooth information icon by drawing large and downsampling."""
    render_size = int(size) * int(scale)
    image = Image.new("RGBA", (render_size, render_size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    margin = 2 * scale
    draw.ellipse(
        (margin, margin, render_size - margin - 1, render_size - margin - 1),
        fill="#087fd1",
    )

    center = render_size // 2
    dot_radius = 2 * scale
    draw.ellipse(
        (
            center - dot_radius,
            10 * scale - dot_radius,
            center + dot_radius,
            10 * scale + dot_radius,
        ),
        fill="white",
    )
    draw.rounded_rectangle(
        (
            center - 2 * scale,
            17 * scale,
            center + 2 * scale,
            32 * scale,
        ),
        radius=2 * scale,
        fill="white",
    )
    image = image.resize((size, size), Image.Resampling.LANCZOS)
    return ImageTk.PhotoImage(image)


def _pack_popup_sections(body, footer, close_button):
    """Keep the fixed footer visible while the message body resizes."""
    footer.pack(fill="x", side="bottom")
    footer.pack_propagate(False)
    close_button.pack(pady=13)
    body.pack(fill="both", expand=True)


def _pack_message_viewport(message_text, message_scrollbar):
    """Reserve the scrollbar before allowing the text area to expand."""
    message_scrollbar.pack(side="right", fill="y", padx=(8, 0))
    message_text.pack(side="left", fill="both", expand=True)


def open_sms_popup(parent, message, center_on_screen, on_close):
    """Open a non-modal SMS alert styled like a native information box."""
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("\u77ed\u4fe1\u63d0\u9192")
    win.minsize(536, 180)
    win.resizable(True, True)
    try:
        win.attributes("-topmost", True)
    except Exception:
        pass

    body = tk.Frame(win, bg="white", padx=26, pady=30)

    info_icon = create_information_icon()
    info = tk.Label(
        body,
        image=info_icon,
        bg="white",
        bd=0,
    )
    info.pack(side="left", anchor="n", padx=(0, 18))
    win.sms_popup_icon = info_icon

    message_frame = tk.Frame(body, bg="white")
    message_frame.pack(side="left", fill="both", expand=True)

    message_text = tk.Text(
        message_frame,
        width=42,
        height=MIN_VISIBLE_MESSAGE_LINES,
        wrap="char",
        bg="white",
        fg="#111111",
        relief="flat",
        bd=0,
        highlightthickness=0,
        padx=0,
        pady=0,
        cursor="arrow",
        takefocus=False,
        font=("Microsoft YaHei UI", 10),
    )

    message_scrollbar = ttk.Scrollbar(
        message_frame,
        orient="vertical",
        command=message_text.yview,
    )
    message_text.configure(yscrollcommand=message_scrollbar.set)
    _pack_message_viewport(message_text, message_scrollbar)
    resize_after_id = None
    last_text_width = None
    recenter_on_refresh = False

    def apply_viewport(lines):
        visible_lines, _is_capped = message_viewport(lines)
        message_text.configure(height=visible_lines)

    def refresh_actual_viewport():
        nonlocal resize_after_id, recenter_on_refresh
        resize_after_id = None
        should_recenter = recenter_on_refresh
        recenter_on_refresh = False
        try:
            counted = message_text.count("1.0", "end-1c", "displaylines")
            display_lines = int(counted[0]) if counted else 1
            apply_viewport(display_lines)
            message_text.yview_moveto(0.0)
            win.update_idletasks()
            if should_recenter:
                center_on_screen(win)
        except Exception:
            pass

    def schedule_viewport_refresh(event=None, *, recenter=False):
        nonlocal resize_after_id, last_text_width, recenter_on_refresh
        recenter_on_refresh = recenter_on_refresh or bool(recenter)
        if event is not None:
            width = int(getattr(event, "width", 0) or 0)
            if width <= 1 or width == last_text_width:
                return
            last_text_width = width
        if resize_after_id is not None:
            try:
                win.after_cancel(resize_after_id)
            except Exception:
                pass
        resize_after_id = win.after_idle(refresh_actual_viewport)

    def replace_message(next_message):
        message_text.configure(state="normal")
        message_text.delete("1.0", "end")
        message_text.insert("1.0", str(next_message or ""))
        message_text.configure(state="disabled")
        lines = estimate_message_lines(next_message)
        apply_viewport(lines)
        message_text.yview_moveto(0.0)

    message_text.bind("<Button-1>", lambda _event: "break")
    message_text.bind("<Key>", lambda _event: "break")
    message_text.bind("<Configure>", schedule_viewport_refresh, add="+")

    footer = tk.Frame(win, bg="#f0f0f0", height=58)
    close_button = ttk.Button(footer, text="\u786e\u5b9a", command=on_close, width=12)
    _pack_popup_sections(body, footer, close_button)

    def update_message(next_message):
        try:
            replace_message(next_message)
            win.update_idletasks()
            center_on_screen(win)
            win.deiconify()
            win.lift()
            win.focus_force()
            close_button.focus_set()
            schedule_viewport_refresh(recenter=True)
        except Exception:
            pass
        return win

    win.sms_popup_update = update_message
    win.protocol("WM_DELETE_WINDOW", on_close)
    win.bind("<Escape>", lambda _event: on_close())
    update_message(message)
    return win
