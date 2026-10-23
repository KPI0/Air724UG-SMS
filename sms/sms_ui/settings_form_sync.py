import tkinter as tk


def bind_settings_form_sync(
    win, fields, *, get_values=None, register_external_refresh=None,
):
    """Refresh untouched fields and release the subscription with the dialog."""
    baseline = {name: read() for name, (read, _write) in fields.items()}

    def refresh(*, committed=False, committed_fields=()):
        try:
            if not win.winfo_exists():
                return
        except tk.TclError:
            return
        values = get_values() if get_values is not None else {
            name: read() for name, (read, _write) in fields.items()
        }
        for name, (read, write) in fields.items():
            if committed or name in committed_fields or read() == baseline[name]:
                value = str(values[name])
                if read() != value:
                    write(value)
                baseline[name] = read()

    unregister = None
    if register_external_refresh is not None and get_values is not None:
        unregister = register_external_refresh(refresh)

    def release(event):
        nonlocal unregister
        if event.widget is win and unregister is not None:
            unregister()
            unregister = None

    win.bind("<Destroy>", release, add="+")
    return lambda: refresh(committed=True)
