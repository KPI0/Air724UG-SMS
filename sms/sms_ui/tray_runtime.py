import time

import pystray
from PIL import Image


def load_tray_image(icon_path, image_module=Image, fallback_size=(16, 16), fallback_color=(200, 30, 30)):
    try:
        return image_module.open(icon_path)
    except Exception:
        pass

    try:
        return image_module.new("RGB", fallback_size, color=fallback_color)
    except Exception:
        return None


def create_tray_icon_runtime(
    *,
    icon_path,
    title,
    show_window,
    hide_window,
    cleanup_and_exit,
    set_tray_icon,
    tray_name="sms_tray",
    pystray_module=pystray,
    image_loader=load_tray_image,
):
    image = image_loader(icon_path)
    if image is None:
        return None

    menu = pystray_module.Menu(
        pystray_module.MenuItem("显示", lambda: show_window(), default=True),
        pystray_module.MenuItem("隐藏", lambda: hide_window()),
        pystray_module.MenuItem("退出", lambda: cleanup_and_exit()),
    )
    icon = pystray_module.Icon(tray_name, image, title, menu)
    set_tray_icon(icon)
    icon.run()
    return icon


def stop_tray_icon_runtime(*, tray_icon, clear_tray_icon, wait_after=0.45, sleep=time.sleep):
    icon = tray_icon
    clear_tray_icon()
    if icon is None:
        return False

    try:
        icon.visible = False
    except Exception:
        pass
    try:
        icon.stop()
    except Exception:
        pass

    try:
        if wait_after:
            sleep(wait_after)
    except Exception:
        pass
    return True
