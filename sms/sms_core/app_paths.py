import os
import sys


def get_app_dir():
    """Return the directory where the script or frozen executable lives."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    try:
        package_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if os.path.basename(package_dir).lower() == "sms":
            return os.path.dirname(package_dir)
        return package_dir
    except Exception:
        return os.path.dirname(os.path.abspath(sys.argv[0] or "."))


def get_startup_dir():
    appdata = os.environ.get("APPDATA", "")
    return os.path.join(appdata, r"Microsoft\Windows\Start Menu\Programs\Startup")


def get_startup_lnk(name="sms.lnk"):
    return os.path.join(get_startup_dir(), name)


def resource_path(relative):
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, relative)
    app_path = os.path.join(get_app_dir(), relative)
    if os.path.exists(app_path):
        return app_path
    return os.path.join(os.path.dirname(get_app_dir()), relative)
