import re
import sys

from sms_app.version import APP_VERSION


STABLE_TAG_RE = re.compile(r"^v(\d+)\.(\d+)\.(\d+)$")


def validate_release_tag(tag, app_version=APP_VERSION):
    tag = str(tag or "").strip()
    app_version = str(app_version or "").strip()
    if not STABLE_TAG_RE.fullmatch(tag):
        return False, f"stable release tag must use vX.Y.Z format: {tag!r}"
    if not re.fullmatch(r"\d+\.\d+\.\d+", app_version):
        return False, f"APP_VERSION must use X.Y.Z format: {app_version!r}"
    expected = f"v{app_version}"
    if tag != expected:
        return False, f"release tag {tag!r} does not match client version {expected!r}"
    return True, expected


def main(argv=None):
    args = list(sys.argv[1:] if argv is None else argv)
    tag = args[0] if args else ""
    ok, message = validate_release_tag(tag)
    if not ok:
        print(message, file=sys.stderr)
        return 1
    print(message)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
