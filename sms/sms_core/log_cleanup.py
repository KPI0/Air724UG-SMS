import os
import re
from datetime import datetime, timedelta


def parse_date_from_log_filename(filename: str):
    """
    Parse dates from names like sms_system_YYYY-MM-DD.txt,
    sms_system_2_YYYY-MM-DD.txt, or sms_COM5_YYYY-MM-DD.txt.
    Returns None when the filename does not contain a supported suffix date.
    """
    match = re.search(r"_(\d{4}-\d{2}-\d{2})\.txt$", str(filename or ""))
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%Y-%m-%d").date()
    except Exception:
        return None


def cleanup_old_logs_in_dir(log_dir: str, days: int, now=None) -> int:
    """Delete sms_*.txt files older than the retention window and return the count."""
    try:
        days = int(days)
    except Exception:
        days = 0
    if days < 0:
        days = 0

    now_dt = now or datetime.now()
    cutoff = (now_dt - timedelta(days=days)).date()
    deleted = 0

    if not os.path.isdir(log_dir):
        return 0

    for name in os.listdir(log_dir):
        path = os.path.join(log_dir, name)
        if not os.path.isfile(path):
            continue
        if not name.lower().endswith(".txt"):
            continue
        if not name.lower().startswith("sms_"):
            continue

        file_date = parse_date_from_log_filename(name)
        try:
            if file_date is None:
                file_date = datetime.fromtimestamp(os.path.getmtime(path)).date()

            if file_date < cutoff:
                os.remove(path)
                deleted += 1
        except Exception:
            pass

    return deleted
