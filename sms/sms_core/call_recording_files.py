"""Cross-process ownership of serial recording temporary files."""

import os
import time
import uuid

if os.name == "nt":
    import msvcrt
else:
    import fcntl


# Old clients did not create a guard. Keep their recent partial files; modern
# transfers are protected by OS locks regardless of how old a file becomes.
LEGACY_PARTIAL_MAX_AGE_SECONDS = 24 * 60 * 60


def _lock_guard(file):
    file.seek(0)
    if os.name == "nt":
        msvcrt.locking(file.fileno(), msvcrt.LK_NBLCK, 1)
    else:
        fcntl.flock(file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)


def _same_guard(file, path):
    try:
        return os.path.samestat(os.fstat(file.fileno()), os.stat(path))
    except OSError:
        return False


def _close_guard(file, path, *, remove):
    # POSIX can unlink open files: remove while locked so a paused creator
    # cannot acquire an obsolete inode and then publish an unprotected part.
    # Windows prevents unlink while another ordinary file handle is open.
    if remove and os.name != "nt" and _same_guard(file, path):
        try:
            os.remove(path)
        except OSError:
            pass
    file.close()  # Closing releases the OS lock, including across threads.
    if remove and os.name == "nt":
        try:
            os.remove(path)
        except OSError:
            pass


def _open_locked_guard(path):
    if os.path.islink(path):
        return None
    try:
        file = open(path, "r+b")
    except OSError:
        return None
    try:
        _lock_guard(file)
        if not _same_guard(file, path):
            file.close()
            return None
        return file
    except OSError:
        file.close()
        return None


class IncomingRecordingLease:
    def __init__(self, path, guard):
        self.path = path
        self.guard = guard

    def close(self):
        guard = self.guard
        if guard is None:
            return
        self.guard = None
        # If deletion/commit failed, retain an unlocked guard so a later
        # startup can safely reclaim the orphan instead of guessing its age.
        _close_guard(guard, self.path + ".lock", remove=not os.path.exists(self.path))


def create_incoming_file(directory, recording_id):
    path = os.path.join(directory, recording_id + "." + uuid.uuid4().hex + ".part")
    guard_path = path + ".lock"
    guard = open(guard_path, "x+b")
    lease = IncomingRecordingLease(path, guard)
    try:
        guard.write(b"0")
        guard.flush()
        _lock_guard(guard)
        if not _same_guard(guard, guard_path):
            raise OSError("Recording file reservation was removed")
        # The guard is locked before the part becomes visible to cleanup.
        file = open(path, "xb")
        return path, file, lease
    except Exception:
        lease.close()
        raise


def cleanup_incoming_files(directory):
    removed = 0
    now = time.time()
    try:
        names = os.listdir(directory)
    except OSError:
        return 0

    def old_file(path):
        try:
            return now - os.stat(path).st_mtime >= LEGACY_PARTIAL_MAX_AGE_SECONDS
        except OSError:
            return False

    for name in names:
        path = os.path.join(directory, name)
        if os.path.islink(path) or not os.path.isfile(path):
            continue
        if name.endswith(".part"):
            guard_path = path + ".lock"
            if os.path.lexists(guard_path):
                guard = _open_locked_guard(guard_path)
                if guard is None:
                    continue
                try:
                    os.remove(path)
                    removed += 1
                except OSError:
                    pass
                finally:
                    _close_guard(guard, guard_path, remove=not os.path.exists(path))
            elif old_file(path):
                try:
                    os.remove(path)
                    removed += 1
                except OSError:
                    pass
        elif name.endswith(".part.lock") and not os.path.exists(path[:-5]) and old_file(path):
            # Leave newly created guards alone during the create/lock window.
            guard = _open_locked_guard(path)
            if guard is not None:
                _close_guard(guard, path, remove=not os.path.exists(path[:-5]))
    return removed
