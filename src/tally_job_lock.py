"""Serialize shared stock export files across web and WhatsApp processes."""
import functools
import os
import threading
import time
from .settings import PROJECT_ROOT

_lock = threading.RLock()
_local = threading.local()


def serialized(function):
    @functools.wraps(function)
    def run(*args, **kwargs):
        with _lock:
            if getattr(_local, 'active', False):
                return function(*args, **kwargs)
            path = PROJECT_ROOT / 'data' / 'tally_export.lock'
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open('a+b') as handle:
                handle.seek(0, 2)
                if handle.tell() == 0:
                    handle.write(b'0'); handle.flush()
                deadline = time.monotonic()+1800
                while True:
                    try:
                        handle.seek(0)
                        if os.name == 'nt':
                            import msvcrt
                            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                        else:
                            import fcntl
                            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                        break
                    except OSError:
                        if time.monotonic() > deadline:
                            raise TimeoutError('Another Tally export is still running. Retry later.')
                        time.sleep(1)
                _local.active = True
                try:
                    return function(*args, **kwargs)
                finally:
                    _local.active = False
                    handle.seek(0)
                    if os.name == 'nt':
                        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
                    else:
                        fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
    return run
