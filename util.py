"""Small shared helpers."""
import os
import time
from pathlib import Path


def launch_persistent(p, user_data_dir, attempts=3, wait_s=10, **kw):
    """p.chromium.launch_persistent_context with retries.

    Chrome sometimes exits right at launch (TargetClosedError at the 08:30 run on
    2026-10-05 - the machine had just woken / Chrome was updating); a retry a few seconds
    later works. Persistent contexts only: Browser.close() hangs on this machine.
    """
    for i in range(attempts):
        try:
            return p.chromium.launch_persistent_context(str(user_data_dir), **kw)
        except Exception as e:
            if i == attempts - 1:
                raise
            print(f"chrome launch failed ({type(e).__name__}) - retrying in {wait_s}s")
            time.sleep(wait_s)


def save_atomic(wb, path):
    """Save to a temp file next to `path`, then replace it.

    A crash or power cut mid-save leaves the old workbook intact, and if Excel has the
    target open the replace raises PermissionError with the old file untouched.
    """
    path = Path(path)
    tmp = path.with_name(f".tmp_{path.name}")
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        if tmp.exists():
            tmp.unlink()
