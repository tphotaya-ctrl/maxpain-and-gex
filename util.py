"""Small shared helpers."""
import os
from pathlib import Path


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
