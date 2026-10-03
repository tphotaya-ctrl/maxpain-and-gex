"""Bookkeeping around run_daily.bat (date handling in Python, not locale-dependent batch):

    python run_state.py check                 exit 1 if today already ran OK/WARN, else 0
    python run_state.py finish STATUS LOGFILE  STATUS = OK | WARN | FAIL

The task fires at 08:30 and at logon (and late, when the PC was off at 08:30), so `check`
keeps it to one real run a day. `finish` records the day, pings the healthchecks.io
dead-man's switch (its alert is what tells you when the PC hasn't run at all), and sends a
Telegram message on FAIL with the end of the log.
"""
import sys
from datetime import date
from pathlib import Path

from notify import notify, ping_health

HERE = Path(__file__).parent
STAMP = HERE / "last_ok.txt"


def already_ran(today=None, stamp=STAMP):
    try:
        return stamp.read_text().strip() == str(today or date.today())
    except OSError:
        return False


def finish(status, log_text, today=None, stamp=STAMP):
    if status in ("OK", "WARN"):
        stamp.write_text(str(today or date.today()))
    ping_health(fail=status == "FAIL", body=log_text)
    if status == "FAIL":
        tail = "\n".join(log_text.strip().splitlines()[-8:])
        notify("MaxPain/GEX", f"daily run FAILED - see run.log\n{tail}")


def main(argv):
    if argv[1] == "check":
        return 1 if already_ran() else 0
    if argv[1] == "finish":
        try:
            text = Path(argv[3]).read_text(encoding="utf-8", errors="replace")
        except OSError:
            text = ""
        finish(argv[2], text)
        return 0
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
