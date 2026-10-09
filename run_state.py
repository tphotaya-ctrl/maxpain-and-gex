"""Bookkeeping around run_daily.bat (date handling in Python, not locale-dependent batch):

    python run_state.py check                 exit 1 if today already ran OK/WARN, else 0
    python run_state.py finish STATUS LOGFILE  STATUS = OK | WARN | FAIL

The task fires at 09:00, 11:00, 14:00 and at logon (and late, when the PC was off), so
`check` keeps it to one real run a day - unless that run only got an older trade date than
the previous weekday (CME often hasn't published by 09:00), in which case the next trigger
runs again; run_daily.bat's --if-new then makes a run with nothing new a cheap SKIP. `finish` records the day, pings the healthchecks.io
dead-man's switch (its alert is what tells you when the PC hasn't run at all), and sends a
Telegram message on FAIL with the end of the log.
"""
import re
import sys
from datetime import date
from pathlib import Path

from notify import notify, ping_health

HERE = Path(__file__).parent
STAMP = HERE / "last_ok.txt"


def already_ran(today=None, stamp=STAMP):
    """Today already ran OK/WARN with fresh data. Stamp = 'YYYY-MM-DD[ trade-date]'; a stamp
    without a trade date (older format, or a run whose log had none) counts as fresh."""
    from telegram_report import expected_trade_date
    today = today or date.today()
    try:
        parts = stamp.read_text().split()
    except OSError:
        return False
    if not parts or parts[0] != str(today):
        return False
    return len(parts) < 2 or date.fromisoformat(parts[1]) >= expected_trade_date(today)


def logged_trade_date(log_text):
    """The trade date the run wrote or found already logged ('OK 2026-10-07 ...' / 'SKIP ...')."""
    m = re.search(r"^(?:OK|SKIP) (\d{4}-\d{2}-\d{2}) ", log_text, re.M)
    return m.group(1) if m else None


def finish(status, log_text, today=None, stamp=STAMP):
    if status in ("OK", "WARN"):
        td = logged_trade_date(log_text)
        stamp.write_text(str(today or date.today()) + (f" {td}" if td else ""))
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
