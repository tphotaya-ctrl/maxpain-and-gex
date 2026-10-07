"""Best-effort desktop notification.

Tries msg.exe first (built into Windows Pro/Enterprise), then a PowerShell
WScript.Shell popup, which every Windows edition has - msg.exe is missing on Windows Home
(confirmed on the current machine, 2026-10-01).

The popup is intrusive, but that is the point: these are same-day "go look at this" alerts
(a stale CME login, a market regime change), not routine status. A failure here must never
break the caller - printing to stdout is the last fallback, so `run.log` still has the
message even if no popup can show (e.g. no interactive session).
"""
import os
import subprocess
import sys

# message goes through an env var, not the command line, so quotes/Thai text need no escaping
_PS_POPUP = "(New-Object -ComObject WScript.Shell).Popup($env:NOTIFY_TEXT, 0, $env:NOTIFY_TITLE, 48) | Out-Null"


def notify(title, message):
    text = f"{title}: {message}"
    try:  # the phone copy - a no-op until telegram.json exists (telegram_report.py --setup)
        from telegram_report import send_text
        send_text(f"⚠️ {text}")
    except Exception as e:
        print(f"notify: telegram failed ({type(e).__name__}: {e})")
    try:
        subprocess.run(["msg", os.environ.get("USERNAME", "*"), text],
                       timeout=10, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        return
    except Exception as e:
        first = e
    try:
        # Popen, not run: the popup blocks until dismissed and the daily run mustn't wait for it
        subprocess.Popen(["powershell", "-NoProfile", "-Command", _PS_POPUP],
                         env={**os.environ, "NOTIFY_TEXT": message, "NOTIFY_TITLE": title},
                         creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception as e:
        print(f"notify: could not show popup ({type(first).__name__}: {first}; "
              f"{type(e).__name__}: {e}) - {text}")


if __name__ == "__main__":
    notify(sys.argv[1] if len(sys.argv) > 1 else "MaxPain/GEX", sys.argv[2] if len(sys.argv) > 2 else "")
