"""Best-effort desktop notification via msg.exe (built into Windows Pro/Enterprise,
confirmed present on this machine - no extra package to install).

msg.exe pops a modal dialog in the interactive session, which is intrusive, but that is
the point: these are same-day "go look at this" alerts (a stale CME login, a market
regime change), not routine status. A failure here must never break the caller - printing
to stdout is the fallback, so `run.log` still has the message even if the popup can't show
(e.g. no interactive session, as when testing under a non-interactive shell).
"""
import os
import subprocess
import sys


def notify(title, message):
    text = f"{title}: {message}"
    try:
        subprocess.run(["msg", os.environ.get("USERNAME", "*"), text],
                        timeout=10, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception as e:
        print(f"notify: could not show popup ({type(e).__name__}: {e}) - {text}")


if __name__ == "__main__":
    notify(sys.argv[1] if len(sys.argv) > 1 else "MaxPain/GEX", sys.argv[2] if len(sys.argv) > 2 else "")
