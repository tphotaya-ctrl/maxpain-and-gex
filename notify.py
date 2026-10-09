"""Alerts: a desktop popup and, when configured, a Telegram message to the phone.

The popup is msg.exe, or a PowerShell WScript.Shell popup where msg.exe doesn't exist
(Windows Home - confirmed on the 26200 PC, 2026-10-01). Either pops a modal dialog in the interactive session - intrusive on purpose, these are
same-day "go look at this" alerts. Telegram is what reaches you when you're away from the PC.
Its bot token / chat id live in secrets.json (git-ignored - the repo is public), never in
config.json. Nothing here may break the caller: every failure falls back to printing, so
run.log still has the message.
"""
import json
import os
import subprocess
import sys
import urllib.parse
import urllib.request
from pathlib import Path

HERE = Path(__file__).parent
SECRETS = Path(os.environ.get("MAXPAIN_SECRETS", HERE / "secrets.json"))
# message goes through an env var, not the command line, so quotes/Thai text need no escaping
_PS_POPUP = "(New-Object -ComObject WScript.Shell).Popup($env:NOTIFY_TEXT, 0, $env:NOTIFY_TITLE, 48) | Out-Null"


def secrets():
    try:
        return json.loads(SECRETS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def send_telegram(text):
    """True if delivered; False when not configured or the request failed."""
    s = secrets()
    token, chat = s.get("telegram_token"), s.get("telegram_chat_id")
    if not token or not chat:
        return False
    data = urllib.parse.urlencode({"chat_id": chat, "text": text[:4000]}).encode()
    try:
        with urllib.request.urlopen(f"https://api.telegram.org/bot{token}/sendMessage", data, timeout=15) as r:
            return r.status == 200
    except Exception as e:
        print(f"notify: Telegram failed ({type(e).__name__}) - {text[:200]}")
        return False


def ping_health(fail=False, body=""):
    """healthchecks.io dead-man's switch: a missing ping is what triggers its alert."""
    url = secrets().get("healthcheck_url")
    if not url:
        return False
    try:
        with urllib.request.urlopen(url.rstrip("/") + ("/fail" if fail else ""), body.encode()[-10000:], timeout=15) as r:
            return r.status == 200
    except Exception as e:
        print(f"notify: healthcheck ping failed ({type(e).__name__})")
        return False


def notify(title, message):
    text = f"{title}: {message}"
    try:
        subprocess.run(["msg", os.environ.get("USERNAME", "*"), text],
                       timeout=10, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception as first:
        try:
            # Popen, not run: the popup blocks until dismissed and the daily run mustn't wait
            subprocess.Popen(["powershell", "-NoProfile", "-Command", _PS_POPUP],
                             env={**os.environ, "NOTIFY_TEXT": message, "NOTIFY_TITLE": title},
                             creationflags=subprocess.CREATE_NO_WINDOW)
        except Exception as e:
            print(f"notify: could not show popup ({type(first).__name__}: {first}; "
                  f"{type(e).__name__}: {e}) - {text}")
    send_telegram(text)


if __name__ == "__main__":
    notify(sys.argv[1] if len(sys.argv) > 1 else "MaxPain/GEX", sys.argv[2] if len(sys.argv) > 2 else "")
