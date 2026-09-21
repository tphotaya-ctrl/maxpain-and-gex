"""One-time login helper: opens Chrome with a persistent profile so the CME
session is remembered. Log in yourself in the window, open the QuikStrike
Gamma view once, then close the window. Credentials are never seen by this script."""
from pathlib import Path

from playwright.sync_api import sync_playwright

PROFILE = Path(__file__).parent / ".chrome_profile"
URL = "https://www.cmegroup.com/tools-information/quikstrike/open-interest-heatmap.html"

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        str(PROFILE), channel="chrome", headless=False,
        args=["--disable-blink-features=AutomationControlled"], viewport=None)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto(URL, wait_until="domcontentloaded")
    print("Chrome opened. Log in, open the Gamma view, then close the window.", flush=True)
    try:
        ctx.wait_for_event("close", timeout=0)
    except Exception:
        pass
print("Profile saved to", PROFILE)
