"""Can CME's data be fetched from *this* machine? Run it on a candidate cloud host, or via
the manual GitHub Actions workflow 'cloud-probe', before moving the daily job off the PC.

Public CME endpoints only - no login, no secrets, nothing is written. One line per
(mode, endpoint), then a verdict. QuikStrike isn't probed: it needs the CME login session,
which this deliberately doesn't touch.

    python cloud_probe.py               # Windows / a machine with a desktop
    xvfb-run -a python cloud_probe.py   # Linux server: gives the headed mode a display
"""
import json
import os
import sys
import urllib.request

from playwright.sync_api import sync_playwright

from fetch_cme import PAGE

BASE = "https://www.cmegroup.com"
PRODUCT = 437
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/126.0 Safari/537.36")


def probe(get):
    """[(endpoint, ok, detail)] for the three calls the daily job depends on."""
    out = []
    try:
        dates = get("/CmeWS/mvc/Volume/TradeDates?exchange=CBOT&isProtected")
        td = dates[0]["tradeDate"]
        out.append(("TradeDates", True, f"latest {td}"))
    except Exception as e:
        return [("TradeDates", False, f"{type(e).__name__}: {str(e)[:120]}")]
    for name, url in (
        ("Options expirations",
         f"/CmeWS/mvc/Volume/Options/Expirations?productid={PRODUCT}&tradedate={td}&isProtected"),
        ("Futures settlements",
         f"/CmeWS/mvc/Settlements/Futures/Settlements/{PRODUCT}/FUT?strategy=DEFAULT"
         f"&tradeDate={td[4:6]}%2F{td[6:]}%2F{td[:4]}&pageSize=50&isProtected"),
    ):
        try:
            d = get(url)
            out.append((name, True, f"{len(d) if isinstance(d, list) else len(d.get('settlements', []))} item(s)"))
        except Exception as e:
            out.append((name, False, f"{type(e).__name__}: {str(e)[:120]}"))
    return out


def plain_http():
    def get(path):
        req = urllib.request.Request(BASE + path, headers={"User-Agent": UA, "Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read())
    return probe(get)


def browser(headless, channel=None):
    with sync_playwright() as p:
        b = p.chromium.launch(headless=headless, channel=channel,
                              args=["--disable-blink-features=AutomationControlled"])
        try:
            pg = b.new_page()
            pg.goto(PAGE, timeout=60000, wait_until="domcontentloaded")
            pg.wait_for_timeout(6000)

            def get(u):
                status, text = pg.evaluate("u=>fetch(u).then(async r=>[r.status,await r.text()])", u)
                if status != 200:
                    raise RuntimeError(f"HTTP {status}: {text[:80]}")
                return json.loads(text)
            return probe(get)
        finally:
            b.close()


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    has_display = os.name == "nt" or bool(os.environ.get("DISPLAY"))
    modes = [("plain HTTP", plain_http), ("headless Chromium", lambda: browser(True))]
    if has_display:
        modes.append(("headed Chrome (what the daily job uses)", lambda: browser(False, "chrome")))
    else:
        print("note: no display - headed mode skipped (run under xvfb-run to include it)")
    working = []
    for label, run in modes:
        try:
            results = run()
        except Exception as e:
            results = [("browser start/page load", False, f"{type(e).__name__}: {str(e)[:120]}")]
        ok = all(r[1] for r in results) and len(results) == 3
        print(f"\n## {label}: {'WORKS' if ok else 'BLOCKED/FAILED'}")
        for name, good, detail in results:
            print(f"- [{'ok' if good else 'FAIL'}] {name} - {detail}")
        if ok:
            working.append(label)
    print("\nVerdict: " + (f"CME data reachable here via {', '.join(working)}." if working else
                          "CME blocks every mode from this machine - a move here would not work."))
    print("(QuikStrike/GEX not probed: it needs the CME login session.)")


if __name__ == "__main__":
    main()
