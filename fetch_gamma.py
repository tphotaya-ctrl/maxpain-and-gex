"""Fetch per-strike Call/Put Gamma (1 Pct) from CME QuikStrike (OI Matrix view).

QuikStrike needs a CME login, so this reuses the persistent Chrome profile
created by quikstrike_login.py. The matrix for a single expiration lists one
C/P column pair per trade date, so any recent day can be read (backfill too).
"""
import json
from datetime import date
from pathlib import Path

from playwright.sync_api import sync_playwright

PAGE = "https://www.cmegroup.com/tools-information/quikstrike/open-interest-heatmap.html"
PROFILE = Path(__file__).parent / ".chrome_profile"
MONTH_CODE = "FGHJKMNQUVXZ"
WEEKDAY_LETTER = {"MW1": "M", "AB1": "T", "WD1": "W", "BB1": "R"}
GREEK = "#MainContent_ucViewControl_IntegratedVOIHeatMap_ucMatrixTB_ddlGreek"
STRIKES = "#MainContent_ucViewControl_IntegratedVOIHeatMap_ucMatrixTB_ddlStrikes"
MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


class LoginRequired(RuntimeError):
    pass


def qs_code(family, label):
    """('MW1', 'Week 3 - SEP 2026') -> 'G3MU6'; Friday weeklies are 'OG4U6'."""
    week, mon, year = label[5], label[9:12], label[16]
    mc = MONTH_CODE[MONTHS.index(mon)]
    if family == "E21":
        return f"OG{week}{mc}{year}"
    return f"G{week}{WEEKDAY_LETTER[family]}{mc}{year}"


def _num(s):
    s = s.replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def fetch_gamma(cfg, code, trade_date: date):
    """Return {strike: (call_gamma, put_gamma)} for `code` on `trade_date`."""
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(PROFILE), channel=cfg.get("browser_channel", "chrome"), headless=False,
            args=["--disable-blink-features=AutomationControlled"], viewport=None)
        try:
            pg = ctx.pages[0] if ctx.pages else ctx.new_page()
            pg.goto(PAGE, timeout=90000, wait_until="domcontentloaded")
            pg.wait_for_timeout(12000)

            def frame():
                return next((f for f in pg.frames if "quikstrike.net" in f.url), None)

            f = frame()
            if f is None:
                raise LoginRequired("QuikStrike not available - CME session expired; run quikstrike_login.py")

            f.locator("#ctl11_hlProductArrow").click()
            pg.wait_for_timeout(2500)
            pop = f.locator("#ctl11_ucProductSelectorPopup_pnlProductSelectorPopup")
            pop.get_by_text(cfg.get("qs_asset_class", "Metals"), exact=True).first.click()
            pg.wait_for_timeout(2500)
            pop.get_by_text(cfg.get("qs_product", "Gold (OG|GC)"), exact=True).first.click()
            pg.wait_for_timeout(9000)

            f = frame()
            f.locator(GREEK).select_option(label="* Gamma (1 Pct)")
            pg.wait_for_timeout(7000)
            f = frame()
            f.locator(STRIKES).select_option(label="(All)")
            pg.wait_for_timeout(7000)

            f = frame()
            f.get_by_text("EXPIRATION:", exact=False).first.click()
            pg.wait_for_timeout(2500)
            # popup entries read "<code> <date>"; anchor at the start so the toolbar label
            # "EXPIRATION: <code>" (which also contains the code) is never the click target
            f.locator(rf"text=/^\s*{code}(\s|$)/ >> visible=true").last.click()
            pg.wait_for_timeout(9000)

            f = frame()
            greek = f.locator(GREEK).evaluate("e=>e.options[e.selectedIndex].text")
            if "Gamma (1 Pct)" not in greek:
                raise RuntimeError(f"Greek selector reset to {greek!r}")
            table = f.evaluate("""()=>{const t=document.querySelector('table.grid-thm');
                return [...t.rows].map(r=>[...r.cells].map(c=>c.innerText.replace(/\\s+/g,' ').trim()))}""")
        finally:
            ctx.close()

    header = next(r for r in table if r and r[0] == "STRIKE")
    want = f"{trade_date.month}/{trade_date.day}/{trade_date.year}"
    if want not in header:
        raise RuntimeError(f"{code}: no gamma column for {want} (have {header[1:4]}...)")
    col = 1 + 2 * (header.index(want) - 1)
    out = {}
    for r in table:
        if r and r[0].isdigit() and len(r) > col + 1:
            out[int(r[0])] = (_num(r[col]), _num(r[col + 1]))
    return out


if __name__ == "__main__":
    import sys
    from fetch_cme import fetch
    cfg = json.load(open("config.json", encoding="utf-8"))
    d = fetch(cfg)
    code = qs_code(d["family"], d["label"])
    g = fetch_gamma(cfg, code, d["trade_date"])
    nz = {k: v for k, v in g.items() if v[0] or v[1]}
    print(code, d["trade_date"], len(g), "strikes,", len(nz), "nonzero", sum(v[0] for v in nz.values()), sum(v[1] for v in nz.values()))
