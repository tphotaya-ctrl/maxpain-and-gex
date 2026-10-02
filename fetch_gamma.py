"""Fetch per-strike Call/Put Gamma (1 Pct) from CME QuikStrike (OI Matrix view).

QuikStrike needs a CME login, so this reuses the persistent Chrome profile
created by quikstrike_login.py. The matrix for a single expiration lists one
C/P column pair per trade date, so any recent day can be read (backfill too).
"""
import json
import re
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


def _step(pg, frame_getter, name, action, verify, waits=(3000, 6000, 10000)):
    """Run `action(frame)`, wait, then check `verify(frame)`; retry with longer waits if it fails.

    QuikStrike's ASP.NET postbacks have no reliable "done" event to await, so this is a
    poll-and-retry rather than a real wait condition. Used for the two dropdowns, which have
    an obvious pass/fail readback (the selected option's own text).
    """
    last = None
    for wait in waits:
        f = frame_getter()
        try:
            if f is None:
                raise RuntimeError("QuikStrike frame not found")
            action(f)
            pg.wait_for_timeout(wait)
            f = frame_getter()
            if f is not None and verify(f):
                return f
            last = RuntimeError(f"{name}: not verified after a {wait}ms wait")
        except Exception as e:
            last = e
        print(f"fetch_gamma: retrying {name} ({last})")
    raise RuntimeError(f"{name} failed after {len(waits)} attempts: {last}")


def fetch_gamma(cfg, code, trade_date: date, retries: int = 3):
    """Return {strike: (call_gamma, put_gamma)} for `code` on `trade_date`.

    QuikStrike is a fixed-wait ASP.NET UI (see _step) with steps - product popup, expiration
    popup - that have no cheap readback to poll. Those are covered by retrying the whole
    fetch from a fresh page instead: cheaper to reason about than guessing at internal state,
    and directly fixes the timeout seen in practice (see verify.py run, 2026-09-21).
    """
    last = None
    for attempt in range(1, retries + 1):
        try:
            return _fetch_gamma_once(cfg, code, trade_date)
        except LoginRequired:
            raise  # retrying won't fix a missing/expired login
        except DateNotInHistory as e:
            # The per-expiration history view lags a day behind the default expiry matrix
            # (seen 2026-10-01: history topped out at 9/29 while the matrix already showed
            # "Wed, Sep 30, 2026"). The matrix only shows its latest trade date, so it's only
            # usable when that date is the one we want - _fetch_latest_once checks that.
            print(f"fetch_gamma: {e} - trying the latest-day expiry matrix instead")
            try:
                return _fetch_latest_once(cfg, code, trade_date)
            except LoginRequired:
                raise
            except Exception as e2:
                last = e2
        except Exception as e:
            last = e
        print(f"fetch_gamma: attempt {attempt}/{retries} failed ({type(last).__name__}: {last})")
    raise RuntimeError(f"fetch_gamma failed after {retries} attempts: {last}") from last


class DateNotInHistory(RuntimeError):
    pass


def _open_gamma_matrix(cfg, p):
    """Open QuikStrike with Gold / Gamma (1 Pct) / Strikes (All); return (ctx, pg, frame_getter)."""
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

        def selected(sel, want):
            return lambda fr: want in fr.locator(sel).evaluate("e=>e.options[e.selectedIndex].text")

        _step(pg, frame, "select Greek=Gamma (1 Pct)",
              lambda fr: fr.locator(GREEK).select_option(label="* Gamma (1 Pct)"),
              selected(GREEK, "Gamma (1 Pct)"))
        _step(pg, frame, "select Strikes=(All)",
              lambda fr: fr.locator(STRIKES).select_option(label="(All)"),
              selected(STRIKES, "(All)"))
        return ctx, pg, frame
    except Exception:
        ctx.close()
        raise


def _read_table(f):
    greek = f.locator(GREEK).evaluate("e=>e.options[e.selectedIndex].text")
    if "Gamma (1 Pct)" not in greek:
        raise RuntimeError(f"Greek selector reset to {greek!r}")
    return f.evaluate("""()=>{const t=document.querySelector('table.grid-thm');
        return [...t.rows].map(r=>[...r.cells].map(c=>c.innerText.replace(/\\s+/g,' ').trim()))}""")


def _wait_table(pg, frame, ready, timeout_ms=40000, every_ms=1500):
    """Poll the matrix until `ready(table)` - the postbacks have no "done" event, and a fixed
    wait was sometimes too short (expiry view still showing, or only the default 31 strikes)."""
    table = None
    for _ in range(max(1, timeout_ms // every_ms)):
        pg.wait_for_timeout(every_ms)
        try:
            table = _read_table(frame())
            if ready(table):
                return table
        except Exception:
            pass
    return table  # caller's own checks report what was wrong with it


def _header(table):
    return next((r for r in table or [] if r and r[0] == "STRIKE"), [])


def _strike_rows(table):
    return sum(1 for r in table or [] if r and r[0].isdigit())


def _column(table, header_index):
    """Rows of the C/P pair under header cell `header_index` -> {strike: (call, put)}."""
    col = 1 + 2 * (header_index - 1)
    out = {}
    for r in table:
        if r and r[0].isdigit() and len(r) > col + 1:
            out[int(r[0])] = (_num(r[col]), _num(r[col + 1]))
    return out


def _fetch_gamma_once(cfg, code, trade_date: date):
    with sync_playwright() as p:
        ctx, pg, frame = _open_gamma_matrix(cfg, p)
        try:
            f = frame()
            f.get_by_text("EXPIRATION:", exact=False).first.click()
            pg.wait_for_timeout(2500)
            # popup entries read "<code> <date>"; anchor at the start so the toolbar label
            # "EXPIRATION: <code>" (which also contains the code) is never the click target
            f.locator(rf"text=/^\s*{code}(\s|$)/ >> visible=true").last.click()
            # ready = the per-trade-date view (date headers) with the (All) strike list
            table = _wait_table(pg, frame, lambda t: any("/" in h for h in _header(t))
                                and _strike_rows(t) > 60)
            if table is None:
                raise RuntimeError("QuikStrike matrix never rendered")
        finally:
            ctx.close()

    header = next(r for r in table if r and r[0] == "STRIKE")
    want = f"{trade_date.month}/{trade_date.day}/{trade_date.year}"
    if want not in header:
        if any(h.startswith(f"{code} ") for h in header):
            # expiration click didn't take - the default expiry matrix is still showing
            raise RuntimeError(f"{code}: expiration not selected (have {header[1:4]}...)")
        raise DateNotInHistory(f"{code}: no gamma column for {want} (have {header[1:4]}...)")
    return _column(table, header.index(want))


def _fetch_latest_once(cfg, code, trade_date: date):
    """Read `code`'s column from the default expiry matrix (one column per expiration,
    latest trade date only), refusing unless the matrix's date is `trade_date`."""
    with sync_playwright() as p:
        ctx, pg, frame = _open_gamma_matrix(cfg, p)
        try:
            # the (All) strikes table renders after the dropdown readback
            table = _wait_table(pg, frame, lambda t: _strike_rows(t) > 60
                                and any(h.startswith(f"{code} ") for h in _header(t)))
            if table is None:
                raise RuntimeError("QuikStrike matrix never rendered")
            body = frame().evaluate("()=>document.body.innerText")
        finally:
            ctx.close()

    return _latest_column(table, body, code, trade_date)


def matrix_date(body):
    """The expiry matrix's trade date, from its toolbar text e.g. 'Wed, Sep 30, 2026'."""
    m = re.search(r"\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun), (\w{3}) (\d{1,2}), (\d{4})", body)
    if not m:
        raise RuntimeError("expiry matrix: trade date not found on page")
    return date(int(m[3]), MONTHS.index(m[1].upper()) + 1, int(m[2]))


def _latest_column(table, body, code, trade_date):
    shown = matrix_date(body)
    if shown != trade_date:
        raise RuntimeError(f"expiry matrix shows {shown}, wanted {trade_date}")
    header = _header(table)
    idx = next((i for i, h in enumerate(header) if h.startswith(f"{code} ")), None)
    if idx is None:
        raise RuntimeError(f"expiry matrix has no {code} column (have {header[1:4]}...)")
    out = _column(table, idx)
    if len(out) < 60:
        raise RuntimeError(f"expiry matrix only had {len(out)} strikes - Strikes (All) not applied yet")
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
