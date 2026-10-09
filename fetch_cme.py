"""Fetch option open interest per strike from CME Group.

CME blocks plain HTTP clients and headless browsers, so this drives a real
(headed) Chrome via Playwright and calls the same JSON endpoints the CME
volume/OI page uses.
"""
import calendar
import json
import os
import re
import tempfile
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from util import launch_persistent

PAGE = "https://www.cmegroup.com/markets/metals/precious/gold.volume.options.html"
MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
MONTH_CODE = "FGHJKMNQUVXZ"  # futures month letters, Jan..Dec


WEEKDAY = {"E21": 4, "MW1": 0, "AB1": 1, "WD1": 2, "BB1": 3}  # Fri, Mon, Tue, Wed, Thu


def expiry_date(family, label):
    """'Week 3 - SEP 2026' on a Monday product -> 3rd Monday of Sep 2026."""
    w, mon, year = int(label[5]), label[9:12], int(label[13:17])
    month = MONTHS.index(mon) + 1
    days = [d for d in range(1, calendar.monthrange(year, month)[1] + 1)
            if date(year, month, d).weekday() == WEEKDAY[family]]
    return date(year, month, days[min(w, len(days)) - 1])


def _num(s):
    s = str(s).replace(",", "").strip()
    return int(s) if s.lstrip("-").isdigit() else 0


def _label_key(label):
    """'Week 2 - SEP 2026' -> (2026, 9, 2); 'OCT 2026' -> (2026, 10, 0) for chronological sorting."""
    m = re.match(r"Week (\d+) - (\w{3}) (\d{4})", label)
    if m:
        return (int(m[3]), MONTHS.index(m[2]) + 1, int(m[1]))
    m = re.match(r"(\w{3}) (\d{4})", label)
    return (int(m[2]), MONTHS.index(m[1]) + 1, 0) if m else (9999, 0, 0)


def _monthly_dates(get, cfg):
    """{(month, year): last_trade_date} for the standard 'AME' series.

    The Volume/Options/Expirations response used below (via `groups`) doesn't carry an expiry
    date for AME the way it lets weeklies compute one via expiry_date(), so this asks the
    options-quotes page instead - same source and caveats as _underlying_future_month.

    Join key is the raw (month, year), not the label: this endpoint's own AME label is off by
    one from Volume/Options/Expirations' for the *same* contract (e.g. this endpoint calls the
    option that Volume/Options/Expirations calls "OCT 2026" a "Nov 2026" - a CME display
    convention, confirmed live 2026-09-27) but the underlying (month, year) numbers agree.
    """
    try:
        groups = get(f"/CmeWS/mvc/atm/expirations/{cfg['underlying_product_id']}?isProtected")
        group = next(g for g in groups if g["optionType"] == "AME")
        return {(e["expirationMonth"], e["expirationYear"]):
                 datetime.strptime(e["lastTradeDate"][:10], "%Y-%m-%d").date()
                for e in group["contractExpirations"]}
    except (StopIteration, KeyError):
        return {}


def _candidates(get, groups, cfg, allow_next=True):
    """Contracts to try, in order.

    target: 'today' | 'YYYY-MM-DD' | 'auto' | 'Week N - MON YYYY' | 'MON YYYY' (standard/AME).
    Date targets ('today'/'YYYY-MM-DD') search every family, weekly and standard, and pick
    whichever contract actually expires that day; 'auto'/a label only search cfg["family"]
    (set it to "AME" to target the standard monthly series by label).
    """
    target = cfg["target"]
    wanted = None
    if target == "today":
        wanted = date.today()
    elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", target):
        wanted = date.fromisoformat(target)

    monthly_dates = None  # fetched lazily - only if an AME group actually needs it
    out = []
    for g in groups:
        fam = g["optionType"]
        if wanted is None and fam != cfg["family"]:
            continue
        if fam in WEEKDAY:
            for e in g.get("expirations", []):
                if not e["label"].startswith("Week"):
                    continue
                out.append({**e, "family": fam, "_date": expiry_date(fam, e["label"])})
        elif fam == "AME" and cfg.get("include_monthly", True):
            monthly_dates = monthly_dates if monthly_dates is not None else _monthly_dates(get, cfg)
            for e in g.get("expirations", []):
                d = monthly_dates.get((e["expiration"]["month"], e["expiration"]["year"]))
                if d:
                    out.append({**e, "family": fam, "_date": d})
    if wanted is None:
        out.sort(key=lambda e: _label_key(e["label"]))
        return out if target == "auto" else [e for e in out if e["label"] == target]
    hit = [e for e in out if e["_date"] == wanted]
    if not hit and not allow_next:
        return []
    if not hit:  # nothing expires that day (weekend/holiday): next upcoming expiry instead
        later = sorted((e for e in out if e["_date"] > wanted), key=lambda e: e["_date"])
        if not later:
            raise RuntimeError(f"no contract expires on or after {wanted}")
        print(f"no contract expires on {wanted}; using next expiry {later[0]['_date']}")
        hit = [e for e in later if e["_date"] == later[0]["_date"]]
    return hit


@contextmanager
def cme_session(cfg):
    """Yields get(url) -> parsed JSON, fetched from inside a real CME page.

    A throwaway persistent context rather than launch()+Browser.close(): on the Windows 11
    26200 PC (Playwright 1.61) Browser.close() hangs forever, while BrowserContext.close()
    on a persistent context returns normally. util.launch_persistent retries a Chrome that
    exits right at launch."""
    with sync_playwright() as p, tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        b = launch_persistent(
            p, tmp,
            channel=cfg.get("browser_channel", "chrome"),
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
        )
        try:
            pg = b.pages[0] if b.pages else b.new_page()
            pg.goto(PAGE, timeout=60000, wait_until="domcontentloaded")
            pg.wait_for_timeout(6000)

            def get(u):
                # fetch() has no timeout of its own and page.evaluate() waits forever, so a
                # stalled CME request would hang the whole run - race it against 30s
                status, text = pg.evaluate(
                    """u=>Promise.race([fetch(u).then(async r=>[r.status,await r.text()]),
                        new Promise(res=>setTimeout(()=>res([0,'timed out after 30s']),30000))])""", u)
                if status != 200:
                    raise RuntimeError(f"CME {status} for {u}: {text[:120]}")
                return json.loads(text)

            yield get
        finally:
            b.close()


def fetch(cfg):
    with cme_session(cfg) as get:
        dates = get("/CmeWS/mvc/Volume/TradeDates?exchange=CBOT&isProtected")
        if cfg.get("trade_date"):  # verify.py / backfill re-read one specific day
            td = next((d for d in dates if d["tradeDate"] == cfg["trade_date"]), None)
            if td is None:
                raise RuntimeError(f"CME has no data for trade date {cfg['trade_date']}")
            tries = [td]
        else:
            tries = dates[:2]
        # A trade date CME has only just started publishing is incomplete for a while: first
        # its groups have no "expirations", then the contracts show zero OI and today's weekly
        # may not be listed yet (all seen 2026-10-02 for 10/01 PRELIMINARY). Exact-expiry pass
        # over both days first, so a half-published day can neither crash the run nor make it
        # silently take the *next* expiry; "next expiry" is the last resort.
        for allow_next in (False, True):
            for i, td in enumerate(tries):
                trade_date, report = td["tradeDate"], "P" if td["reportType"] == "PRELIMINARY" else "F"
                out = _fetch_trade_date(get, cfg, trade_date, report, allow_next)
                if out:
                    if i:
                        print(f"CME trade date {tries[0]['tradeDate']} not fully published yet; "
                              f"used {trade_date}")
                    return out
        raise RuntimeError("no contract with open interest found")


def _fetch_trade_date(get, cfg, trade_date, report, allow_next=True):
    """The target contract's OI on one trade date, or None if CME has nothing for it yet."""
    groups = get(f"/CmeWS/mvc/Volume/Options/Expirations?productid="
                 f"{cfg['underlying_product_id']}&tradedate={trade_date}&isProtected")
    if not any(g.get("expirations") for g in groups):
        return None
    try:
        exps = _candidates(get, groups, cfg, allow_next)
    except RuntimeError:
        return None
    for e in exps:
        d = get(f"/CmeWS/mvc/Volume/Options/Details?productid={e['productId']}"
                f"&tradedate={trade_date}&expirationcode={e['expirationCode']}"
                f"&reporttype={report}&isProtected")
        calls = puts = None
        oi = {}
        for md in d["monthData"]:
            side = "call" if md["monthID"].endswith("Calls") else "put"
            total = _num(md["totalData"]["atClose"])
            if side == "call":
                calls = (calls or 0) + total
            else:
                puts = (puts or 0) + total
            for s in md["strikeData"]:
                k = _num(s["strike"])
                oi.setdefault(k, {"call": 0, "put": 0})[side] += _num(s["atClose"])
        if (calls or 0) + (puts or 0) == 0:
            continue  # expired / empty contract
        lo, hi = cfg.get("strike_min", 0), cfg.get("strike_max", 10**9)
        strikes = [(k, v["call"], v["put"]) for k, v in sorted(oi.items())
                   if lo <= k <= hi and k > 0]
        derived = _underlying_future_month(get, cfg, e["family"], e["_date"])
        price_cfg = {**cfg, "price_month": cfg.get("price_month") or derived}
        price, price_month, change = _settle_price(get, price_cfg, trade_date)
        live, live_time = _live_quote(get, cfg, price_month) if price_month else (None, None)
        return {
            "live_price": live,  # CME's 10-min delayed quote at run time (update_gex.ref_price)
            "live_time": live_time,
            "trade_date": datetime.strptime(trade_date, "%Y%m%d").date(),
            "report": "PRELIMINARY" if report == "P" else "FINAL",
            "label": e["label"],
            "family": e["family"],
            "code": e["expirationCode"],
            "strikes": strikes,
            "call_total": calls,
            "put_total": puts,
            "price": price,
            "price_month": price_month,
            "change": change,  # futures settle change vs the previous day (rules.R2)
        }
    return None


def _underlying_future_month(get, cfg, family, exp_date):
    """CME's own answer for which futures month a given option series settles against
    (e.g. 'GCZ6' -> 'DEC 26'), read from the options-quotes page - no login needed, no guessing.

    This list only covers the next ~4 expirations per weekday, so it returns None for
    contracts outside that rolling window (verify.py re-checking an old date, etc.); callers
    should fall back to the OI-based heuristic in that case.
    """
    try:
        groups = get(f"/CmeWS/mvc/atm/expirations/{cfg['underlying_product_id']}?isProtected")
        group = next(g for g in groups if g["optionType"] == family)
        match = next(e for e in group["contractExpirations"] if e["lastTradeDate"][:10] == str(exp_date))
        code = match["underlyingFutureContract"]  # e.g. "GCZ6": product, month letter, single-digit year
        month = MONTHS[MONTH_CODE.index(code[-2])]
        return f"{month} 2{code[-1]}"  # "2" + digit: holds for 2020-2029, same as CME's own settle rows
    except (StopIteration, KeyError, IndexError, ValueError):
        return None


def _settle_price(get, cfg, trade_date):
    """(settle, month) of the futures used as the price; (None, None) if unavailable.

    cfg["price_month"] (e.g. "DEC 26") - usually filled in by fetch() from CME's own
    underlying-future answer, see _underlying_future_month - picks the settlement row.
    Falls back to the first month with OI > 10,000 if that's unset or not found.
    """
    try:
        row = futures_row(get, cfg, trade_date)
        if row:
            return px(row["settle"]), row["month"], px(row.get("change"))
    except Exception:
        pass
    return None, None, None


def px(s):
    """CME price text ('4,381.0B', '+23.2', '-', '') -> float, or None if there's no number."""
    m = re.match(r"[-+]?\d[\d,]*\.?\d*", str(s or "").strip())
    return float(m.group().replace(",", "")) if m else None


def futures_row(get, cfg, trade_date):
    """The raw CME settlement row (open/high/low/last/settle/...) for the futures month
    cfg["price_month"], else the first month with OI > 10,000; None if neither exists."""
    td = f"{trade_date[4:6]}%2F{trade_date[6:]}%2F{trade_date[:4]}"
    rows = get(f"/CmeWS/mvc/Settlements/Futures/Settlements/{cfg['underlying_product_id']}"
               f"/FUT?strategy=DEFAULT&tradeDate={td}&pageSize=500&isProtected")["settlements"]
    want = cfg.get("price_month")
    hit = next((r for r in rows if want and r["month"] == want), None)
    return hit or next((r for r in rows if _num(r["openInterest"]) > 10000), None)


def _live_quote(get, cfg, price_month):
    """(last, updated datetime UTC) of the price futures right now - CME's own quote, 10 min
    delayed - or (None, None). The settle is the previous session's close, which is what the
    OI/gamma data belong to, but it can sit tens of dollars away from the live chart."""
    try:
        want = f"{price_month[:3]} 20{price_month[-2:]}"  # "DEC 26" -> "DEC 2026"
        d = get(f"/CmeWS/mvc/quotes/v2/{cfg['underlying_product_id']}?isProtected")
        q = next(q for q in d["quotes"] if q["expirationMonth"] == want)
        last = px(q["last"])
        if last is None:  # "-": no trade yet this session
            return None, None
        return last, datetime.fromisoformat(q["updated"].replace("Z", "+00:00"))
    except Exception:
        return None, None


if __name__ == "__main__":
    cfg = json.load(open(os.environ.get("MAXPAIN_CONFIG", Path(__file__).parent / "config.json"), encoding="utf-8"))
    r = fetch(cfg)
    print({k: v for k, v in r.items() if k != "strikes"}, len(r["strikes"]), "strikes")
    print(r["strikes"][:5])
