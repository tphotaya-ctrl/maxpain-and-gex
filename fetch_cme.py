"""Fetch option open interest per strike from CME Group.

CME blocks plain HTTP clients and headless browsers, so this drives a real
(headed) Chrome via Playwright and calls the same JSON endpoints the CME
volume/OI page uses.
"""
import calendar
import json
import os
import re
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

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


def _candidates(get, groups, cfg):
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
            for e in g["expirations"]:
                if not e["label"].startswith("Week"):
                    continue
                out.append({**e, "family": fam, "_date": expiry_date(fam, e["label"])})
        elif fam == "AME" and cfg.get("include_monthly", True):
            monthly_dates = monthly_dates if monthly_dates is not None else _monthly_dates(get, cfg)
            for e in g["expirations"]:
                d = monthly_dates.get((e["expiration"]["month"], e["expiration"]["year"]))
                if d:
                    out.append({**e, "family": fam, "_date": d})
    if wanted is None:
        out.sort(key=lambda e: _label_key(e["label"]))
        return out if target == "auto" else [e for e in out if e["label"] == target]
    hit = [e for e in out if e["_date"] == wanted]
    if not hit:  # nothing expires that day (weekend/holiday): next upcoming expiry instead
        later = sorted((e for e in out if e["_date"] > wanted), key=lambda e: e["_date"])
        if not later:
            raise RuntimeError(f"no contract expires on or after {wanted}")
        print(f"no contract expires on {wanted}; using next expiry {later[0]['_date']}")
        hit = [e for e in later if e["_date"] == later[0]["_date"]]
    return hit


@contextmanager
def cme_session(cfg):
    """Yields get(url) -> parsed JSON, fetched from inside a real CME page."""
    with sync_playwright() as p:
        b = p.chromium.launch(
            channel=cfg.get("browser_channel", "chrome"),
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
        )
        try:
            pg = b.new_page()
            pg.goto(PAGE, timeout=60000, wait_until="domcontentloaded")
            pg.wait_for_timeout(6000)

            def get(u):
                status, text = pg.evaluate(
                    "u=>fetch(u).then(async r=>[r.status,await r.text()])", u)
                if status != 200:
                    raise RuntimeError(f"CME {status} for {u}: {text[:120]}")
                return json.loads(text)

            yield get
        finally:
            b.close()


def fetch(cfg):
    with cme_session(cfg) as get:
        dates = get("/CmeWS/mvc/Volume/TradeDates?exchange=CBOT&isProtected")
        td = dates[0]
        if cfg.get("trade_date"):  # verify.py re-reads the same day the workbook holds
            td = next((d for d in dates if d["tradeDate"] == cfg["trade_date"]), None)
            if td is None:
                raise RuntimeError(f"CME has no data for trade date {cfg['trade_date']}")
        trade_date, report = td["tradeDate"], "P" if td["reportType"] == "PRELIMINARY" else "F"

        groups = get(f"/CmeWS/mvc/Volume/Options/Expirations?productid="
                     f"{cfg['underlying_product_id']}&tradedate={trade_date}&isProtected")
        exps = _candidates(get, groups, cfg)

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
            return {
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
        raise RuntimeError("no contract with open interest found")


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


if __name__ == "__main__":
    cfg = json.load(open(os.environ.get("MAXPAIN_CONFIG", Path(__file__).parent / "config.json"), encoding="utf-8"))
    r = fetch(cfg)
    print({k: v for k, v in r.items() if k != "strikes"}, len(r["strikes"]), "strikes")
    print(r["strikes"][:5])
