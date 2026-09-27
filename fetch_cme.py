"""Fetch option open interest per strike from CME Group.

CME blocks plain HTTP clients and headless browsers, so this drives a real
(headed) Chrome via Playwright and calls the same JSON endpoints the CME
volume/OI page uses.
"""
import calendar
import json
import os
import re
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
    """'Week 2 - SEP 2026' -> (2026, 9, 2) for chronological sorting."""
    m = re.match(r"Week (\d+) - (\w{3}) (\d{4})", label)
    return (int(m[3]), MONTHS.index(m[2]) + 1, int(m[1])) if m else (9999, 0, 0)


def _candidates(groups, cfg):
    """Contracts to try, in order. target: 'today' | 'YYYY-MM-DD' | 'auto' | 'Week N - MON YYYY'."""
    target = cfg["target"]
    wanted = None
    if target == "today":
        wanted = date.today()
    elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", target):
        wanted = date.fromisoformat(target)
    out = []
    for g in groups:
        fam = g["optionType"]
        if fam not in WEEKDAY or (wanted is None and fam != cfg["family"]):
            continue
        for e in g["expirations"]:
            if not e["label"].startswith("Week"):
                continue
            out.append({**e, "family": fam, "_date": expiry_date(fam, e["label"])})
    if wanted is None:
        out.sort(key=lambda e: _label_key(e["label"]))
        return out if target == "auto" else [e for e in out if e["label"] == target]
    hit = [e for e in out if e["_date"] == wanted]
    if not hit:  # nothing expires that day (weekend/holiday): next upcoming expiry instead
        later = sorted((e for e in out if e["_date"] > wanted), key=lambda e: e["_date"])
        if not later:
            raise RuntimeError(f"no weekly contract expires on or after {wanted}")
        print(f"no contract expires on {wanted}; using next expiry {later[0]['_date']}")
        hit = [e for e in later if e["_date"] == later[0]["_date"]]
    return hit


def fetch(cfg):
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

            dates = get("/CmeWS/mvc/Volume/TradeDates?exchange=CBOT&isProtected")
            td = dates[0]
            if cfg.get("trade_date"):  # verify.py re-reads the same day the workbook holds
                td = next((d for d in dates if d["tradeDate"] == cfg["trade_date"]), None)
                if td is None:
                    raise RuntimeError(f"CME has no data for trade date {cfg['trade_date']}")
            trade_date, report = td["tradeDate"], "P" if td["reportType"] == "PRELIMINARY" else "F"

            groups = get(f"/CmeWS/mvc/Volume/Options/Expirations?productid="
                         f"{cfg['underlying_product_id']}&tradedate={trade_date}&isProtected")
            exps = _candidates(groups, cfg)

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
                price, price_month = _settle_price(get, price_cfg, trade_date)
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
                }
            raise RuntimeError("no contract with open interest found")
        finally:
            b.close()


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
        td = f"{trade_date[4:6]}%2F{trade_date[6:]}%2F{trade_date[:4]}"
        d = get(f"/CmeWS/mvc/Settlements/Futures/Settlements/{cfg['underlying_product_id']}"
                f"/FUT?strategy=DEFAULT&tradeDate={td}&pageSize=500&isProtected")
        rows = d["settlements"]
        want = cfg.get("price_month")
        if want:
            hit = next((r for r in rows if r["month"] == want), None)
            if hit:
                return float(hit["settle"].replace(",", "")), hit["month"]
        for row in rows:
            if _num(row["openInterest"]) > 10000:
                return float(row["settle"].replace(",", "")), row["month"]
    except Exception:
        pass
    return None, None


if __name__ == "__main__":
    cfg = json.load(open(os.environ.get("MAXPAIN_CONFIG", Path(__file__).parent / "config.json"), encoding="utf-8"))
    r = fetch(cfg)
    print({k: v for k, v in r.items() if k != "strikes"}, len(r["strikes"]), "strikes")
    print(r["strikes"][:5])
