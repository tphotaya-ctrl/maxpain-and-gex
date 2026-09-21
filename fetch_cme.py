"""Fetch option open interest per strike from CME Group.

CME blocks plain HTTP clients and headless browsers, so this drives a real
(headed) Chrome via Playwright and calls the same JSON endpoints the CME
volume/OI page uses.
"""
import calendar
import json
import re
from datetime import date, datetime

from playwright.sync_api import sync_playwright

PAGE = "https://www.cmegroup.com/markets/metals/precious/gold.volume.options.html"
MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


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
                return {
                    "trade_date": datetime.strptime(trade_date, "%Y%m%d").date(),
                    "report": "PRELIMINARY" if report == "P" else "FINAL",
                    "label": e["label"],
                    "family": e["family"],
                    "code": e["expirationCode"],
                    "strikes": strikes,
                    "call_total": calls,
                    "put_total": puts,
                    "price": _settle_price(get, cfg, trade_date),
                }
            raise RuntimeError("no contract with open interest found")
        finally:
            b.close()


def _settle_price(get, cfg, trade_date):
    """Settle of the first liquid futures month (OI > 10,000); None if unavailable."""
    try:
        td = f"{trade_date[4:6]}%2F{trade_date[6:]}%2F{trade_date[:4]}"
        d = get(f"/CmeWS/mvc/Settlements/Futures/Settlements/{cfg['underlying_product_id']}"
                f"/FUT?strategy=DEFAULT&tradeDate={td}&pageSize=500&isProtected")
        for row in d["settlements"]:
            if _num(row["openInterest"]) > 10000:
                return float(row["settle"].replace(",", ""))
    except Exception:
        pass
    return None


if __name__ == "__main__":
    cfg = json.load(open("config.json", encoding="utf-8"))
    r = fetch(cfg)
    print({k: v for k, v in r.items() if k != "strikes"}, len(r["strikes"]), "strikes")
    print(r["strikes"][:5])
