"""Hourly price alerts: a Telegram message the first time each trading day that gold reaches
one of the logged levels (Max Pain, Call/Put Wall, Gamma Flip, the all-expiration walls).

Runs in GitHub Actions (.github/workflows/price-alerts.yml), so it works with the PC off. The
levels come from the private Google Sheet the main PC syncs (sheets_sync.py) - QuikStrike, and
so GEX, only exists on the PC. The price is Yahoo's GC=F (COMEX front month, plain HTTP, no key),
shifted onto the logged futures month by the settle difference (basis). Which levels were
already sent, and up to when bars were checked, live in the Sheet's "Alerts" tab - the job
never writes to the (public) repo.

    python price_alerts.py                                    # the cloud job
    python price_alerts.py --dry-run --levels-from-workbook   # local check: print, send nothing
"""
import json
import os
import sys
import urllib.request
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

HERE = Path(__file__).parent
CONFIG = Path(os.environ.get("MAXPAIN_CONFIG", HERE / "config.json"))
YAHOO = "https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?interval={interval}&range={range}"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"
BKK = timezone(timedelta(hours=7))
ALERTS = "Alerts"
ALERT_HEAD = ["วันที่ข้อมูล", "ระดับ", "ค่า", "ส่งเมื่อ (UTC)"]
LAST_CHECKED = "last_checked"
STALE_DAYS = 3        # Friday's levels still count on Monday; older ones belong to an expired contract
FIRST_LOOKBACK = timedelta(minutes=75)  # first run ever: one hourly slot plus cron jitter
MAX_LOOKBACK = timedelta(hours=6)       # after a long gap (runner outage) don't replay old touches
SETTLE_UTC = 18                          # COMEX gold settles 13:30 New York, by 18:30 UTC all year
BASIS_WARN = 0.01

# (Sheet tab, column) -> alert name. Tab names are the ones sheets_sync.py writes.
SOURCES = {
    "Max Pain Log": {"Max Pain": "Max Pain"},
    "GEX Log": {"Call Wall": "Call Wall", "Put Wall": "Put Wall", "Gamma Flip (ใกล้ราคา)": "Gamma Flip"},
    "GEX รวม": {"Call Wall": "Call Wall (รวม)", "Put Wall": "Put Wall (รวม)"},
}


def _day(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return date.fromisoformat(str(v)[:10])
    except ValueError:
        return None


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return None


def latest(rows):
    """{column: value} of the row with the newest date in column A (rows[0] is the header)."""
    if not rows or len(rows) < 2:
        return None
    header = rows[0]
    dated = [r for r in rows[1:] if r and _day(r[0])]
    if not dated:
        return None
    row = max(dated, key=lambda r: _day(r[0]))
    return {h: (row[i] if i < len(row) else None) for i, h in enumerate(header) if h}


def levels(tabs, today):
    """What to watch today: {"trade_date", "settle", "mode", "levels": {name: value}}.

    `tabs` is {tab name: rows incl. header}. Each tab's newest row is used only if it is at
    most STALE_DAYS old, so an expired contract's levels never alert."""
    out = {"trade_date": None, "settle": None, "mode": None, "levels": {}}
    for tab, cols in SOURCES.items():
        row = latest(tabs.get(tab))
        if not row:
            continue
        d = _day(row.get("วันที่ข้อมูล"))
        if (today - d).days > STALE_DAYS:
            print(f"{tab}: newest row is {d} - too old, skipped")
            continue
        for col, name in cols.items():
            v = _num(row.get(col))
            if v:
                out["levels"][name] = v
        if tab == "Max Pain Log":
            out["trade_date"], out["settle"] = d, _num(row.get("ราคา"))
        elif tab == "GEX Log":
            out["mode"] = row.get("สถานะ")
            out["trade_date"] = out["trade_date"] or d
    return out


def _get_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())


def parse_chart(d):
    """Yahoo chart JSON -> [(UTC datetime, high, low, close)], bars with no prices dropped."""
    res = d["chart"]["result"][0]
    q = res["indicators"]["quote"][0]
    out = []
    for i, ts in enumerate(res.get("timestamp") or []):
        h, l, c = q["high"][i], q["low"][i], q["close"][i]
        if None not in (h, l, c):
            out.append((datetime.fromtimestamp(ts, timezone.utc), h, l, c))
    return out


def yahoo_bars(symbol="GC=F", interval="5m", range_="2d"):
    return parse_chart(_get_json(YAHOO.format(symbol=symbol, interval=interval, range=range_)))


def daily_closes(bars):
    """Daily bars -> {trade date: close}. Yahoo stamps a futures day at midnight New York time
    (04:00/05:00 UTC), so the UTC date is the session; its close is the settle."""
    return {t.date(): c for t, _, _, c in bars}


def basis(settle, closes, trade_date):
    """(offset to add to Yahoo prices, warning or None). The levels are in the logged futures
    month (e.g. DEC 26); GC=F is the front month, which can be a different contract."""
    y = closes.get(trade_date)
    if settle is None or y is None:
        return 0.0, f"เทียบเดือนสัญญาไม่ได้ (ไม่มีราคาปิด {trade_date} จาก Yahoo)"
    off = settle - y
    return off, ("เดือนสัญญาต่างกัน: ราคา Yahoo ถูกเลื่อน {:+,.1f}".format(off)
                 if abs(off) > settle * BASIS_WARN else None)


def touches(bars, lv, since, offset=0.0):
    """[(name, level, 'up'|'down', UTC time, price then)] - the first bar after `since` whose
    high-low range reaches each level. Direction is where the price came from: the close of
    the bar before (or the bar's own open side when there is none)."""
    out = []
    for name, level in lv.items():
        prev = None
        for t, h, l, c in bars:
            h, l, c = h + offset, l + offset, c + offset
            if t > since and l <= level <= h:
                came_from = prev if prev is not None else (h + l) / 2
                out.append((name, level, "up" if came_from < level else "down", t, c))
                break
            prev = c
    return out


def read_state(rows):
    """Alerts tab rows -> (set of (trade date, name, value) already sent, last_checked or None)."""
    sent, last = set(), None
    for r in (rows or [])[1:]:
        if not r:
            continue
        if r[0] == LAST_CHECKED:
            try:
                last = datetime.fromisoformat(str(r[3]))
            except (IndexError, ValueError):
                pass
        elif len(r) >= 3 and _num(r[2]) is not None:
            sent.add((str(r[0])[:10], str(r[1]), _num(r[2])))
    return sent, last


def window_start(last, now):
    if last is None:
        return now - FIRST_LOOKBACK
    return max(last, now - MAX_LOOKBACK)


def new_alerts(hits, sent, trade_date):
    return [h for h in hits if (str(trade_date), h[0], h[1]) not in sent]


def format_alert(hit, info, price_now, warn=None):
    from telegram_report import reading
    name, level, direction, t, _ = hit
    arrow = "ขึ้นแตะ" if direction == "up" else "ลงแตะ"
    lines = [f"ราคาทอง{arrow} {name} {level:,.0f} เวลา {t.astimezone(BKK):%H:%M} (ไทย)",
             f"ราคาตอนนี้ ~{price_now:,.1f} · ระดับจากข้อมูลวันที่ {info['trade_date']}"]
    if info.get("mode"):
        lines.append(f"โหมด {info['mode']}")
    lv = info["levels"]
    lines += reading(info.get("mode"), price_now, lv.get("Call Wall"), lv.get("Put Wall"))
    if warn:
        lines.append(warn)
    return "\n".join(lines)


def _rows_from_workbook(cfg):
    """Same shape as the Sheet tabs, read from the local workbook (for --levels-from-workbook)."""
    from publish_report import _read_log
    from sheets_sync import _cell
    from update_gex import AGG_LOG, GEX_LOG
    tabs = {}
    for tab, sheet in (("Max Pain Log", "Log"), ("GEX Log", GEX_LOG), ("GEX รวม", AGG_LOG)):
        try:
            header, rows = _read_log(HERE / cfg["workbook"], n=5, sheet=sheet)
        except KeyError:
            continue
        header = header[:]
        if tab == "Max Pain Log" and header and header[-1] is None:
            header[-1] = "Futures ที่ใช้เป็นราคา"
        tabs[tab] = [header] + [[_cell(v) for v in r] for r in rows]
    return tabs


def _write_state(ws, new_rows, now, existing):
    if not existing:
        ws.append_row(ALERT_HEAD)
        existing = [ALERT_HEAD]
    if new_rows:
        ws.append_rows(new_rows)
    stamp = [LAST_CHECKED, "", "", now.isoformat(timespec="seconds")]
    hit = next((i for i, r in enumerate(existing, start=1) if r and r[0] == LAST_CHECKED), None)
    if hit:
        ws.update(f"A{hit}", [stamp])
    else:
        ws.append_row(stamp)


def run(dry_run=False, from_workbook=False, now=None):
    now = now or datetime.now(timezone.utc)
    cfg = json.load(open(CONFIG, encoding="utf-8"))
    ws = existing = None
    if from_workbook:
        tabs = _rows_from_workbook(cfg)
    else:
        import gspread
        from gspread.utils import ValueRenderOption
        from sheets_sync import open_sheet, _worksheet
        sh = open_sheet(cfg)
        tabs = {}
        for tab in SOURCES:
            try:
                tabs[tab] = sh.worksheet(tab).get_all_values(value_render_option=ValueRenderOption.unformatted)
            except gspread.WorksheetNotFound:
                pass
        ws = _worksheet(sh, ALERTS, ALERT_HEAD)
        existing = ws.get_all_values()
    info = levels(tabs, now.astimezone(BKK).date())
    print(f"levels from {'workbook' if from_workbook else 'Sheet'} ({info['trade_date']}): "
          + (", ".join(f"{k} {v:,.0f}" for k, v in info["levels"].items()) or "none"))
    sent, last = read_state(existing)
    since = window_start(last, now)
    sent_rows = []
    if info["levels"]:
        bars = yahoo_bars()
        offset, warn = basis(info["settle"], daily_closes(yahoo_bars(interval="1d", range_="10d")),
                             info["trade_date"])
        # only bars after the levels' own session closed: before that they describe the past
        since = max(since, datetime.combine(info["trade_date"], datetime.min.time(), timezone.utc)
                    + timedelta(hours=SETTLE_UTC, minutes=30))
        price_now = bars[-1][3] + offset if bars else None
        hits = new_alerts(touches(bars, info["levels"], since, offset), sent, info["trade_date"])
        print(f"{len(bars)} bars, offset {offset:+.1f}, checking since {since:%Y-%m-%d %H:%M} UTC: "
              f"{len(hits)} new alert(s)")
        for h in hits:
            text = format_alert(h, info, price_now, warn)
            print(text)
            if not dry_run:
                from telegram_report import send_text
                if not send_text(text):
                    raise RuntimeError("Telegram send failed")
            sent_rows.append([str(info["trade_date"]), h[0], h[1], now.isoformat(timespec="seconds")])
    if ws is not None and not dry_run:
        _write_state(ws, sent_rows, now, existing)
    return sent_rows


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    try:
        run(dry_run="--dry-run" in sys.argv, from_workbook="--levels-from-workbook" in sys.argv)
    except Exception as e:
        # no Telegram here: an hourly job that fails would message every hour. The red
        # Actions run (and GitHub's failure e-mail) is the signal.
        print(f"price alerts failed: {type(e).__name__}: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
