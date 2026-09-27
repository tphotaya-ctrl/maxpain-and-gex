"""Backfill Log-sheet history for trade dates CME still has data for.

CME's own trade-date list is a short rolling window (currently ~5 trading days), so this
catches up a few missed days (machine off, expired login) - it is not a way to build long-run
history, since CME simply stops serving older days through these endpoints. Never touches
`OI Data` / `Gamma Data`, which only ever hold the most recent day's snapshot; only appends
to `Log`, and skips any date already logged (safe to run repeatedly).

    python backfill.py            # Max Pain workbook only
    python backfill.py --gex      # also backfill the GEX workbook (needs the QuikStrike login)

For each historical trade date, the contract picked is whichever one CME's own date-matching
would pick for a `target` of *that day* (mirrors `target: today`, anchored to the day being
backfilled instead of today's wall-clock date) - not whatever `config.json`'s own `target` says,
which usually means "today" and would be wrong for every day but the most recent one.
"""
import argparse
import json
import os
from datetime import datetime
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright

from fetch_cme import PAGE, fetch
from fetch_gamma import fetch_gamma, qs_code
from update_gex import LOG_HEAD as GEX_LOG_HEAD
from update_gex import gamma_flip, select_rows
from update_workbooks import LOG_HEAD as MP_LOG_HEAD
from update_workbooks import max_pain
from util import save_atomic

HERE = Path(__file__).parent
CONFIG = Path(os.environ.get("MAXPAIN_CONFIG", HERE / "config.json"))


def _trade_dates(cfg):
    """All trade dates CME's own history currently covers, newest first."""
    with sync_playwright() as p:
        b = p.chromium.launch(channel=cfg.get("browser_channel", "chrome"), headless=False,
                               args=["--disable-blink-features=AutomationControlled"])
        try:
            pg = b.new_page()
            pg.goto(PAGE, timeout=60000, wait_until="domcontentloaded")
            pg.wait_for_timeout(6000)
            text = pg.evaluate("u=>fetch(u).then(r=>r.text())",
                                "/CmeWS/mvc/Volume/TradeDates?exchange=CBOT&isProtected")
            return [d["tradeDate"] for d in json.loads(text)]
        finally:
            b.close()


def _existing_dates(path, sheet):
    if not Path(path).exists():
        return set()
    wb = openpyxl.load_workbook(path, read_only=True)
    if sheet not in wb.sheetnames:
        return set()
    return {r[0].date() for r in wb[sheet].iter_rows(min_row=2, values_only=True)
            if isinstance(r[0], datetime)}


def backfill_max_pain(cfg, trade_dates):
    path = HERE / cfg["workbook"]
    have = _existing_dates(path, "Log")
    wb = openpyxl.load_workbook(path)
    if "Log" not in wb.sheetnames:
        wb.create_sheet("Log").append(MP_LOG_HEAD)
    log = wb["Log"]
    added = 0
    for td in trade_dates:
        if datetime.strptime(td, "%Y%m%d").date() in have:
            continue
        target = f"{td[:4]}-{td[4:6]}-{td[6:]}"
        try:
            d = fetch({**cfg, "trade_date": td, "target": target, "strike_min": 0, "strike_max": 10**9})
        except Exception as e:
            print(f"Max Pain: skip {td} ({type(e).__name__}: {e})")
            continue
        mp, low, valley = max_pain(d["strikes"])
        calls, puts = sum(s[1] for s in d["strikes"]), sum(s[2] for s in d["strikes"])
        total = d["call_total"] + d["put_total"]
        log.append([d["trade_date"], f"{d['family']} {d['label']} ({d['code']})", None, d["price"], mp,
                    (mp - d["price"]) / d["price"] if d["price"] else None, valley,
                    puts / calls if calls else None, (calls + puts) / total if total else None,
                    d["report"], datetime.now(), d["price_month"]])
        added += 1
        print(f"Max Pain: logged {td} - {d['family']} {d['label']} maxpain={mp}")
    if added:
        save_atomic(wb, path)
    print(f"Max Pain backfill: {added} day(s) added, {len(trade_dates) - added} already had a row or failed")


def backfill_gex(cfg, trade_dates):
    path = HERE / cfg["gex_workbook"]
    have = _existing_dates(path, "Log")
    wb = openpyxl.load_workbook(path)
    if "Log" not in wb.sheetnames:
        wb.create_sheet("Log").append(GEX_LOG_HEAD)
    log = wb["Log"]
    added = 0
    for td in trade_dates:
        if datetime.strptime(td, "%Y%m%d").date() in have:
            continue
        target = f"{td[:4]}-{td[4:6]}-{td[6:]}"
        try:
            d = fetch({**cfg, "trade_date": td, "target": target, "strike_min": 0, "strike_max": 10**9})
            code = qs_code(d["family"], d["label"])
            g = fetch_gamma(cfg, code, d["trade_date"])
        except Exception as e:
            print(f"GEX: skip {td} ({type(e).__name__}: {e})")
            continue
        rows = select_rows(g)
        calls, puts = sum(r[1] for r in rows), sum(r[2] for r in rows)
        net = calls - puts
        status = "Positive GEX (นิ่ง)" if net > 0 else "Negative GEX (แกว่งแรง)"
        log.append([d["trade_date"], code, d["price"], calls, puts, net, status,
                    gamma_flip(rows, d["price"]), max(rows, key=lambda r: r[1] + r[2])[0], datetime.now()])
        added += 1
        print(f"GEX: logged {td} - {code} net={net:+.0f}")
    if added:
        save_atomic(wb, path)
    print(f"GEX backfill: {added} day(s) added, {len(trade_dates) - added} already had a row or failed")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--gex", action="store_true", help="also backfill the GEX workbook")
    args = ap.parse_args()
    cfg = json.load(open(CONFIG, encoding="utf-8"))
    trade_dates = _trade_dates(cfg)
    print(f"CME currently has {len(trade_dates)} trade date(s) available: {trade_dates[-1]}..{trade_dates[0]}")
    backfill_max_pain(cfg, trade_dates)
    if args.gex:
        backfill_gex(cfg, trade_dates)


if __name__ == "__main__":
    main()
