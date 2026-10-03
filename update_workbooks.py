"""Fetch today's CME data and write it into the Max Pain workbook + Log sheet."""
import json
import os
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
from copy import copy
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formula.translate import Translator

from charts import rebuild_charts
from fetch_cme import expiry_date, fetch
from notify import notify
from util import save_atomic

HERE = Path(__file__).parent
CONFIG = Path(os.environ.get("MAXPAIN_CONFIG", HERE / "config.json"))  # override to work on cloned files
MAX_ROWS = 200  # OI Data rows 2-201, the range covered by the Max Pain Calc formulas
OLD_LAST, NEW_LAST = 102, 201
LOG_HEAD = ["วันที่ข้อมูล", "สัญญา", "DTE", "ราคา", "Max Pain", "ห่างจากราคา %",
            "ก้นหุบ (จำนวน strike)", "P/C Ratio", "ครอบคลุม CME %", "รายงาน", "บันทึกเมื่อ", "Futures ที่ใช้เป็นราคา"]


def zone(measure, mid, edge):
    """Same bucketing as Max Pain Calc!F11 ('เทียบกับงานวิจัย'): None if there's no reading yet."""
    if measure is None:
        return None
    a = abs(measure)
    return "กลาง" if a < mid else ("เปลี่ยนผ่าน" if a < edge else "มีผล")


def max_pain(strikes):
    """Same maths as the Max Pain Calc sheet: pain at each strike, lowest wins."""
    pain = {}
    for k, _, _ in strikes:
        pain[k] = sum((k - s) * c for s, c, _ in strikes if s < k) + \
                  sum((s - k) * p for s, _, p in strikes if s > k)
    low = min(pain.values())
    return min(k for k, v in pain.items() if v == low), low, sum(1 for v in pain.values() if v < low * 1.02)


def extend_formulas(wb):
    """One-off: widen the Max Pain Calc formula range from row 102 to row 201 (idempotent)."""
    calc, oi = wb["Max Pain Calc"], wb["OI Data"]
    if "$A$%d" % NEW_LAST in str(calc["B2"].value):
        return
    for row in calc.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = c.value.replace("$%d" % OLD_LAST, "$%d" % NEW_LAST)
    for r in range(OLD_LAST + 1, NEW_LAST + 1):
        for col in "ABC":
            src, dst = calc[f"{col}{OLD_LAST}"], calc[f"{col}{r}"]
            dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
            dst._style = copy(src._style)
    old_cf, calc.conditional_formatting = calc.conditional_formatting, ConditionalFormattingList()
    for rng, rules in old_cf._cf_rules.items():
        for rule in rules:
            rule.formula = [f.replace("$%d" % OLD_LAST, "$%d" % NEW_LAST) for f in rule.formula]
            calc.conditional_formatting.add(str(rng.sqref).replace(str(OLD_LAST), str(NEW_LAST)), rule)
    for r in range(OLD_LAST + 1, NEW_LAST + 1):  # match OI Data cell styles
        for col in "ABC":
            oi[f"{col}{r}"]._style = copy(oi[f"{col}{OLD_LAST}"]._style)
    note = oi["A104"].value
    oi.unmerge_cells("A104:C106")
    oi["A104"].value = None
    oi.merge_cells(f"A{NEW_LAST + 2}:C{NEW_LAST + 4}")
    oi[f"A{NEW_LAST + 2}"].value = note.replace("2-102 (101 strike)", f"2-{NEW_LAST} ({MAX_ROWS} strike)")


def main():
    cfg = json.load(open(CONFIG, encoding="utf-8"))
    d = fetch(cfg)
    strikes = d["strikes"]
    if len(strikes) > MAX_ROWS:
        sys.exit(f"{len(strikes)} strikes exceeds formula range ({MAX_ROWS}); narrow strike_min/strike_max")

    path = HERE / cfg["workbook"]
    wb = openpyxl.load_workbook(path)
    oi, calc = wb["OI Data"], wb["Max Pain Calc"]
    extend_formulas(wb)

    for r in range(2, 2 + MAX_ROWS):
        for c in (1, 2, 3):
            oi.cell(r, c).value = None
    for i, (k, c, p) in enumerate(strikes):
        oi.cell(2 + i, 1).value, oi.cell(2 + i, 2).value, oi.cell(2 + i, 3).value = k, c, p

    exp = expiry_date(d["family"], d["label"])
    dte = (exp - d["trade_date"]).days
    oi["F2"] = f"{d['family']} {d['label']} ({d['code']})"
    oi["F3"] = datetime.combine(d["trade_date"], datetime.min.time())
    oi["F4"] = dte
    oi["F5"], oi["F6"] = d["call_total"], d["put_total"]
    if d["price"]:
        calc["F4"] = d["price"]
    else:  # never leave the previous contract's price behind
        calc["F4"] = None
        print("WARNING: no futures settle price from CME - price cleared, fill Max Pain Calc!F4 by hand")

    mp, low, valley = max_pain(strikes)
    calls, puts = sum(s[1] for s in strikes), sum(s[2] for s in strikes)
    coverage = (calls + puts) / (d["call_total"] + d["put_total"])
    measure = (mp - d["price"]) / d["price"] if d["price"] else None
    mid = calc["F9"].value if isinstance(calc["F9"].value, (int, float)) else 0.02
    edge = calc["F10"].value if isinstance(calc["F10"].value, (int, float)) else 0.05
    cur_zone = zone(measure, mid, edge)

    if "Log" not in wb.sheetnames:
        wb.create_sheet("Log").append(LOG_HEAD)
    log = wb["Log"]
    # snapshot before this run's row goes in, so the zone-change alert compares against
    # the last *different* entry, not against a same-day rerun of itself
    prev_zone = zone(log[log.max_row][5].value, mid, edge) if log.max_row > 1 else None
    row = [d["trade_date"], oi["F2"].value, dte, d["price"], mp,
           measure, valley, puts / calls if calls else None, coverage, d["report"],
           datetime.now(), d["price_month"]]
    for r in range(2, log.max_row + 1):  # same day + contract: replace, don't duplicate
        v = log.cell(r, 1).value
        if (v.date() if isinstance(v, datetime) else v) == d["trade_date"] and log.cell(r, 2).value == row[1]:
            log.delete_rows(r)
            break
    log.append(row)
    if cur_zone and prev_zone and cur_zone != prev_zone:
        notify("Max Pain", f"{row[1]}: โซนเปลี่ยนจาก '{prev_zone}' เป็น '{cur_zone}' "
               f"(ห่างราคา {measure:+.2%}, Max Pain {mp})")
    for cell, fmt in zip(log[log.max_row], ["yyyy-mm-dd", None, "0", "0.0", "0", "0.00%", "0", "0.00", "0%", None, "yyyy-mm-dd hh:mm"]):
        if fmt:
            cell.number_format = fmt

    rebuild_charts(wb)  # both charts if this is the combined Max Pain + GEX file
    try:
        save_atomic(wb, path)
    except PermissionError:
        sys.exit(f"Cannot save - close {path.name} in Excel and run again")
    print(f"OK {d['trade_date']} {oi['F2'].value} strikes={len(strikes)} maxpain={mp} "
          f"price={d['price']} ({d['price_month']}) coverage={coverage:.0%} ({d['report']})")
    return d


if __name__ == "__main__":
    data = main()
    try:  # gamma needs the QuikStrike login; a failure must not lose the Max Pain update
        from update_gex import update_gex
        update_gex(json.load(open(CONFIG, encoding="utf-8")), data)
    except SystemExit as e:
        print("GEX skipped:", e)
        notify("MaxPain/GEX", f"GEX skipped - {e}")
    except Exception as e:
        print("GEX skipped:", type(e).__name__, e)
        notify("MaxPain/GEX", f"GEX skipped ({type(e).__name__}) - {e}")
    try:  # record expired contracts' real outcome; CME keeps ~5 days, so it has to run daily
        from outcomes import main as record_outcomes
        record_outcomes()
    except SystemExit as e:
        print("Outcomes skipped:", e)
        notify("Outcomes", f"Outcomes skipped - {e}")
    except Exception as e:
        print("Outcomes skipped:", type(e).__name__, e)
        notify("Outcomes", f"Outcomes skipped ({type(e).__name__}) - {e}")
    try:  # docs/index.html for the mobile/web view; a failure here must not fail the daily run
        from publish_report import main as publish_report
        publish_report()
    except Exception as e:
        print("Report skipped:", type(e).__name__, e)
        notify("Report", f"Report skipped ({type(e).__name__}) - {e}")
    try:  # Google Sheet sync; not set up yet (no service_account.json) just skips quietly
        from sheets_sync import main as sheets_sync
        sheets_sync()
    except Exception as e:
        print("Sheets sync skipped:", type(e).__name__, e)
        notify("Sheets sync", f"Sheets sync skipped ({type(e).__name__}) - {e}")
