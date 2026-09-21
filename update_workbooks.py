"""Fetch today's CME data and write it into the Max Pain workbook + Log sheet."""
import json
import os
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
from copy import copy
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formula.translate import Translator

from fetch_cme import expiry_date, fetch

HERE = Path(__file__).parent
CONFIG = Path(os.environ.get("MAXPAIN_CONFIG", HERE / "config.json"))  # override to work on cloned files
MAX_ROWS = 200  # OI Data rows 2-201, the range covered by the Max Pain Calc formulas
OLD_LAST, NEW_LAST = 102, 201
LOG_HEAD = ["วันที่ข้อมูล", "สัญญา", "DTE", "ราคา", "Max Pain", "ห่างจากราคา %",
            "ก้นหุบ (จำนวน strike)", "P/C Ratio", "ครอบคลุม CME %", "รายงาน", "บันทึกเมื่อ"]


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


def rebuild_chart(calc, n):
    """Replace the old fixed-range chart with one sized to today's strike count."""
    calc._charts = []
    ch = LineChart()
    ch.title, ch.legend = "PAIN รวม ตาม Strike", None
    ch.height, ch.width = 7.5, 26
    ch.add_data(Reference(calc, min_col=2, min_row=1, max_row=1 + n), titles_from_data=True)
    ch.set_categories(Reference(calc, min_col=1, min_row=2, max_row=1 + n))
    calc.add_chart(ch, "H2")


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

    mp, low, valley = max_pain(strikes)
    calls, puts = sum(s[1] for s in strikes), sum(s[2] for s in strikes)
    coverage = (calls + puts) / (d["call_total"] + d["put_total"])

    if "Log" not in wb.sheetnames:
        wb.create_sheet("Log").append(LOG_HEAD)
    log = wb["Log"]
    row = [d["trade_date"], oi["F2"].value, dte, d["price"], mp,
           (mp - d["price"]) / d["price"] if d["price"] else None, valley,
           puts / calls if calls else None, coverage, d["report"], datetime.now()]
    for r in range(2, log.max_row + 1):  # same day + contract: replace, don't duplicate
        v = log.cell(r, 1).value
        if (v.date() if isinstance(v, datetime) else v) == d["trade_date"] and log.cell(r, 2).value == row[1]:
            log.delete_rows(r)
            break
    log.append(row)
    for cell, fmt in zip(log[log.max_row], ["yyyy-mm-dd", None, "0", "0.0", "0", "0.00%", "0", "0.00", "0%", None, "yyyy-mm-dd hh:mm"]):
        if fmt:
            cell.number_format = fmt

    rebuild_chart(calc, len(strikes))
    try:
        wb.save(path)
    except PermissionError:
        sys.exit(f"Cannot save - close {path.name} in Excel and run again")
    print(f"OK {d['trade_date']} {oi['F2'].value} strikes={len(strikes)} maxpain={mp} "
          f"price={d['price']} coverage={coverage:.0%} ({d['report']})")
    return d


if __name__ == "__main__":
    data = main()
    try:  # gamma needs the QuikStrike login; a failure must not lose the Max Pain update
        from update_gex import update_gex
        update_gex(json.load(open(CONFIG, encoding="utf-8")), data)
    except SystemExit as e:
        print("GEX skipped:", e)
    except Exception as e:
        print("GEX skipped:", type(e).__name__, e)
