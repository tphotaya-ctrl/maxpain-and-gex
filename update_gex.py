"""Write QuikStrike gamma into the GEX workbook and append a daily Log row."""
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.formatting import ConditionalFormattingList

from fetch_gamma import fetch_gamma, qs_code
from notify import notify
from util import save_atomic

HERE = Path(__file__).parent
LAST = 201  # formulas in GEX Calc cover data rows 2-201
LOG_HEAD = ["วันที่ข้อมูล", "สัญญา", "ราคา", "รวม Call Gamma", "รวม Put Gamma", "NET GEX",
            "สถานะ", "Gamma Flip (ใกล้ราคา)", "Strike gamma สูงสุด", "บันทึกเมื่อ"]


def select_rows(g):
    """All strikes with gamma, plus one zero strike each side so the edge checks read 'ครบ'."""
    ks = sorted(g)
    nz = [i for i, k in enumerate(ks) if g[k][0] or g[k][1]]
    if not nz:
        raise RuntimeError("gamma table is all zero")
    lo, hi = max(nz[0] - 1, 0), min(nz[-1] + 1, len(ks) - 1)
    return [(k, g[k][0], g[k][1]) for k in ks[lo:hi + 1]]


def gamma_flip(rows, price):
    """Strike where cumulative net gamma changes sign, nearest to price."""
    cum, prev, flips = 0.0, None, []
    for k, c, p in rows:
        cum += c - p
        if prev is not None and (prev < 0 <= cum or prev > 0 >= cum) and prev != cum:
            flips.append(k)
        if cum != 0:
            prev = cum
    if not flips:
        return None
    return min(flips, key=lambda k: abs(k - price)) if price else flips[0]


def rebuild_layout(wb, n):
    gd, gc = wb["Gamma Data"], wb["GEX Calc"]
    for r in range(2, LAST + 1):
        for col in "ABC":
            gd[f"{col}{r}"]._style = copy(gd[f"{col}2"]._style)
        gc[f"A{r}"] = f"=IF('Gamma Data'!A{r}=\"\",\"\",'Gamma Data'!A{r})"
        gc[f"B{r}"] = f"=IF(A{r}=\"\",\"\",'Gamma Data'!B{r}-'Gamma Data'!C{r})"
        gc[f"C{r}"] = f"=IF(A{r}=\"\",\"\",B{r})" if r == 2 else f"=IF(A{r}=\"\",\"\",N(C{r-1})+B{r})"
        for col in "ABC":
            gc[f"{col}{r}"]._style = copy(gc[f"{col}{min(r, 62)}"]._style)
    gc["F2"] = f"=SUM('Gamma Data'!B2:B{LAST})"
    gc["F3"] = f"=SUM('Gamma Data'!C2:C{LAST})"
    gc["F6"] = "=IF(AND('Gamma Data'!B2=0,'Gamma Data'!C2=0),\"ครบ\",\"ยังไม่ครบ - ขยาย Strikes เพิ่ม\")"
    last = f"INDEX('Gamma Data'!{{c}}2:{{c}}{LAST},COUNT('Gamma Data'!A2:A{LAST}))"
    gc["F7"] = (f"=IF(AND({last.format(c='B')}=0,{last.format(c='C')}=0),\"ครบ\","
                "\"ยังไม่ครบ - ขยาย Strikes เพิ่ม\")")
    old, gc.conditional_formatting = gc.conditional_formatting, ConditionalFormattingList()
    for rng, rules in old._cf_rules.items():
        for rule in rules:
            gc.conditional_formatting.add(str(rng.sqref).replace("62", str(LAST)), rule)
    gc._charts = []
    ch = LineChart()
    ch.title, ch.legend = "CUMULATIVE NET GEX", None
    ch.height, ch.width = 7.5, 22
    ch.add_data(Reference(gc, min_col=3, min_row=1, max_row=1 + n), titles_from_data=True)
    ch.set_categories(Reference(gc, min_col=1, min_row=2, max_row=1 + n))
    gc.add_chart(ch, "I2")


def update_gex(cfg, d):
    code = qs_code(d["family"], d["label"])
    g = fetch_gamma(cfg, code, d["trade_date"])
    rows = select_rows(g)

    path = HERE / cfg["gex_workbook"]
    wb = openpyxl.load_workbook(path)
    gd, gc = wb["Gamma Data"], wb["GEX Calc"]
    rebuild_layout(wb, len(rows))
    for r in range(2, LAST + 1):
        for col in (1, 2, 3):
            gd.cell(r, col).value = None
    for i, (k, c, p) in enumerate(rows):
        gd.cell(2 + i, 1).value, gd.cell(2 + i, 2).value, gd.cell(2 + i, 3).value = k, c, p

    calls, puts = sum(r[1] for r in rows), sum(r[2] for r in rows)
    net = calls - puts
    status = "Positive GEX (นิ่ง)" if net > 0 else "Negative GEX (แกว่งแรง)"
    flip = gamma_flip(rows, d["price"])
    peak = max(rows, key=lambda r: r[1] + r[2])[0]

    if "Log" not in wb.sheetnames:
        wb.create_sheet("Log").append(LOG_HEAD)
    log = wb["Log"]
    # snapshot before this run's row goes in, so the sign-flip alert compares against the
    # last *different* entry, not against a same-day rerun of itself
    prev_net = log[log.max_row][5].value if log.max_row > 1 else None
    row = [d["trade_date"], code, d["price"], calls, puts, net, status, flip, peak, datetime.now()]
    for r in range(2, log.max_row + 1):
        v = log.cell(r, 1).value
        if (v.date() if isinstance(v, datetime) else v) == d["trade_date"] and log.cell(r, 2).value == code:
            log.delete_rows(r)
            break
    log.append(row)
    if isinstance(prev_net, (int, float)) and prev_net != 0 and net != 0 and (prev_net > 0) != (net > 0):
        notify("GEX", f"{code}: NET GEX พลิกเครื่องหมาย {prev_net:+.0f} -> {net:+.0f} ({status})")
    for cell, fmt in zip(log[log.max_row], ["yyyy-mm-dd", None, "0.0", "0", "0", "0", None, "0", "0", "yyyy-mm-dd hh:mm"]):
        if fmt:
            cell.number_format = fmt
    try:
        save_atomic(wb, path)
    except PermissionError:
        sys.exit(f"Cannot save - close {path.name} in Excel and run again")
    print(f"GEX OK {d['trade_date']} {code} strikes={len(rows)} call={calls:.0f} put={puts:.0f} "
          f"net={net:+.0f} flip={flip} peak={peak}")
