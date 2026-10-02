"""Write QuikStrike gamma into the GEX workbook and append a daily Log row."""
import re
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from notify import notify
from fetch_gamma import fetch_gamma, qs_code
from util import save_atomic

HERE = Path(__file__).parent
# GEX Calc (V4.1 layout) already has self-contained, blank-safe formulas fixed to this
# range ('Gamma Data'!$x$2:$x$152 everywhere) - Python only ever fills raw Gamma Data
# values, never rewrites GEX Calc's formulas/conditional formatting/chart.
LAST = 152
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


def fill_calc_header(gc, cfg, d, code, n):
    """The few GEX Calc cells the script owns: data-date label (I1/J1), price (J6),
    and the chart's row range. Everything else in GEX Calc is the workbook's own formulas."""
    gc["I1"] = "ข้อมูลวันที่ / สัญญา"
    gc["J1"] = f"{d['trade_date']} {code} ({d['label']})"
    # J6 drives Call/Put Wall and Pins; a stale hand-typed price silently skews them
    if cfg.get("gex_fill_price", True) and d.get("price"):
        gc["J6"] = d["price"]
    for ch in gc._charts:
        for s in ch.series:
            for ref in (s.val and s.val.numRef, s.cat and (s.cat.numRef or s.cat.strRef)):
                if ref is not None and ref.f:
                    ref.f = re.sub(r"\$(\d+)$", f"${1 + n}", ref.f)


def update_gex(cfg, d):
    code = qs_code(d["family"], d["label"])
    g = fetch_gamma(cfg, code, d["trade_date"])
    rows = select_rows(g)

    path = HERE / cfg["gex_workbook"]
    wb = openpyxl.load_workbook(path)
    gd = wb["Gamma Data"]
    # Gamma Data holds raw values only (no formulas) - GEX Calc's own formulas already
    # cover A2:A{LAST} and self-guard against blanks, so we never touch GEX Calc here.
    for r in range(2, LAST + 1):
        for col_letter, col in zip("ABC", (1, 2, 3)):
            gd.cell(r, col)._style = copy(gd[f"{col_letter}2"]._style)
            gd.cell(r, col).value = None
    for i, (k, c, p) in enumerate(rows):
        gd.cell(2 + i, 1).value, gd.cell(2 + i, 2).value, gd.cell(2 + i, 3).value = k, c, p
    fill_calc_header(wb["GEX Calc"], cfg, d, code, len(rows))

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
