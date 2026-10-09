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


def mode_label(rows, net):
    """GEX Calc!J5's rule: conviction = |NET| / sum|call-put|; under 0.05 is 'no mode'."""
    den = sum(abs(c - p) for _, c, p in rows)
    conv = abs(net) / den if den else 0.0
    if conv < 0.05:
        return "ไม่มีโหมด", conv
    return ("Positive GEX (นิ่ง)" if net > 0 else "Negative GEX (วิ่ง)"), conv


def aggregate(per_code):
    """{code: {strike: (call, put)}} -> [(strike, call, put)] summed over expirations."""
    tot = {}
    for g in per_code.values():
        for k, (c, p) in g.items():
            a = tot.setdefault(k, [0.0, 0.0])
            a[0] += c
            a[1] += p
    return [(k, c, p) for k, (c, p) in sorted(tot.items())]


def levels(rows, price):
    """The workbook's level rules applied to any gamma table (here: all expirations summed):
    Call Wall = largest call gamma at/above price where call >= 1.5x put, Put Wall the mirror
    below, peak = largest gross; plus totals, mode/conviction and Gamma Flip."""
    calls, puts = sum(r[1] for r in rows), sum(r[2] for r in rows)
    cw = [r for r in rows if r[0] >= price and r[1] > 0 and r[1] >= r[2] * 1.5]
    pw = [r for r in rows if r[0] <= price and r[2] > 0 and r[2] >= r[1] * 1.5]
    mode, conv = mode_label(rows, calls - puts)
    return {"calls": calls, "puts": puts, "net": calls - puts, "mode": mode, "conviction": conv,
            "call_wall": max(cw, key=lambda r: r[1])[0] if cw else None,
            "put_wall": max(pw, key=lambda r: r[2])[0] if pw else None,
            "peak": max(rows, key=lambda r: r[1] + r[2])[0] if rows else None,
            "flip": gamma_flip(rows, price)}


AGG_LOG = "Log รวม"
AGG_HEAD = ["วันที่ข้อมูล", "สัญญาที่รวม", "ราคา", "รวม Call Gamma", "รวม Put Gamma", "NET GEX",
            "โหมด", "Conviction", "Call Wall", "Put Wall", "Gamma Flip", "Strike gamma สูงสุด", "บันทึกเมื่อ"]


def update_gex_all(cfg, d):
    """GEX over every expiration QuikStrike lists by default - its own 'Log รวม' sheet in the
    GEX workbook (Gamma Data / GEX Calc stay single-contract) and the phone card section."""
    from fetch_gamma import fetch_gamma_all
    per_code = fetch_gamma_all(cfg, d["trade_date"])
    rows = aggregate(per_code)
    price = ref_price(cfg, d)
    out = {"trade_date": d["trade_date"], "codes": list(per_code), "price": price, "rows": rows,
           **levels(rows, price)}
    path = HERE / cfg["gex_workbook"]
    wb = openpyxl.load_workbook(path)
    if AGG_LOG not in wb.sheetnames:
        wb.create_sheet(AGG_LOG).append(AGG_HEAD)
    log = wb[AGG_LOG]
    for r in range(2, log.max_row + 1):  # same trade date: replace, don't duplicate
        v = log.cell(r, 1).value
        if (v.date() if isinstance(v, datetime) else v) == d["trade_date"]:
            log.delete_rows(r)
            break
    log.append([d["trade_date"], ", ".join(per_code), price, out["calls"], out["puts"], out["net"],
                out["mode"], round(out["conviction"], 3), out["call_wall"], out["put_wall"],
                out["flip"], out["peak"], datetime.now()])
    for cell, fmt in zip(log[log.max_row], ["yyyy-mm-dd", None, "0.0", "0", "0", "0", None, "0.000",
                                            "0", "0", "0", "0", "yyyy-mm-dd hh:mm"]):
        if fmt:
            cell.number_format = fmt
    try:
        save_atomic(wb, path)
    except PermissionError:
        print(f"Cannot save - close {path.name} in Excel and run again")
    print(f"GEX ALL OK {d['trade_date']} {len(per_code)} expirations net={out['net']:+.0f} "
          f"{out['mode']} cw={out['call_wall']} pw={out['put_wall']} flip={out['flip']}")
    return out


def ref_price(cfg, d):
    """Price the GEX levels are measured from: the live (10-min delayed) futures quote at run
    time by default, so walls/pins/distances match the chart the user trades from; the
    previous session's settle if config says "gex_price_source": "settle" or no live quote."""
    if cfg.get("gex_price_source", "live") != "settle" and d.get("live_price"):
        return d["live_price"]
    return d["price"]


def fill_calc_header(gc, cfg, d, code, n):
    """The few GEX Calc cells the script owns: data-date label (I1/J1), price (J6),
    and the chart's row range. Everything else in GEX Calc is the workbook's own formulas."""
    gc["I1"] = "ข้อมูลวันที่ / สัญญา"
    gc["J1"] = f"{d['trade_date']} {code} ({d['label']})"
    # J6 drives Call/Put Wall and Pins; a stale hand-typed price silently skews them
    price = ref_price(cfg, d)
    if cfg.get("gex_fill_price", True) and price:
        gc["J6"] = price
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
    price = ref_price(cfg, d)
    flip = gamma_flip(rows, price)
    peak = max(rows, key=lambda r: r[1] + r[2])[0]

    if "Log" not in wb.sheetnames:
        wb.create_sheet("Log").append(LOG_HEAD)
    log = wb["Log"]
    # snapshot before this run's row goes in, so the sign-flip alert compares against the
    # last *different* entry, not against a same-day rerun of itself
    prev_net = log[log.max_row][5].value if log.max_row > 1 else None
    row = [d["trade_date"], code, price, calls, puts, net, status, flip, peak, datetime.now()]
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
    out = {"code": code, "label": d["label"], "trade_date": d["trade_date"], "report": d.get("report"),
           "price": price, "settle": d["price"], "live_time": d.get("live_time")
           if price != d["price"] else None, "rows": rows, "calls": calls, "puts": puts, "net": net,
           "flip": flip, "peak": peak, "path": path, "saved": True}
    try:
        save_atomic(wb, path)
    except PermissionError:
        # the phone report still goes out from these numbers; only the workbook is stale
        print(f"Cannot save - close {path.name} in Excel and run again")
        notify("GEX", f"บันทึก {path.name} ไม่ได้ - ไฟล์เปิดค้างใน Excel (รายงานยังส่ง แต่ไม่มี Wall/Pins)")
        out["saved"] = False
        return out
    print(f"GEX OK {d['trade_date']} {code} strikes={len(rows)} call={calls:.0f} put={puts:.0f} "
          f"net={net:+.0f} flip={flip} peak={peak}")
    return out
