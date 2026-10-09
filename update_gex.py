"""Write QuikStrike gamma into the GEX V.4.1 sheets ('Gamma Data' / 'GEX Calc') and append
a daily row to 'GEX Log'. Works on the combined MaxPain_GEX.xlsx or a standalone V.4.1 file."""
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.formula import ArrayFormula

from charts import GEX_LAST as LAST, rebuild_charts
from fetch_gamma import fetch_gamma, qs_code
from notify import notify
from util import save_atomic

HERE = Path(__file__).parent
GEX_LOG = "GEX Log"
MAX_STRIKES = LAST - 1  # Gamma Data rows 2-152
LOG_HEAD = ["วันที่ข้อมูล", "สัญญา", "ราคา", "รวม Call Gamma", "รวม Put Gamma", "NET GEX",
            "สถานะ", "Gamma Flip (ใกล้ราคา)", "Strike gamma สูงสุด", "บันทึกเมื่อ",
            "Call Wall", "Put Wall", "Pin #1", "ความชัดเจน"]
LOG_FMT = ["yyyy-mm-dd", None, "0.0", "0", "0", "0", None, "0", "0", "yyyy-mm-dd hh:mm",
           "0", "0", "0", "0.00"]
# V.4.1 ships with A:F formulas that read blank only from row 67 (rows 2-66 show 0 for an
# empty strike), and with E/F's MEDIAN, H, the edge check J16 and conditional formatting
# stopping at row 62/87. Every row 2-152 gets row 67's blank-safe form, ranges go to 152.
SRC_ROW = 67
J16 = ("=IF(AND(INDEX('Gamma Data'!$B$2:$B$152,COUNT('Gamma Data'!$A$2:$A$152))=0,"
       "INDEX('Gamma Data'!$C$2:$C$152,COUNT('Gamma Data'!$A$2:$A$152))=0),"
       "\"ครบ\",\"ยังไม่ครบ - ขยาย Strikes\")")


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


def conviction(rows):
    """GEX Calc!J18: |sum(call-put)| / sum|call-put|; None when the table has no imbalance."""
    spread = sum(abs(c - p) for _, c, p in rows)
    return abs(sum(c - p for _, c, p in rows)) / spread if spread else None


def gex_status(rows):
    """Same reading as GEX Calc!J5: no mode when |NET| is under 5% of sum|call-put|."""
    conv = conviction(rows)
    if conv is None or conv < 0.05:
        return "ไม่มีโหมด"
    return "Positive GEX (นิ่ง)" if sum(c - p for _, c, p in rows) > 0 else "Negative GEX (วิ่ง)"


def _wall(rows, price, side):
    """J9 (side='call', strikes >= price) / J10 (side='put', strikes <= price): the strike
    with the largest gamma on that side among strikes where it dominates (>0 and >= 1.5x the
    other side); ties go to the lowest strike, like Excel's MATCH."""
    if price is None:
        return None
    if side == "call":
        ok = [(k, c) for k, c, p in rows if k >= price and c > 0 and c >= p * 1.5]
    else:
        ok = [(k, p) for k, c, p in rows if k <= price and p > 0 and p >= c * 1.5]
    if not ok:
        return None
    top = max(v for _, v in ok)
    return next(k for k, v in ok if v == top)


def call_wall(rows, price):
    return _wall(rows, price, "call")


def put_wall(rows, price):
    return _wall(rows, price, "put")


def top_pin(rows, price, window=0.05, min_dist=25):
    """J30 (column AB's pin score, rank 1): gross gamma discounted by distance from price, over
    strikes min_dist..price*window away. The row/1e9 term is the sheet's tie-break (later row,
    i.e. higher strike, wins); rows start at sheet row 2."""
    if not price:
        return None
    best = None
    for i, (k, c, p) in enumerate(rows):
        gross, dist = abs(c) + abs(p), abs(k - price)
        if not gross or dist > price * window or dist < min_dist:
            continue
        score = gross / max(dist / (price * 0.005), 1) + (2 + i) / 1e9
        if best is None or score > best[0]:
            best = (score, k)
    return best[1] if best else None


def levels(rows, price, gc=None):
    """[Call Wall, Put Wall, Pin #1, conviction] as GEX Log stores them; the pin's search
    window / minimum distance come from the sheet's own inputs J26 / J28 when available."""
    window = gc["J26"].value if gc is not None and isinstance(gc["J26"].value, (int, float)) else 0.05
    min_dist = gc["J28"].value if gc is not None and isinstance(gc["J28"].value, (int, float)) else 25
    return [call_wall(rows, price), put_wall(rows, price), top_pin(rows, price, window, min_dist),
            conviction(rows)]


def gex_log(wb):
    """The GEX Log sheet, created if missing; an older, shorter header is extended in place
    (earlier rows simply stay blank in the new columns)."""
    if GEX_LOG not in wb.sheetnames:
        wb.create_sheet(GEX_LOG).append(LOG_HEAD)
    log = wb[GEX_LOG]
    for i, h in enumerate(LOG_HEAD, start=1):
        if log.cell(1, i).value is None:
            log.cell(1, i).value = h
    return log


def format_last_row(log):
    for cell, fmt in zip(log[log.max_row], LOG_FMT):
        if fmt:
            cell.number_format = fmt


def _mode(status):
    """'+', '-' or None, so old Log wording ('แกว่งแรง') and 'ไม่มีโหมด' compare correctly."""
    s = str(status or "")
    return "+" if s.startswith("Positive") else "-" if s.startswith("Negative") else None


def extend_layout(wb):
    """Idempotent: widen GEX V.4.1's GEX Calc to rows 2-152 (see SRC_ROW note)."""
    gc = wb["GEX Calc"]
    if str(gc[f"A{LAST}"].value or "").startswith("=IF("):
        return
    src = {col: gc[f"{col}{SRC_ROW}"].value.replace("$C$62", f"$C${LAST}") for col in "ABCDEF"}
    h_src = gc["H62"].value
    for r in range(2, LAST + 1):
        for col in "ABCDEF":
            cell = gc[f"{col}{r}"]
            cell.value = Translator(src[col], origin=f"{col}{SRC_ROW}").translate_formula(f"{col}{r}")
            if r > 87:
                cell._style = copy(gc[f"{col}{SRC_ROW}"]._style)
        if r > 62:
            gc[f"H{r}"] = Translator(h_src, origin="H62").translate_formula(f"H{r}")
    gc["J16"] = J16
    old, gc.conditional_formatting = gc.conditional_formatting, ConditionalFormattingList()
    for rng, rules in old._cf_rules.items():
        for rule in rules:
            gc.conditional_formatting.add(str(rng.sqref).replace("87", str(LAST)), rule)


def fix_zone_median(gc):
    """Idempotent: GEX Calc's strength/zone columns (E/F) compare each strike with a median
    gross. Over every row that median was 0 once most strikes carried no gamma (2026-10-07)
    and every row read จุดหนืด; J12 now takes the median over strikes with gamma within ±J26
    of the price (Pins' own search band), and E/F read J12."""
    j12 = gc["J12"].value
    if isinstance(j12, ArrayFormula) and "$J$26" in j12.text:
        return
    a, b, c = ("'Gamma Data'!$%s$2:$%s$%d" % (x, x, LAST) for x in "ABC")
    gross = f"(ABS({b})+ABS({c}))"
    gc["J12"] = ArrayFormula("J12", f"=IFERROR(MEDIAN(IF(ISNUMBER({a})*({gross}>0)*"
                                    f"(ABS({a}-$J$6)<=$J$6*$J$26),{gross})),0)")
    gc["I12"] = "ค่ากลาง Gross (strike ที่มี gamma, ±J26 ของราคา)"
    for r in range(2, LAST + 1):
        gc[f"E{r}"] = f'=IF(OR(A{r}="",$J$12=0),"",C{r}/$J$12)'
        gc[f"F{r}"] = f'=IF(A{r}="","",IF(C{r}>=2*$J$12,"จุดหนืด",IF(C{r}<=0.5*$J$12,"air pocket","")))'


def ref_price(cfg, d):
    """Price the GEX levels are measured from: the live (10-min delayed) futures quote at run
    time by default, so walls/pins/distances match the chart the user trades from (BlackBull
    gold futures on TradingView sat ~$26 off the previous settle, 2026-10-07); the settle if
    config says "gex_price_source": "settle" or there's no live quote."""
    if cfg.get("gex_price_source", "live") != "settle" and d.get("live_price"):
        return d["live_price"]
    return d["price"]


def update_gex(cfg, d):
    code = qs_code(d["family"], d["label"])
    g = fetch_gamma(cfg, code, d["trade_date"])
    rows = select_rows(g)
    if len(rows) > MAX_STRIKES:
        raise RuntimeError(f"{len(rows)} strikes exceeds GEX Calc range ({MAX_STRIKES})")

    path = HERE / cfg["gex_workbook"]
    wb = openpyxl.load_workbook(path)
    extend_layout(wb)
    gd, gc = wb["Gamma Data"], wb["GEX Calc"]
    for r in range(2, LAST + 1):
        for col in (1, 2, 3):
            gd.cell(r, col).value = None
    for i, (k, c, p) in enumerate(rows):
        gd.cell(2 + i, 1).value, gd.cell(2 + i, 2).value, gd.cell(2 + i, 3).value = k, c, p
    price = ref_price(cfg, d)
    fix_zone_median(gc)
    if cfg.get("gex_fill_price", True):
        gc["J6"] = price  # None rather than leaving a previous contract's price behind
    gc["I1"] = "ข้อมูลวันที่ / สัญญา"
    gc["J1"] = f"{d['trade_date']} {code} ({d['label']})"

    calls, puts = sum(r[1] for r in rows), sum(r[2] for r in rows)
    net = calls - puts
    status = gex_status(rows)
    flip = gamma_flip(rows, price)
    peak = max(rows, key=lambda r: r[1] + r[2])[0]

    log = gex_log(wb)
    # snapshot before this run's row goes in, so the mode-change alert compares against
    # the last *different* entry, not against a same-day rerun of itself
    prev_status, prev_net = (log[log.max_row][6].value, log[log.max_row][5].value) if log.max_row > 1 else (None, None)
    lv = levels(rows, price, gc)
    row = [d["trade_date"], code, price, calls, puts, net, status, flip, peak, datetime.now(), *lv]
    for r in range(2, log.max_row + 1):
        v = log.cell(r, 1).value
        if (v.date() if isinstance(v, datetime) else v) == d["trade_date"] and log.cell(r, 2).value == code:
            log.delete_rows(r)
            break
    log.append(row)
    if _mode(prev_status) and _mode(status) and _mode(prev_status) != _mode(status):
        notify("GEX", f"{code}: โหมด GEX พลิก {prev_net:+.0f} -> {net:+.0f} ({status})")
    format_last_row(log)

    rebuild_charts(wb)
    out = {"code": code, "label": d["label"], "trade_date": d["trade_date"], "report": d.get("report"),
           "price": price, "settle": d["price"], "live_time": d.get("live_time") if price != d["price"] else None,
           "rows": rows, "calls": calls, "puts": puts, "net": net, "status": status, "flip": flip,
           "peak": peak, "call_wall": lv[0], "put_wall": lv[1], "pin": lv[2], "conviction": lv[3],
           "path": path, "saved": True}
    try:
        save_atomic(wb, path)
    except PermissionError:
        # the phone report still goes out from these numbers; only the workbook is stale
        print(f"Cannot save - close {path.name} in Excel and run again")
        notify("GEX", f"บันทึก {path.name} ไม่ได้ - ไฟล์เปิดค้างใน Excel (รายงานยังส่ง แต่ไม่มี Pins จาก Excel)")
        out["saved"] = False
        return out
    print(f"GEX OK {d['trade_date']} {code} strikes={len(rows)} call={calls:.0f} put={puts:.0f} "
          f"net={net:+.0f} {status} flip={flip} peak={peak} "
          f"walls={lv[1]}/{lv[0]} pin={lv[2]}")
    return out


AGG_LOG = "Log รวม"
AGG_HEAD = ["วันที่ข้อมูล", "สัญญาที่รวม", "ราคา", "รวม Call Gamma", "รวม Put Gamma", "NET GEX",
            "โหมด", "ความชัดเจน", "Call Wall", "Put Wall", "Gamma Flip", "Strike gamma สูงสุด", "บันทึกเมื่อ"]
AGG_FMT = ["yyyy-mm-dd", None, "0.0", "0", "0", "0", None, "0.000", "0", "0", "0", "0", "yyyy-mm-dd hh:mm"]


def aggregate(per_code):
    """{code: {strike: (call, put)}} -> [(strike, call, put)] summed over expirations."""
    tot = {}
    for g in per_code.values():
        for k, (c, p) in g.items():
            a = tot.setdefault(k, [0.0, 0.0])
            a[0] += c
            a[1] += p
    return [(k, c, p) for k, (c, p) in sorted(tot.items())]


def agg_levels(rows, price):
    """The workbook's rules (call_wall/put_wall/gex_status/conviction above) applied to any
    gamma table - here every expiration summed - plus peak and Gamma Flip."""
    calls, puts = sum(r[1] for r in rows), sum(r[2] for r in rows)
    return {"calls": calls, "puts": puts, "net": calls - puts, "mode": gex_status(rows),
            "conviction": conviction(rows) or 0.0, "call_wall": call_wall(rows, price),
            "put_wall": put_wall(rows, price),
            "peak": max(rows, key=lambda r: r[1] + r[2])[0] if rows else None,
            "flip": gamma_flip(rows, price)}


def update_gex_all(cfg, d):
    """GEX over every expiration QuikStrike lists by default (nearest weekly + monthlies):
    its own 'Log รวม' sheet (Gamma Data / GEX Calc stay single-contract) and the phone card."""
    from fetch_gamma import fetch_gamma_all
    per_code = fetch_gamma_all(cfg, d["trade_date"])
    rows = aggregate(per_code)
    price = ref_price(cfg, d)
    out = {"trade_date": d["trade_date"], "codes": list(per_code), "price": price, "rows": rows,
           **agg_levels(rows, price)}
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
    for cell, fmt in zip(log[log.max_row], AGG_FMT):
        if fmt:
            cell.number_format = fmt
    rebuild_charts(wb)
    try:
        save_atomic(wb, path)
    except PermissionError:
        print(f"Cannot save - close {path.name} in Excel and run again")
    print(f"GEX ALL OK {d['trade_date']} {len(per_code)} expirations net={out['net']:+.0f} "
          f"{out['mode']} cw={out['call_wall']} pw={out['put_wall']} flip={out['flip']}")
    return out
