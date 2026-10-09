"""Independent check of the daily update.

Re-fetches the workbook's contract fresh from CME (and QuikStrike for gamma)
and compares it with what is stored in the Excel files. Max Pain / GEX results
are recomputed in plain Python and compared with what Excel itself calculates.

    python verify.py              # everything
    python verify.py --no-gamma   # skip the QuikStrike (login) part
    python verify.py --no-excel   # skip reading Excel's calculated cells
Exit code 1 if any check FAILs.
"""
import argparse
import base64
import json
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

from fetch_cme import expiry_date, fetch
from fetch_gamma import fetch_gamma, qs_code
from charts import GEX_LAST
from update_gex import GEX_LOG, gamma_flip, select_rows

HERE = Path(__file__).parent
CONFIG = Path(os.environ.get("MAXPAIN_CONFIG", HERE / "config.json"))  # override to work on cloned files
results = []


def check(name, ok, detail="", warn=False):
    status = "PASS" if ok else ("WARN" if warn else "FAIL")
    results.append(status)
    print(f"[{status}] {name}" + (f" - {detail}" if detail else ""))


def skip(name, why):
    results.append("SKIP")
    print(f"[SKIP] {name} - {why}")


def excel_cells(path, wanted):
    """Open read-only in Excel, recalc, return {'Sheet!A1': text}; None if Excel is unavailable."""
    quoted = str(path).replace("'", "''")
    # $x.Quit() alone leaves the COM-launched EXCEL.EXE running (4 left behind, ~230MB each,
    # 2026-10-07). Note which EXCEL.EXE this call started right after starting it, and kill
    # exactly that one in `finally` - so an error midway can't leak it either, and a
    # concurrent run's Excel (or the user's own) is never touched.
    lines = ["$ErrorActionPreference='Stop'",
             "$pre=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)",
             "$x=$null; $wb=$null; $mine=@()",
             "try {",
             "$x=New-Object -ComObject Excel.Application",
             "$mine=@(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $pre -notcontains $_.Id "
             "-and -not $_.MainWindowTitle } | ForEach-Object Id)",
             "$x.Visible=$false; $x.DisplayAlerts=$false",
             # Open() sometimes hands back $null while a freshly started Excel is still busy
             # (about 1 run in 6, 2026-10-07) - retry instead of failing the whole read
             "for ($i=0; $i -lt 8 -and -not ($wb -and $wb.Worksheets.Count); $i++) {",
             f"  try {{ $wb=$x.Workbooks.Open('{quoted}',0,$true) }} catch {{ $wb=$null }}",
             "  if (-not ($wb -and $wb.Worksheets.Count)) { Start-Sleep -Milliseconds 750 } }",
             "if (-not $wb) { throw 'Excel did not open the workbook' }",
             "$x.CalculateFull()", "$o=@{}"]
    for sheet, cells in wanted.items():
        for c in cells:
            lines.append(f"$o['{sheet}!{c}']=[string]$wb.Worksheets.Item('{sheet}').Range('{c}').Value2")
    lines += ["[Console]::OutputEncoding=[Text.Encoding]::UTF8; $o|ConvertTo-Json -Compress",
              "} finally {",
              "try { if ($wb) { $wb.Close($false) }; if ($x) { $x.Quit() } } catch {}",
              "foreach ($c in @($wb, $x)) { if ($c) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($c) } }",
              "$mine | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }",
              "}"]
    enc = base64.b64encode("\n".join(lines).encode("utf-16-le")).decode()
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-EncodedCommand", enc],
                           capture_output=True, timeout=180)
        return json.loads(r.stdout.decode("utf-8")) if r.returncode == 0 else None
    except Exception:
        return None


def max_pain_bruteforce(rows):
    """Independent of the sheet formula: total payout to option holders if settled at each strike."""
    pain = {k: sum(max(0, k - s) * c + max(0, s - k) * p for s, c, p in rows) for k, _, _ in rows}
    low = min(pain.values())
    return min(k for k, v in pain.items() if v == low), low


def verify_max_pain(cfg, use_excel):
    print("== Max Pain workbook ==")
    path = HERE / cfg["workbook"]
    wb = openpyxl.load_workbook(path)
    oi, calc = wb["OI Data"], wb["Max Pain Calc"]
    label_cell, dt, dte, ctot, ptot = (oi[f"F{i}"].value for i in range(2, 7))
    m = re.match(r"(\w+) (Week \d - \w{3} \d{4}) \((\w+)\)", str(label_cell))
    if not m or not isinstance(dt, datetime):
        skip("workbook header", f"cannot parse OI Data F2/F3: {label_cell!r} {dt!r}")
        return None
    family, label, _ = m.groups()
    rows = [(int(r[0].value), r[1].value or 0, r[2].value or 0)
            for r in oi.iter_rows(min_row=2, max_row=201) if isinstance(r[0].value, (int, float))]
    print(f"workbook: {label_cell}, trade date {dt.date()}, {len(rows)} strikes")

    fresh = fetch({**cfg, "family": family, "target": label, "trade_date": dt.strftime("%Y%m%d"),
                   "strike_min": 0, "strike_max": 10**9})
    check("same contract", fresh["label"] == label and fresh["family"] == family,
          f"CME returned {fresh['family']} {fresh['label']}")
    print(f"       CME report type now: {fresh['report']}")

    log = wb["Log"]
    last = [c.value for c in log[log.max_row]]
    # a difference is only excusable if CME's report changed since we stored it (PRELIMINARY -> FINAL)
    report_changed = last[9] != fresh["report"]
    fr = {k: (c, p) for k, c, p in fresh["strikes"]}
    stored = {r[0] for r in rows}
    lo, hi = min(stored), max(stored)
    diffs = [(k, (c, p), fr.get(k, (0, 0))) for k, c, p in rows if (c, p) != fr.get(k, (0, 0))]
    missing = [k for k, v in fr.items() if lo <= k <= hi and k not in stored and v != (0, 0)]
    check("OI per strike matches CME", not diffs and not missing,
          f"{len(rows) - len(diffs)}/{len(rows)} equal"
          + (f"; stored {last[9]}, CME now {fresh['report']} - re-run update_workbooks.py to refresh"
             if report_changed else "")
          + (f"; first diffs {diffs[:3]}" if diffs else "")
          + (f"; strikes missing from sheet {missing[:5]}" if missing else ""),
          warn=report_changed)
    check("CME totals (F5/F6)", (ctot, ptot) == (fresh["call_total"], fresh["put_total"]),
          f"sheet {ctot}/{ptot} vs CME {fresh['call_total']}/{fresh['put_total']}")
    sc, sp = sum(r[1] for r in rows), sum(r[2] for r in rows)
    check("strike sums = CME totals", (sc, sp) == (ctot, ptot),
          f"coverage {(sc + sp) / (ctot + ptot):.1%}")
    exp = expiry_date(family, label)
    calc_dte = (exp - dt.date()).days
    check("DTE", dte == calc_dte, f"sheet {dte}, expiry {exp}, computed {calc_dte}")
    price = calc["F4"].value
    check("price vs CME settle", fresh["price"] is not None and abs(price - fresh["price"]) < 0.05,
          f"sheet {price} vs CME {fresh['price']}")

    mp, low = max_pain_bruteforce(rows)
    print(f"       independent Max Pain = {mp} (pain {low})")
    check("Log row matches", last[1] == label_cell and last[4] == mp, f"Log: {last[1]} max pain {last[4]}")
    if use_excel:
        cells = excel_cells(path, {"Max Pain Calc": ["F2", "F3", "F28"]})
        if cells is None:
            skip("Excel-calculated Max Pain", "Excel unavailable or file could not be opened")
        else:
            xl_mp, xl_low = float(cells["Max Pain Calc!F2"]), float(cells["Max Pain Calc!F3"])
            check("Excel Max Pain = independent", xl_mp == mp and xl_low == low,
                  f"Excel {xl_mp:.0f} / {xl_low:.0f} vs Python {mp} / {low}")
            cov = float(cells["Max Pain Calc!F28"])
            check("Excel coverage >= 98%", cov >= 0.98, f"{cov:.1%}")
    return fresh, family, label, dt


def verify_gex(cfg, fresh, family, label, dt, use_excel):
    print("\n== GEX workbook ==")
    path = HERE / cfg["gex_workbook"]
    wb = openpyxl.load_workbook(path)
    rows = [(int(r[0].value), r[1].value or 0, r[2].value or 0)
            for r in wb["Gamma Data"].iter_rows(min_row=2, max_row=GEX_LAST) if isinstance(r[0].value, (int, float))]
    code = qs_code(family, label)
    log = wb[GEX_LOG]
    last = [c.value for c in log[log.max_row]]
    print(f"workbook: {last[1]}, trade date {last[0].date()}, {len(rows)} strikes")
    check("same contract/date as Max Pain file", last[1] == code and last[0].date() == dt.date(),
          f"GEX Log {last[1]} {last[0].date()} vs {code} {dt.date()}")

    g = fetch_gamma(cfg, code, dt.date())
    stored = {r[0] for r in rows}
    diffs = [(k, (c, p), g.get(k)) for k, c, p in rows if g.get(k) != (c, p)]
    missing = [k for k, v in g.items() if (v[0] or v[1]) and k not in stored]
    check("gamma per strike matches QuikStrike", not diffs and not missing,
          f"{len(rows) - len(diffs)}/{len(rows)} equal"
          + (f"; first diffs {diffs[:3]}" if diffs else "")
          + (f"; nonzero strikes not stored {missing[:5]}" if missing else ""))
    calls, puts = sum(r[1] for r in rows), sum(r[2] for r in rows)
    check("Log totals", (last[3], last[4], last[5]) == (calls, puts, calls - puts),
          f"Log {last[3]}/{last[4]}/{last[5]} vs sheet {calls}/{puts}/{calls - puts}")
    best = max(rows, key=lambda r: r[1] + r[2])[0]
    check("peak strike", last[8] == best, f"Log {last[8]} vs recomputed {best}")
    # independent of update_gex.py's own answer: re-select rows from the freshly re-fetched
    # gamma (not the stored sheet) and recompute the flip the same way it does, then compare
    # against what's stored in Log - so a corrupted Gamma Data sheet fails this too, not just
    # a raw per-strike diff.
    # same reference price update_gex used (Log col C: live quote at run time, or the settle)
    flip = gamma_flip(select_rows(g), last[2] or fresh["price"])
    check("Gamma Flip matches independent recompute", last[7] == flip, f"Log {last[7]} vs recomputed {flip}")
    check("gamma peak within 100 of price", abs(best - (fresh["price"] or best)) <= 100,
          f"peak {best}, price {fresh['price']}", warn=True)
    check("edge strikes are zero", rows[0][1:] == (0, 0) and rows[-1][1:] == (0, 0),
          f"first {rows[0]}, last {rows[-1]}", warn=True)
    if use_excel:
        cells = excel_cells(path, {"GEX Calc": ["J2", "J3", "J4", "J5", "J6", "J9", "J10", "J15", "J16", "J18", "J30"]})
        if cells is None:
            skip("Excel-calculated GEX", "Excel unavailable or file could not be opened")
        else:
            x = {k.split("!")[1]: v for k, v in cells.items()}
            check("Excel GEX totals = independent",
                  (float(x["J2"]), float(x["J3"]), float(x["J4"])) == (calls, puts, calls - puts),
                  f"Excel {x['J2']}/{x['J3']}/{x['J4']} vs Python {calls}/{puts}/{calls - puts}")
            # J6 is the reference price update_gex used - the live quote at run time by default
            # (gex_price_source), else the settle - and it's what the Log's price column holds
            check("Excel price (J6) = Log price", last[2] is not None and x["J6"] != ""
                  and abs(float(x["J6"]) - float(last[2])) < 0.05, f"J6 {x['J6']} vs Log {last[2]}")
            check("Excel mode (J5) = Log status", x["J5"] == last[6], f"Excel '{x['J5']}' vs Log '{last[6]}'")
            check("Excel edge checks", x["J15"] == "ครบ" and x["J16"] == "ครบ", f"{x['J15']} / {x['J16']}")
            # GEX Log's levels are computed in Python (openpyxl can't read Excel results);
            # this proves that mirror matches the workbook's own J9/J10/J30/J18
            def num(v):
                try:
                    return float(v)
                except (TypeError, ValueError):
                    return None
            if len(last) > 13:
                xl = [num(x["J9"]), num(x["J10"]), num(x["J30"]), num(x["J18"])]
                py = [num(v) for v in last[10:14]]
                same = all((a is None and b is None) or (a is not None and b is not None and abs(a - b) < 1e-9)
                           for a, b in zip(xl, py))
                check("Log levels = Excel (Call/Put Wall, Pin, conviction)", same, f"Excel {xl} vs Log {py}")
            else:
                skip("Log levels = Excel", "GEX Log row predates the level columns")


def main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # Thai text on a cp1252 console
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-gamma", action="store_true")
    ap.add_argument("--no-excel", action="store_true")
    a = ap.parse_args()
    cfg = json.load(open(CONFIG, encoding="utf-8"))
    out = verify_max_pain(cfg, not a.no_excel)
    if a.no_gamma:
        skip("GEX verification", "--no-gamma")
    elif out:
        try:
            verify_gex(cfg, *out, not a.no_excel)
        except Exception as e:
            check("GEX verification ran", False, f"{type(e).__name__}: {e}")
    n = {s: results.count(s) for s in ("PASS", "WARN", "FAIL", "SKIP")}
    print(f"\nSummary: {n['PASS']} pass, {n['WARN']} warn, {n['FAIL']} fail, {n['SKIP']} skip")
    sys.exit(1 if n["FAIL"] else 0)


if __name__ == "__main__":
    main()
