"""Render docs/index.html from both workbooks' Log sheets, for viewing on a phone or any
browser without opening Excel.

This script only writes the local docs/index.html file. Publishing it (git add/commit/push)
is a deliberate manual step - see README's "Mobile/web view" section - not something the
daily automation does unattended, since that would mean an unattended task pushing to a
public GitHub repo every day with no one looking.
"""
import json
import os
from datetime import datetime
from pathlib import Path

import openpyxl

from update_gex import GEX_LOG

HERE = Path(__file__).parent
CONFIG = Path(os.environ.get("MAXPAIN_CONFIG", HERE / "config.json"))
HISTORY_ROWS = 20


def _fmt(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d") if v.time() == datetime.min.time() else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, float):
        return f"{v:,.4g}" if abs(v) < 1 else f"{v:,.2f}"
    return "" if v is None else str(v)


def _pct(v):
    return f"{v:+.2%}" if isinstance(v, (int, float)) else "-"


def _table(header, rows):
    head = "".join(f"<th>{h or ''}</th>" for h in header)
    body = "".join("<tr>" + "".join(f"<td>{_fmt(v)}</td>" for v in r) + "</tr>" for r in reversed(rows))
    return f"<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"


def _read_log(path, n=HISTORY_ROWS, sheet="Log"):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    log = wb[sheet]
    rows_iter = log.iter_rows(values_only=True)
    header = list(next(rows_iter))
    rows = [list(r) for r in rows_iter]
    wb.close()
    return header, rows[-n:]


def render_html(mp_header, mp_rows, gex_header, gex_rows, generated_at):
    mp_last = mp_rows[-1] if mp_rows else None
    gex_last = gex_rows[-1] if gex_rows else None

    mp_card = ""
    if mp_last:
        mp_card = f"""
        <div class="card">
          <h2>Max Pain</h2>
          <p class="big">{_fmt(mp_last[4])}</p>
          <p>ราคา {_fmt(mp_last[3])} &middot; ห่างราคา {_pct(mp_last[5])} &middot; DTE {_fmt(mp_last[2])}</p>
          <p>{mp_last[1]}</p>
          <p class="muted">{mp_last[9]} &middot; อัปเดต {_fmt(mp_last[10])}</p>
        </div>"""

    gex_card = ""
    if gex_last:
        gex_card = f"""
        <div class="card">
          <h2>GEX</h2>
          <p class="big">{_fmt(gex_last[5])}</p>
          <p>{gex_last[6]}</p>
          <p>Gamma Flip {_fmt(gex_last[7])} &middot; Peak {_fmt(gex_last[8])}</p>
          <p class="muted">{gex_last[1]} &middot; อัปเดต {_fmt(gex_last[9])}</p>
        </div>"""

    return f"""<!doctype html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Max Pain / GEX</title>
<style>
  body {{ font-family: -apple-system, "Segoe UI", sans-serif; margin: 0; padding: 16px;
          background: #0b0d12; color: #e8e8e8; }}
  h1 {{ font-size: 1.05rem; color: #999; font-weight: 400; }}
  h3 {{ font-size: 0.9rem; color: #999; margin-top: 28px; }}
  .cards {{ display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 8px; }}
  .card {{ background: #161a22; border-radius: 10px; padding: 16px; flex: 1; min-width: 220px; }}
  .card h2 {{ margin: 0 0 8px; font-size: 0.9rem; color: #8ab4f8; }}
  .big {{ font-size: 1.8rem; font-weight: 600; margin: 0 0 4px; }}
  .muted {{ color: #888; font-size: 0.8rem; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 0.8rem; }}
  th, td {{ border-bottom: 1px solid #262c38; padding: 6px 8px; text-align: right; white-space: nowrap; }}
  th:first-child, td:first-child {{ text-align: left; }}
</style>
</head>
<body>
<h1>CME Gold Max Pain / GEX &middot; สร้างเมื่อ {_fmt(generated_at)}</h1>
<div class="cards">{mp_card}{gex_card}</div>
<h3>Max Pain - ประวัติล่าสุด</h3>
{_table(mp_header, mp_rows)}
<h3>GEX - ประวัติล่าสุด</h3>
{_table(gex_header, gex_rows)}
</body>
</html>
"""


def generate(cfg):
    mp_header, mp_rows = _read_log(HERE / cfg["workbook"])
    gex_header, gex_rows = _read_log(HERE / cfg["gex_workbook"], sheet=GEX_LOG)
    html = render_html(mp_header, mp_rows, gex_header, gex_rows, datetime.now())
    out_dir = HERE / "docs"
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / "index.html"
    out_path.write_text(html, encoding="utf-8")
    return out_path


def main():
    cfg = json.load(open(CONFIG, encoding="utf-8"))
    path = generate(cfg)
    print(f"Report written to {path}")


if __name__ == "__main__":
    main()
