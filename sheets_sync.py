"""Sync each workbook's latest Log row into a private Google Sheet, so results can be
checked from the Google Sheets app on a phone - see README's "Mobile/web view" section for
the one-time Google Cloud service-account setup. Only syncs the current run's row into each
tab (same match-and-replace-or-append behaviour the workbooks' own Log sheets already use);
it does not backfill older history.
"""
import json
import os
from datetime import datetime
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials

from publish_report import _read_log
from update_gex import AGG_LOG, GEX_LOG

HERE = Path(__file__).parent
CONFIG = Path(os.environ.get("MAXPAIN_CONFIG", HERE / "config.json"))
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def _cell(v):
    """Sheet-safe value: gspread/JSON can't serialize datetime, everything else passes through
    as a native number/string so the sheet keeps real numeric values, not display strings."""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d") if v.time() == datetime.min.time() else v.strftime("%Y-%m-%d %H:%M")
    return "" if v is None else v


def _match_row(existing, date_str, contract):
    """1-based sheet row number of an existing row with the same date+contract (row 1 is the
    header), or None if there's no match yet."""
    for i, r in enumerate(existing[1:], start=2):
        if len(r) > 1 and r[0] == date_str and r[1] == contract:
            return i
    return None


def _worksheet(sh, title, header):
    try:
        return sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=max(len(header), 1))
        ws.append_row(header)
        return ws


def _sync_tab(ws, header, row):
    values = [_cell(v) for v in row]
    existing = ws.get_all_values()
    if not existing:
        ws.append_row(header)
        existing = [header]
    match = _match_row(existing, values[0], values[1])
    if match:
        ws.update(f"A{match}", [values])
    else:
        ws.append_row(values)


def sheet_id(cfg):
    return os.environ.get("GOOGLE_SHEET_ID") or cfg.get("google_sheet_id")


def open_sheet(cfg):
    """The private Google Sheet. In GitHub Actions the key and id come from repository
    secrets (env vars) - there is no service_account.json on the runner."""
    info = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if info:
        creds = Credentials.from_service_account_info(json.loads(info), scopes=SCOPES)
    else:
        creds_path = HERE / cfg.get("google_credentials_file", "service_account.json")
        creds = Credentials.from_service_account_file(str(creds_path), scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(sheet_id(cfg))


def sync(cfg):
    sh = open_sheet(cfg)

    mp_header, mp_rows = _read_log(HERE / cfg["workbook"], n=1)
    if mp_rows:
        _sync_tab(_worksheet(sh, "Max Pain Log", mp_header), mp_header, mp_rows[-1])

    gex_header, gex_rows = _read_log(HERE / cfg["gex_workbook"], n=1, sheet=GEX_LOG)
    if gex_rows:
        _sync_tab(_worksheet(sh, "GEX Log", gex_header), gex_header, gex_rows[-1])

    try:  # all-expiration walls, for the cloud price alerts (price_alerts.py)
        agg_header, agg_rows = _read_log(HERE / cfg["gex_workbook"], n=1, sheet=AGG_LOG)
    except KeyError:
        agg_rows = None
    if agg_rows:
        _sync_tab(_worksheet(sh, "GEX รวม", agg_header), agg_header, agg_rows[-1])


def main():
    cfg = json.load(open(CONFIG, encoding="utf-8"))
    if not sheet_id(cfg):
        # not set up (see README "Mobile/web view") - skip quietly; raising here made
        # update_workbooks send a "Sheets sync skipped" Telegram alert on every run
        print("Sheets sync: not configured (no google_sheet_id) - skipped")
        return
    sync(cfg)
    print("Synced to Google Sheet")


if __name__ == "__main__":
    main()
