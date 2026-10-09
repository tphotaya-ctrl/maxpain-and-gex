"""Pure-logic tests for sheets_sync.py - no network, no real Google API.

The live parts (sync()'s actual gspread/service-account calls) need a real credential and
Google Sheet and aren't covered here; this locks down the value-formatting and same-day
match-or-append logic that's cheap and safe to test in CI.
"""
from datetime import datetime

from sheets_sync import _cell, _match_row


def test_cell_formats_datetime_as_date_only_when_no_time_component():
    assert _cell(datetime(2026, 9, 25)) == "2026-09-25"


def test_cell_formats_datetime_with_time_component():
    assert _cell(datetime(2026, 9, 27, 16, 40, 48)) == "2026-09-27 16:40"


def test_cell_passes_through_numbers_and_strings():
    assert _cell(4330) == 4330
    assert _cell(4321.2) == 4321.2
    assert _cell("MW1 Week 4 - SEP 2026 (U26)") == "MW1 Week 4 - SEP 2026 (U26)"


def test_cell_turns_none_into_empty_string():
    assert _cell(None) == ""


HEADER = ["วันที่ข้อมูล", "สัญญา", "ราคา"]
EXISTING = [
    HEADER,
    ["2026-09-23", "E21 Week 4 - SEP 2026 (U26)", 4318.4],
    ["2026-09-25", "MW1 Week 4 - SEP 2026 (U26)", 4321.2],
]


def test_match_row_finds_existing_same_day_same_contract():
    # row 1 is the header, so "2026-09-25" (2nd data row) is sheet row 3
    assert _match_row(EXISTING, "2026-09-25", "MW1 Week 4 - SEP 2026 (U26)") == 3


def test_match_row_returns_none_for_a_new_day_or_contract():
    assert _match_row(EXISTING, "2026-09-27", "MW1 Week 4 - SEP 2026 (U26)") is None
    assert _match_row(EXISTING, "2026-09-25", "AB1 Week 4 - SEP 2026 (U26)") is None


def test_match_row_returns_none_when_sheet_only_has_a_header():
    assert _match_row([HEADER], "2026-09-25", "MW1 Week 4 - SEP 2026 (U26)") is None


def test_main_skips_quietly_when_not_configured(tmp_path, monkeypatch, capsys):
    import json
    import sheets_sync
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({"workbook": "x.xlsx", "gex_workbook": "x.xlsx"}))
    monkeypatch.setattr(sheets_sync, "CONFIG", cfg)
    monkeypatch.setattr(sheets_sync, "sync", lambda c: (_ for _ in ()).throw(AssertionError("must not sync")))
    sheets_sync.main()  # no exception -> no "skipped" alert from update_workbooks
    assert "not configured" in capsys.readouterr().out
