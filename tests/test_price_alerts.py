"""price_alerts - levels, basis, touches and the once-per-day rule, on hand-made rows and bars
(no Sheet, no Yahoo, no Telegram)."""
from datetime import date, datetime, timedelta, timezone

import price_alerts as pa

T0 = datetime(2026, 10, 8, 1, 0, tzinfo=timezone.utc)
MP_HEAD = ["วันที่ข้อมูล", "สัญญา", "DTE", "ราคา", "Max Pain"]
GEX_HEAD = ["วันที่ข้อมูล", "สัญญา", "ราคา", "สถานะ", "Gamma Flip (ใกล้ราคา)", "Call Wall", "Put Wall"]


def bars(*hlc):
    return [(T0 + timedelta(minutes=5 * i), h, l, c) for i, (h, l, c) in enumerate(hlc)]


def tabs(mp_date="2026-10-07", gex_date="2026-10-07"):
    return {
        "Max Pain Log": [MP_HEAD, ["2026-10-06", "x", 1, 4187.1, 4200], [mp_date, "y", 2, 4140.7, 4190]],
        "GEX Log": [GEX_HEAD, [gex_date, "OG2V6", 4212.2, "Negative GEX (วิ่ง)", "", 4250, 4100]],
    }


def test_levels_take_the_newest_rows_and_skip_blanks():
    info = pa.levels(tabs(), date(2026, 10, 8))
    assert info["trade_date"] == date(2026, 10, 7) and info["settle"] == 4140.7
    assert info["levels"] == {"Max Pain": 4190, "Call Wall": 4250, "Put Wall": 4100}  # no Gamma Flip
    assert info["mode"] == "Negative GEX (วิ่ง)"


def test_levels_skip_stale_rows():
    # GEX failed for days: its last row belongs to an expired contract
    info = pa.levels(tabs(gex_date="2026-10-01"), date(2026, 10, 8))
    assert info["levels"] == {"Max Pain": 4190}
    assert pa.levels(tabs(), date(2026, 10, 12))["levels"] == {}   # Wed's levels on Mon: too old


def test_levels_from_workbook_style_values():
    t = {"Max Pain Log": [MP_HEAD, [datetime(2026, 10, 7), "y", 2, 4140.7, 4190]]}
    assert pa.levels(t, date(2026, 10, 8))["levels"] == {"Max Pain": 4190}


def test_touches_up_down_none_and_since():
    b = bars((4185, 4180, 4184), (4192, 4183, 4191), (4192, 4170, 4172))
    hits = pa.touches(b, {"Max Pain": 4190, "Put Wall": 4175, "Call Wall": 4250}, T0 - timedelta(1))
    assert [(h[0], h[2], h[3]) for h in hits] == [("Max Pain", "up", b[1][0]), ("Put Wall", "down", b[2][0])]
    # bars up to and including `since` are already checked
    assert pa.touches(b, {"Max Pain": 4190}, b[2][0]) == []
    assert pa.touches(b, {"Max Pain": 4190}, b[1][0])[0][3] == b[2][0]  # bar 3 spans it again


def test_touches_apply_the_basis_offset():
    b = bars((4170, 4165, 4168))
    assert pa.touches(b, {"Max Pain": 4190}, T0 - timedelta(1)) == []
    assert pa.touches(b, {"Max Pain": 4190}, T0 - timedelta(1), offset=20)[0][4] == 4188


def test_basis():
    assert pa.basis(4140.7, {date(2026, 10, 7): 4140.7}, date(2026, 10, 7)) == (0, None)
    off, warn = pa.basis(4200, {date(2026, 10, 7): 4100}, date(2026, 10, 7))
    assert off == 100 and "เดือนสัญญาต่างกัน" in warn
    assert pa.basis(4200, {}, date(2026, 10, 7))[0] == 0


def test_daily_closes_use_the_session_date():
    d = [(datetime(2026, 10, 7, 4, tzinfo=timezone.utc), 1, 1, 4140.7)]
    assert pa.daily_closes(d) == {date(2026, 10, 7): 4140.7}


def test_state_and_once_per_trade_date():
    rows = [pa.ALERT_HEAD, ["2026-10-07", "Max Pain", 4190, "x"],
            [pa.LAST_CHECKED, "", "", "2026-10-08T01:00:00+00:00"]]
    sent, last = pa.read_state(rows)
    assert last == T0 and sent == {("2026-10-07", "Max Pain", 4190.0)}
    hits = [("Max Pain", 4190.0, "up", T0, 4191), ("Put Wall", 4100.0, "down", T0, 4099)]
    assert [h[0] for h in pa.new_alerts(hits, sent, date(2026, 10, 7))] == ["Put Wall"]
    assert len(pa.new_alerts(hits, sent, date(2026, 10, 8))) == 2   # next day's levels: fresh


def test_window_start():
    now = T0 + timedelta(hours=1)
    assert pa.window_start(None, now) == now - pa.FIRST_LOOKBACK
    assert pa.window_start(T0, now) == T0
    assert pa.window_start(T0 - timedelta(days=2), now) == now - pa.MAX_LOOKBACK


def test_format_alert():
    info = pa.levels(tabs(), date(2026, 10, 8))
    text = pa.format_alert(("Max Pain", 4190.0, "up", T0, 4191), info, 4191.0, None)
    assert "ราคาทองขึ้นแตะ Max Pain 4,190 เวลา 08:00 (ไทย)" in text
    assert "โหมด Negative GEX (วิ่ง)" in text and "ราคาอยู่ในกรอบ Put Wall–Call Wall" in text


def test_parse_chart_drops_empty_bars():
    d = {"chart": {"result": [{"timestamp": [1, 2], "indicators": {"quote": [
        {"high": [2.0, None], "low": [1.0, None], "close": [1.5, None]}]}}]}}
    assert pa.parse_chart(d) == [(datetime.fromtimestamp(1, timezone.utc), 2.0, 1.0, 1.5)]


class FakeWS:
    def __init__(self, rows):
        self.rows = rows

    def get_all_values(self, value_render_option=None):
        return [list(r) for r in self.rows]

    def append_row(self, r):
        self.rows.append(list(r))

    def append_rows(self, rs):
        self.rows += [list(r) for r in rs]

    def update(self, a1, values):
        self.rows[int(a1[1:]) - 1] = list(values[0])


class FakeSheet:
    def __init__(self, tabs):
        self.ws = {k: FakeWS(v) for k, v in tabs.items()}

    def worksheet(self, title):
        import gspread
        if title not in self.ws:
            raise gspread.WorksheetNotFound(title)
        return self.ws[title]

    def add_worksheet(self, title, rows, cols):
        self.ws[title] = FakeWS([])
        return self.ws[title]


def test_run_sends_once_and_keeps_state_in_the_sheet(monkeypatch):
    import sheets_sync
    import telegram_report
    sheet = FakeSheet(tabs())
    sent = []
    monkeypatch.setattr(sheets_sync, "open_sheet", lambda cfg: sheet)
    monkeypatch.setenv("GOOGLE_SHEET_ID", "test-sheet")
    monkeypatch.setattr(telegram_report, "send_text", lambda t: sent.append(t) or True)
    b = bars((4185, 4180, 4184), (4192, 4183, 4191))
    monkeypatch.setattr(pa, "yahoo_bars", lambda interval="5m", range_="2d":
                        b if interval == "5m" else [(datetime(2026, 10, 7, 4, tzinfo=timezone.utc), 0, 0, 4140.7)])
    now = T0 + timedelta(minutes=30)
    assert [r[1] for r in pa.run(now=now)] == ["Max Pain"]
    assert len(sent) == 1 and "Max Pain 4,190" in sent[0]
    alerts = sheet.ws[pa.ALERTS].rows
    assert alerts[0] == pa.ALERT_HEAD and alerts[1][:3] == ["2026-10-07", "Max Pain", 4190.0]
    # next hour: the same touch is not sent again, last_checked row is updated in place
    pa.run(now=now + timedelta(hours=1))
    assert len(sent) == 1
    assert [r[0] for r in sheet.ws[pa.ALERTS].rows].count(pa.LAST_CHECKED) == 1


def test_run_skips_quietly_until_configured(monkeypatch, capsys):
    import sheets_sync
    monkeypatch.delenv("GOOGLE_SHEET_ID", raising=False)
    monkeypatch.setattr(sheets_sync, "sheet_id", lambda cfg: None)
    assert pa.run() == []
    assert "not configured" in capsys.readouterr().out


def test_public_actions_log_has_no_levels_or_prices(monkeypatch, capsys):
    import sheets_sync
    import telegram_report
    monkeypatch.setattr(pa, "PUBLIC_LOG", True)
    monkeypatch.setattr(sheets_sync, "open_sheet", lambda cfg: FakeSheet(tabs()))
    monkeypatch.setenv("GOOGLE_SHEET_ID", "test-sheet")
    monkeypatch.setattr(telegram_report, "send_text", lambda t: True)
    b = bars((4185, 4180, 4184), (4192, 4183, 4191))
    monkeypatch.setattr(pa, "yahoo_bars", lambda interval="5m", range_="2d":
                        b if interval == "5m" else [(datetime(2026, 10, 7, 4, tzinfo=timezone.utc), 0, 0, 4140.7)])
    assert len(pa.run(now=T0 + timedelta(minutes=30))) == 1
    out = capsys.readouterr().out
    assert out.strip() == "price alerts: checked"   # not even which level fired
