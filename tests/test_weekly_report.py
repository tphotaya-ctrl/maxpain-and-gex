"""weekly_report scheduling and text - no workbook, no Telegram."""
from datetime import date

import weekly_report as wr

A = {"start": 100, "mp": 110, "settle": 108, "net": 5, "range_pct": 0.02, "entry": 100, "hi": 112,
     "lo": 99, "change": 3, "mode": "Negative GEX (วิ่ง)", "call_wall": None, "put_wall": None}


def test_target_week_is_this_week_from_friday_else_last_week():
    assert wr.target_week(date(2026, 10, 9)) == "2026-W41"   # Fri
    assert wr.target_week(date(2026, 10, 10)) == "2026-W41"  # Sat
    assert wr.target_week(date(2026, 10, 12)) == "2026-W41"  # Mon: Friday's run never happened
    assert wr.target_week(date(2026, 10, 8)) == "2026-W40"   # Thu: still last week's


def test_due_once_per_week(tmp_path):
    stamp = tmp_path / "weekly_sent.txt"
    assert wr.due(date(2026, 10, 9), stamp) is True
    stamp.write_text("2026-W41")
    assert wr.due(date(2026, 10, 9), stamp) is False
    assert wr.due(date(2026, 10, 12), stamp) is False         # Monday: W41 already sent
    assert wr.due(date(2026, 10, 16), stamp) is True          # next Friday: W42


def test_gex_coverage_counts_this_weeks_trade_dates_only():
    log = [date(2026, 10, 2), date(2026, 10, 5), date(2026, 10, 6), date(2026, 10, 7), date(2026, 10, 8)]
    gex = [date(2026, 10, 2), date(2026, 10, 7)]
    assert wr.week_bounds("2026-W41") == (date(2026, 10, 5), date(2026, 10, 9))
    assert wr.gex_coverage(log, gex, "2026-W41") == (4, 1)


def test_format_weekly_text():
    text = wr.format_weekly("2026-W41", [A], 0.02, 0.05, (4, 1))
    assert "สรุปรายสัปดาห์ 2026-W41" in text
    assert "ตัวอย่างผลจริง: 1/30 สัญญา (ยังเป็น noise)" in text
    assert "Max Pain คลาด 2.0 จุด vs เดาไม่ขยับ 8.0 -> Max Pain แม่นกว่า" in text
    assert "R2 Momentum (-GEX): 1 เทรด ชนะ 100% รวม +8.0 จุด" in text
    assert "R1 Range fade (+GEX): ยังไม่มีเทรด" in text
    assert "GEX สัปดาห์นี้: ได้ 1/4 วัน" in text


def test_main_only_stamps_after_a_delivered_send(tmp_path, monkeypatch):
    import telegram_report
    monkeypatch.setattr(wr, "STAMP", tmp_path / "weekly_sent.txt")
    monkeypatch.setattr(wr, "build", lambda cfg, week: f"report {week}")
    monkeypatch.setattr(wr, "CONFIG", tmp_path / "config.json")
    (tmp_path / "config.json").write_text("{}")
    monkeypatch.setattr(telegram_report, "send_text", lambda text: False)
    assert wr.main(today=date(2026, 10, 9)) == "report 2026-W41"
    assert not (tmp_path / "weekly_sent.txt").exists()     # not delivered -> try again next run
    monkeypatch.setattr(telegram_report, "send_text", lambda text: True)
    wr.main(today=date(2026, 10, 9))
    assert (tmp_path / "weekly_sent.txt").read_text() == "2026-W41"
    assert wr.main(today=date(2026, 10, 9)) is None         # already sent this week
