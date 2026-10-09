"""run_state.py bookkeeping - temp stamp file, alert/ping functions replaced."""
from datetime import date

import run_state


def test_already_ran_compares_stamp_with_today(tmp_path):
    stamp = tmp_path / "last_ok.txt"
    assert run_state.already_ran(date(2026, 10, 5), stamp) is False  # no file yet
    stamp.write_text("2026-10-05")
    assert run_state.already_ran(date(2026, 10, 5), stamp) is True
    assert run_state.already_ran(date(2026, 10, 6), stamp) is False


def _record(monkeypatch):
    calls = {"ping": [], "notify": []}
    monkeypatch.setattr(run_state, "ping_health", lambda fail=False, body="": calls["ping"].append(fail))
    monkeypatch.setattr(run_state, "notify", lambda t, m: calls["notify"].append(m))
    return calls


def test_finish_ok_and_warn_stamp_the_day_and_ping(tmp_path, monkeypatch):
    calls = _record(monkeypatch)
    stamp = tmp_path / "last_ok.txt"
    run_state.finish("WARN", "GEX skipped", date(2026, 10, 5), stamp)
    assert stamp.read_text() == "2026-10-05"
    assert calls == {"ping": [False], "notify": []}


def test_finish_fail_pings_fail_alerts_and_leaves_no_stamp(tmp_path, monkeypatch):
    calls = _record(monkeypatch)
    stamp = tmp_path / "last_ok.txt"
    log = "\n".join(f"line {i}" for i in range(20)) + "\nTraceback: boom"
    run_state.finish("FAIL", log, date(2026, 10, 5), stamp)
    assert not stamp.exists()  # a later trigger the same day must retry
    assert calls["ping"] == [True]
    (msg,) = calls["notify"]
    assert "FAILED" in msg and "Traceback: boom" in msg and "line 0" not in msg  # only the tail


def test_stale_trade_date_lets_a_later_trigger_run_again(tmp_path):
    stamp = tmp_path / "last_ok.txt"
    # Fri 2026-10-09 09:00 run only got 10-07 (CME hadn't published 10-08): not done yet
    stamp.write_text("2026-10-09 2026-10-07")
    assert run_state.already_ran(date(2026, 10, 9), stamp) is False
    stamp.write_text("2026-10-09 2026-10-08")
    assert run_state.already_ran(date(2026, 10, 9), stamp) is True


def test_finish_stamps_the_logged_trade_date(tmp_path, monkeypatch):
    calls = _record(monkeypatch)
    stamp = tmp_path / "last_ok.txt"
    run_state.finish("OK", "x\nOK 2026-10-08 E21 Week 2 - OCT 2026 (V26) strikes=1\n", date(2026, 10, 9), stamp)
    assert stamp.read_text() == "2026-10-09 2026-10-08"
    run_state.finish("OK", "SKIP 2026-10-08 E21 ...", date(2026, 10, 9), stamp)
    assert stamp.read_text() == "2026-10-09 2026-10-08"
