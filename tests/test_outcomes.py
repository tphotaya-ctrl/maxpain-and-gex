"""Pure-logic tests for outcomes.py - no network, no browser, no workbook."""
from datetime import date

import pytest

from outcomes import outcome_metrics, parse_contract, px, trading_days


def test_parse_contract_weekly():
    assert parse_contract("MW1 Week 4 - SEP 2026 (U26)") == ("MW1", "Week 4 - SEP 2026")
    assert parse_contract("E21 Week 1 - OCT 2026 (V26)") == ("E21", "Week 1 - OCT 2026")


def test_parse_contract_rejects_monthly_and_junk():
    assert parse_contract("AME OCT 2026 (V26)") is None
    assert parse_contract(None) is None
    assert parse_contract("") is None


def test_px_strips_cme_suffixes_and_commas():
    assert px("4,381.0B") == 4381.0
    assert px("4321.2") == 4321.2
    assert px("-41.0") == -41.0
    assert px("+23.2") == 23.2
    assert px("-") is None
    assert px("") is None
    assert px(None) is None


def test_trading_days_skips_weekend_and_excludes_start():
    # Fri 25 Sep -> Mon 28 Sep: only the Monday
    assert trading_days(date(2026, 9, 25), date(2026, 9, 28)) == [date(2026, 9, 28)]
    # Fri 18 Sep -> Tue 22 Sep: Mon 21 and Tue 22
    assert trading_days(date(2026, 9, 18), date(2026, 9, 22)) == [date(2026, 9, 21), date(2026, 9, 22)]


def test_outcome_metrics_price_moved_away_from_max_pain():
    # start 4321.2, MP 4330 (8.8 away); settled 4200 (130 away) -> moved away
    m = outcome_metrics(4321.2, 4330, 4200, 4280, 4180)
    assert m["d_start"] == pytest.approx(8.8)
    assert m["d_end"] == pytest.approx(130)
    assert m["move"] == pytest.approx(121.2)
    assert m["toward"] is False
    assert m["range_pct"] == pytest.approx(100 / 4321.2)


def test_outcome_metrics_toward_and_missing_range():
    m = outcome_metrics(100, 110, 108, None, None)
    assert m["toward"] is True  # 2 away at settle vs 10 away at start
    assert m["range_pct"] is None
