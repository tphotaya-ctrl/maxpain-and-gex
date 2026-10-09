"""Pure-logic tests for update_gex.py - no network, no Excel.

Expected values were cross-checked by hand and by running the functions directly before
being written into these assertions.
"""
import pytest

from update_gex import gamma_flip, select_rows


def test_select_rows_includes_one_zero_strike_each_side():
    g = {100: (0, 0), 105: (5, 2), 110: (2, 8), 115: (0, 0)}
    assert select_rows(g) == [(100, 0, 0), (105, 5, 2), (110, 2, 8), (115, 0, 0)]


def test_select_rows_all_zero_raises():
    with pytest.raises(RuntimeError):
        select_rows({100: (0, 0), 105: (0, 0)})


def test_gamma_flip_finds_sign_change():
    rows = select_rows({100: (0, 0), 105: (5, 2), 110: (2, 8), 115: (0, 0)})
    # cumulative net gamma: 0 -> +3 (at 105) -> -3 (at 110, sign change) -> stays -3 (at 115)
    assert gamma_flip(rows, price=108) == 110
    assert gamma_flip(rows, price=100) == 110  # only one flip point either way


def test_gamma_flip_none_when_cumulative_never_changes_sign():
    rows = select_rows({100: (5, 1), 105: (3, 1)})  # always net-positive
    assert gamma_flip(rows, price=102) is None


def test_ref_price_prefers_live_unless_settle_configured():
    from update_gex import ref_price
    d = {"price": 4187.1, "live_price": 4168.4}
    assert ref_price({}, d) == 4168.4
    assert ref_price({"gex_price_source": "settle"}, d) == 4187.1
    assert ref_price({}, {"price": 4187.1, "live_price": None}) == 4187.1


def test_aggregate_sums_per_strike_across_expirations():
    from update_gex import aggregate
    per_code = {"OG2V6": {4100: (1, 5), 4200: (3, 0)}, "OGX6": {4200: (10, 2), 4300: (4, 0)}}
    assert aggregate(per_code) == [(4100, 1.0, 5.0), (4200, 13.0, 2.0), (4300, 4.0, 0.0)]


def test_levels_apply_the_workbook_rules():
    from update_gex import levels
    rows = [(4000, 0, 40), (4050, 5, 30), (4100, 20, 18), (4150, 30, 2), (4200, 60, 10), (4250, 25, 0)]
    lv = levels(rows, 4120)
    assert lv["call_wall"] == 4200          # biggest call >= 1.5x put at/above price
    assert lv["put_wall"] == 4000           # biggest put >= 1.5x call at/below price; 4100 is mixed
    assert lv["peak"] == 4200 and lv["net"] == 140 - 100
    assert lv["flip"] == 4200               # cumulative -40,-65,-63,-35,+15: turns positive at 4200
    assert lv["mode"].startswith("Positive")
