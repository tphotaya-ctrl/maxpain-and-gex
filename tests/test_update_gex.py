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
