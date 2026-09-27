"""Pure-logic tests for update_workbooks.py - no network, no Excel.

Expected values below were cross-checked by hand and by running max_pain() directly before
being written into these assertions - see the function's own docstring for the formula.
"""
from update_workbooks import max_pain, zone


def test_zone_buckets():
    mid, edge = 0.02, 0.05
    assert zone(0.001, mid, edge) == "กลาง"
    assert zone(-0.019, mid, edge) == "กลาง"
    assert zone(0.03, mid, edge) == "เปลี่ยนผ่าน"
    assert zone(-0.03, mid, edge) == "เปลี่ยนผ่าน"
    assert zone(0.06, mid, edge) == "มีผล"
    assert zone(-0.10, mid, edge) == "มีผล"
    assert zone(None, mid, edge) is None


def test_max_pain_single_strike_has_zero_pain():
    # only one strike with OI: settling there pays out nothing (no other strike to be ITM against)
    assert max_pain([(100, 10, 10)]) == (100, 0, 0)


def test_max_pain_picks_lowest_total_payout():
    # calls at 100 (20 lots), puts at 110 (10 lots), nothing at 105:
    #   pain(100) = (110-100)*10          = 100  <- minimum
    #   pain(105) = (105-100)*20 + (110-105)*10 = 150
    #   pain(110) = (110-100)*20          = 200
    assert max_pain([(100, 20, 0), (105, 0, 0), (110, 0, 10)]) == (100, 100, 1)


def test_max_pain_valley_counts_strikes_within_2pct_of_minimum():
    # symmetric OI around the middle strike: all three strikes work out to the same total pain
    assert max_pain([(90, 50, 0), (100, 0, 0), (110, 0, 50)]) == (90, 1000, 3)
