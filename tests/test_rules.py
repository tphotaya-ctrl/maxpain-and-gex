"""Hand-worked cases for the paper-trading rules - no data files, no network."""
from rules import baseline_long, r1_range_fade, r2_momentum, r3_max_pain, stats

POS, NEG, NONE = "Positive GEX (นิ่ง)", "Negative GEX (วิ่ง)", "ไม่มีโหมด"


def rec(**kw):
    base = dict(entry=100, settle=110, hi=115, lo=95, mp=None, mode=POS,
                call_wall=120, put_wall=90, change=5)
    base.update(kw)
    return base


def test_r1_buys_lower_half_and_holds_to_settle():
    # walls 90..120, mid 105, entry 100 -> long, stop 90 not touched (low 95) -> +10
    t = r1_range_fade(rec())
    assert (t.direction, t.exit, t.stopped, t.pnl) == (1, 110, False, 10)


def test_r1_stop_is_worst_case():
    # low 89 reaches the 90 stop -> assume stopped there even though settle was 110
    t = r1_range_fade(rec(lo=89))
    assert (t.exit, t.stopped, t.pnl) == (90, True, -10)


def test_r1_sells_upper_half():
    # entry 110 > mid 105 -> short, stop 120 not touched (high 115), settle 104 -> +6
    t = r1_range_fade(rec(entry=110, settle=104))
    assert (t.direction, t.pnl) == (-1, 6)


def test_r1_needs_positive_mode_and_entry_inside_walls():
    assert r1_range_fade(rec(mode=NEG)) is None
    assert r1_range_fade(rec(mode=NONE)) is None
    assert r1_range_fade(rec(mode="Positive GEX (นิ่ง)", entry=125)) is None
    assert r1_range_fade(rec(put_wall=None)) is None
    assert r1_range_fade(rec(entry=105)) is None  # exactly at mid: no edge, no trade


def test_r2_follows_change_in_negative_mode():
    # change -5 -> short at 100, stop at call wall 120 (high 115, not hit), settle 110 -> -10
    t = r2_momentum(rec(mode=NEG, change=-5))
    assert (t.direction, t.stopped, t.pnl) == (-1, False, -10)
    # change +5 -> long, stop at put wall 90 (low 95, not hit) -> +10
    assert r2_momentum(rec(mode=NEG, change=5)).pnl == 10


def test_r2_skips_without_mode_or_change():
    assert r2_momentum(rec(mode=POS, change=-5)) is None
    assert r2_momentum(rec(mode=NEG, change=0)) is None
    assert r2_momentum(rec(mode=NEG, change=None)) is None


def test_r3_trades_toward_max_pain_only_when_far_enough():
    assert r3_max_pain(rec(mp=103)).pnl == 10        # 3% above -> long
    assert r3_max_pain(rec(mp=97)).pnl == -10        # 3% below -> short
    assert r3_max_pain(rec(mp=100.3)) is None        # 0.3% < 0.5%


def test_baseline_and_missing_prices():
    assert baseline_long(rec()).pnl == 10
    assert baseline_long(rec(entry=None)) is None
    assert baseline_long(rec(settle=None)) is None
    assert r1_range_fade(rec(hi=None, lo=None)) is None  # stop can't be judged


def test_stats():
    trades = [baseline_long(rec()), baseline_long(rec(settle=90)), baseline_long(rec(settle=100))]
    s = stats(trades)
    assert s["n"] == 3 and s["total"] == 0 and s["worst"] == -10
    assert s["win_rate"] == 1 / 3
    assert stats([]) == {"n": 0}


def test_all_expiration_variants_read_the_rolled_up_gex():
    from rules import r1_range_fade_all, r2_momentum_all
    # single-contract says negative (so R1 is off), the all-expiration reading says positive
    r = rec(mode=NEG, mode_all=POS, call_wall_all=120, put_wall_all=90)
    assert r1_range_fade(r) is None
    t = r1_range_fade_all(r)
    assert (t.direction, t.pnl) == (1, 10)  # same R1 logic, the รวม walls 90..120
    assert r2_momentum_all(r) is None       # รวม is positive -> no momentum trade
    assert r1_range_fade_all(rec()) is None  # no รวม data on the row -> no trade
