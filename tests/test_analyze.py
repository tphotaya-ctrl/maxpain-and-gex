"""Pure-logic tests for analyze.py's statistics - no workbook."""
import pytest

from analyze import MIN_N, format_report, summarize

# A: MP 110 above start 100, settled 108 -> 2 from MP (was 10): toward; naive error 8
# B: MP 90 below start 100, settled 105 -> 15 from MP (was 10): away; naive error 5
A = {"start": 100, "mp": 110, "settle": 108, "net": 5, "range_pct": 0.02}
B = {"start": 100, "mp": 90, "settle": 105, "net": -3, "range_pct": 0.06}


def test_summarize_core_numbers():
    s = summarize([A, B], 0.02, 0.05)
    assert s["n"] == 2
    assert s["toward_rate"] == 0.5
    assert s["mae_mp"] == pytest.approx(8.5)     # (2 + 15) / 2
    assert s["mae_naive"] == pytest.approx(6.5)  # (8 + 5) / 2


def test_summarize_range_split_by_gex_sign():
    s = summarize([A, B], 0.02, 0.05)
    assert s["range_by_gex"]["NET GEX > 0"] == (1, 0.02)
    assert s["range_by_gex"]["NET GEX < 0"] == (1, 0.06)


def test_summarize_zones_use_workbook_buckets():
    # both are 10% away from MP at log time -> beyond edge 5% -> 'มีผล'
    s = summarize([A, B], 0.02, 0.05)
    assert set(s["by_zone"]) == {"มีผล"}
    assert s["by_zone"]["มีผล"]["n"] == 2


def test_summarize_ignores_missing_gex_and_range():
    c = {"start": 100, "mp": 101, "settle": 100.5, "net": None, "range_pct": None}
    s = summarize([c], 0.02, 0.05)
    assert s["range_by_gex"]["NET GEX > 0"] == (0, None)
    assert set(s["by_zone"]) == {"กลาง"}


def test_report_warns_on_small_sample_and_names_the_better_predictor():
    text = format_report(summarize([A, B], 0.02, 0.05))
    assert f"fewer than {MIN_N}" in text
    assert "naive (no change)" in text  # 8.5 vs 6.5: the naive guess was closer


def test_paper_trades_only_use_rows_with_an_entry():
    from analyze import format_paper, paper_trades
    full = dict(A, hi=112, lo=99, entry=100, change=3, mode="Negative GEX (วิ่ง)",
                call_wall=None, put_wall=None)
    old = dict(B, hi=None, lo=None, entry=None, change=None, mode=None, call_wall=None, put_wall=None)
    res = paper_trades([full, old])
    assert res["Baseline: always long"]["n"] == 1        # old row has no entry -> skipped
    assert res["R2 Momentum (-GEX)"]["total"] == 8       # long 100 -> settle 108
    assert res["R1 Range fade (+GEX)"]["n"] == 0
    assert "WARNING" in format_paper(res)


def test_report_with_no_data():
    assert "nothing to analyze" in format_report(summarize([], 0.02, 0.05))
