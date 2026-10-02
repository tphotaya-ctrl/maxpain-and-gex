"""Pure-logic tests for fetch_gamma.py - no network, no browser, no CME login.

Live QuikStrike scraping (_fetch_gamma_once) isn't covered here; this locks down the retry
wrapper's behaviour and the code-naming logic, both observed and fixed live this session
(see README "How it works" / commit history for fetch_gamma.py).
"""
import pytest

import fetch_gamma
from fetch_gamma import LoginRequired, fetch_gamma as fetch_gamma_public, qs_code


def test_qs_code_weekly():
    # confirmed against the real QuikStrike codes for these contracts, 2026-09
    assert qs_code("MW1", "Week 3 - SEP 2026") == "G3MU6"
    assert qs_code("AB1", "Week 4 - SEP 2026") == "G4TU6"


def test_qs_code_friday_uses_og_prefix():
    assert qs_code("E21", "Week 4 - SEP 2026") == "OG4U6"


def test_fetch_gamma_retries_transient_failures_then_succeeds(monkeypatch):
    calls = {"n": 0}

    def flaky(cfg, code, trade_date):
        calls["n"] += 1
        if calls["n"] < 3:
            raise RuntimeError(f"simulated transient failure #{calls['n']}")
        return {1: (1.0, 2.0)}

    monkeypatch.setattr(fetch_gamma, "_fetch_gamma_once", flaky)
    result = fetch_gamma_public({}, "FAKE", None)
    assert result == {1: (1.0, 2.0)}
    assert calls["n"] == 3


def test_fetch_gamma_gives_up_after_max_retries(monkeypatch):
    def always_fails(cfg, code, trade_date):
        raise RuntimeError("still broken")

    monkeypatch.setattr(fetch_gamma, "_fetch_gamma_once", always_fails)
    with pytest.raises(RuntimeError):
        fetch_gamma_public({}, "FAKE", None, retries=2)


def test_matrix_date_parses_toolbar_text():
    from datetime import date
    body = "(Set Expiration List) \tHeat By: Expiry\t Call/Put Combined\tWed, Sep 30, 2026\n0 DTE\tOG1V6"
    assert fetch_gamma.matrix_date(body) == date(2026, 9, 30)


# shape of the real expiry matrix (2026-10-01): a price row, STRIKE header with one
# "<code> <n> DTE" cell per expiration, a C/P row with one extra leading blank, then strikes
def _matrix(n_strikes):
    rows = [["", "GCZ6 4185.4", "GCZ6 4185.4"],
            ["STRIKE", "G1RV6 0 DTE", "OG1V6 1 DTE"],
            ["", "", "C", "P", "C", "P"]]
    rows += [[str(4000 + 5 * i), str(i), str(2 * i), "", "7"] for i in range(n_strikes)]
    return rows


def test_latest_column_picks_the_codes_pair():
    from datetime import date
    body = "Wed, Sep 30, 2026"
    g1 = fetch_gamma._latest_column(_matrix(70), body, "G1RV6", date(2026, 9, 30))
    assert g1[4005] == (1.0, 2.0)
    og = fetch_gamma._latest_column(_matrix(70), body, "OG1V6", date(2026, 9, 30))
    assert og[4005] == (0.0, 7.0)  # blank cell reads as 0


def test_latest_column_refuses_wrong_date_or_partial_table():
    from datetime import date
    with pytest.raises(RuntimeError, match="shows 2026-09-30"):
        fetch_gamma._latest_column(_matrix(70), "Wed, Sep 30, 2026", "G1RV6", date(2026, 10, 1))
    with pytest.raises(RuntimeError, match="only had 31 strikes"):
        fetch_gamma._latest_column(_matrix(31), "Wed, Sep 30, 2026", "G1RV6", date(2026, 9, 30))


def test_history_gap_falls_back_to_latest_matrix(monkeypatch):
    def no_date(cfg, code, trade_date):
        raise fetch_gamma.DateNotInHistory("no gamma column for 9/30/2026")
    monkeypatch.setattr(fetch_gamma, "_fetch_gamma_once", no_date)
    monkeypatch.setattr(fetch_gamma, "_fetch_latest_once", lambda cfg, code, td: {1: (3.0, 4.0)})
    assert fetch_gamma_public({}, "FAKE", None) == {1: (3.0, 4.0)}


def test_fetch_gamma_does_not_retry_login_required(monkeypatch):
    calls = {"n": 0}

    def always_login(cfg, code, trade_date):
        calls["n"] += 1
        raise LoginRequired("no session")

    monkeypatch.setattr(fetch_gamma, "_fetch_gamma_once", always_login)
    with pytest.raises(LoginRequired):
        fetch_gamma_public({}, "FAKE", None)
    assert calls["n"] == 1  # not retried - a stale login won't fix itself
