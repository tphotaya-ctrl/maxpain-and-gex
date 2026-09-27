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


def test_fetch_gamma_does_not_retry_login_required(monkeypatch):
    calls = {"n": 0}

    def always_login(cfg, code, trade_date):
        calls["n"] += 1
        raise LoginRequired("no session")

    monkeypatch.setattr(fetch_gamma, "_fetch_gamma_once", always_login)
    with pytest.raises(LoginRequired):
        fetch_gamma_public({}, "FAKE", None)
    assert calls["n"] == 1  # not retried - a stale login won't fix itself
