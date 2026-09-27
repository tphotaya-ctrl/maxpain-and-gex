"""Pure-logic tests for fetch_cme.py - no network, no browser.

The live parts (fetch(), _underlying_future_month(), _settle_price()'s actual HTTP call)
need a real CME session and aren't covered here; this locks down the parsing/matching logic
that's cheap and safe to test in CI.
"""
from datetime import date

from fetch_cme import _candidates, _label_key, _monthly_dates, _num, _settle_price, expiry_date


def test_expiry_date_monday():
    # 'Week 3 - SEP 2026' on the Monday product -> 3rd Monday of Sep 2026
    assert expiry_date("MW1", "Week 3 - SEP 2026") == date(2026, 9, 21)


def test_expiry_date_friday():
    assert expiry_date("E21", "Week 1 - OCT 2026") == date(2026, 10, 2)


def test_label_key_weekly_sorts_chronologically():
    labels = ["Week 2 - OCT 2026", "Week 4 - SEP 2026", "Week 1 - OCT 2026"]
    assert sorted(labels, key=_label_key) == ["Week 4 - SEP 2026", "Week 1 - OCT 2026", "Week 2 - OCT 2026"]


def test_label_key_monthly():
    assert _label_key("OCT 2026") == (2026, 10, 0)
    assert _label_key("garbage") == (9999, 0, 0)


def test_num_parses_commas_and_dashes():
    assert _num("1,234") == 1234
    assert _num("-56") == -56
    assert _num("-") == 0
    assert _num("") == 0


WEEKLY_GROUP = {
    "optionType": "MW1",
    "expirations": [
        {"label": "Week 3 - SEP 2026", "productId": 1, "expirationCode": "U26"},
        {"label": "Week 4 - SEP 2026", "productId": 2, "expirationCode": "U26"},
    ],
}
AME_GROUP = {
    "optionType": "AME",
    "expirations": [
        {"label": "OCT 2026", "productId": 192, "expirationCode": "V26",
         "expiration": {"month": 10, "year": 2026}},
        {"label": "DEC 2026", "productId": 192, "expirationCode": "Z26",
         "expiration": {"month": 12, "year": 2026}},
    ],
}


def _stub_get(monthly_dates):
    def get(url):
        assert "atm/expirations" in url
        return [{"optionType": "AME", "contractExpirations": [
            {"expirationMonth": m, "expirationYear": y, "lastTradeDate": d.isoformat() + "T00:00:00"}
            for (m, y), d in monthly_dates.items()
        ]}]
    return get


def test_candidates_today_picks_weekly_exact_match():
    cfg = {"target": "2026-09-21", "family": "MW1", "include_monthly": False}
    hit = _candidates(_stub_get({}), [WEEKLY_GROUP], cfg)
    assert [e["label"] for e in hit] == ["Week 3 - SEP 2026"]  # 3rd Monday of Sep 2026


def test_candidates_falls_back_to_next_expiry_when_none_matches_exactly():
    cfg = {"target": "2026-09-20", "family": "MW1", "include_monthly": False}  # a Sunday
    hit = _candidates(_stub_get({}), [WEEKLY_GROUP], cfg)
    assert [e["label"] for e in hit] == ["Week 3 - SEP 2026"]  # next Monday after the 20th


def test_candidates_raises_when_nothing_expires_on_or_after():
    cfg = {"target": "2027-01-01", "family": "MW1", "include_monthly": False}
    try:
        _candidates(_stub_get({}), [WEEKLY_GROUP], cfg)
        assert False, "expected RuntimeError"
    except RuntimeError:
        pass


def test_candidates_auto_sorts_and_filters_to_family():
    cfg = {"target": "auto", "family": "MW1"}
    hit = _candidates(_stub_get({}), [WEEKLY_GROUP], cfg)
    assert [e["label"] for e in hit] == ["Week 3 - SEP 2026", "Week 4 - SEP 2026"]


def test_candidates_monthly_joined_by_month_year_not_label():
    # this is the behaviour that matters: CME's two endpoints disagree on the AME label for
    # the same contract, so the join must use the numeric (month, year), not the label string
    monthly_dates = {(10, 2026): date(2026, 10, 27), (12, 2026): date(2026, 12, 28)}
    cfg = {"target": "2026-10-27", "family": "AME", "include_monthly": True, "underlying_product_id": 437}
    hit = _candidates(_stub_get(monthly_dates), [AME_GROUP], cfg)
    assert [e["label"] for e in hit] == ["OCT 2026"]


def test_candidates_monthly_disabled_by_include_monthly_false():
    monthly_dates = {(10, 2026): date(2026, 10, 27)}
    cfg = {"target": "2026-10-27", "family": "AME", "include_monthly": False, "underlying_product_id": 437}
    try:
        _candidates(_stub_get(monthly_dates), [AME_GROUP], cfg)
        assert False, "expected RuntimeError - monthly excluded"
    except RuntimeError:
        pass


def test_monthly_dates_missing_group_returns_empty():
    def get(url):
        return [{"optionType": "MW1", "contractExpirations": []}]  # no AME group at all
    assert _monthly_dates(get, {"underlying_product_id": 437}) == {}


def test_settle_price_prefers_requested_month_falls_back_to_oi_heuristic():
    rows = [
        {"month": "OCT 26", "settle": "4,300.5", "openInterest": "5,000"},
        {"month": "DEC 26", "settle": "4,350.2", "openInterest": "300,000"},
    ]

    def get(url):
        return {"settlements": rows}

    # requested month found -> used even though it's not the highest-OI one
    price, month = _settle_price(get, {"underlying_product_id": 437, "price_month": "OCT 26"}, "20260918")
    assert (price, month) == (4300.5, "OCT 26")

    # requested month not found -> falls back to first month with OI > 10,000
    price, month = _settle_price(get, {"underlying_product_id": 437, "price_month": "NOPE 99"}, "20260918")
    assert (price, month) == (4350.2, "DEC 26")

    # no request at all -> same fallback
    price, month = _settle_price(get, {"underlying_product_id": 437}, "20260918")
    assert (price, month) == (4350.2, "DEC 26")


def test_settle_price_returns_none_on_error():
    def get(url):
        raise RuntimeError("network down")
    assert _settle_price(get, {"underlying_product_id": 437}, "20260918") == (None, None)
