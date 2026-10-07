"""Shared test setup: never let a test reach the real Telegram bot.

Once telegram.json exists on the machine (after `telegram_report.py --setup`), any test that
goes through notify() would otherwise send a real message to the user's phone.
"""
import pytest

import telegram_report


@pytest.fixture(autouse=True)
def no_real_telegram(monkeypatch, tmp_path):
    monkeypatch.setattr(telegram_report, "SECRETS", tmp_path / "telegram.json")  # absent
