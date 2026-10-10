"""Shared test setup: never let a test reach the real Telegram bot or healthchecks URL.

Once secrets.json / telegram.json exist on the machine, any test that goes through notify()
or the phone report would otherwise send a real message to the user's phone. Tests that need
secrets point SECRETS at their own temp file, which overrides this.
"""
import pytest

import notify
import telegram_report


@pytest.fixture(autouse=True)
def no_real_telegram(monkeypatch, tmp_path):
    monkeypatch.setattr(notify, "SECRETS", tmp_path / "secrets.json")  # absent
    monkeypatch.setattr(telegram_report, "SECRETS", tmp_path / "telegram.json")  # absent
    for var in notify.ENV_SECRETS.values():  # the cloud jobs pass secrets as env vars
        monkeypatch.delenv(var, raising=False)
