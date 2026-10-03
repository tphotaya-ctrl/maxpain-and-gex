"""Tests for notify.py - no real popup, no real Telegram/healthchecks request.

Every test points notify.SECRETS at a temp file, so a real secrets.json on the machine can
never make a test send an actual message."""
import json
import subprocess
import urllib.request

import pytest

import notify


@pytest.fixture(autouse=True)
def no_real_secrets(tmp_path, monkeypatch):
    monkeypatch.setattr(notify, "SECRETS", tmp_path / "secrets.json")
    return tmp_path / "secrets.json"


class _Resp:
    status = 200

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def _capture_urlopen(monkeypatch):
    sent = []

    def fake(url, data=None, timeout=None):
        sent.append((url, data))
        return _Resp()
    monkeypatch.setattr(urllib.request, "urlopen", fake)
    return sent


def test_notify_calls_msg_exe_with_title_and_message(monkeypatch):
    calls = []
    monkeypatch.setattr(subprocess, "run", lambda cmd, **kw: calls.append(cmd))
    notify.notify("Title", "Message")
    assert len(calls) == 1
    assert calls[0][0] == "msg"
    assert "Title: Message" in calls[0]


def test_notify_falls_back_to_print_on_failure(monkeypatch, capsys):
    def boom(cmd, **kw):
        raise FileNotFoundError("msg.exe not found")
    monkeypatch.setattr(subprocess, "run", boom)
    notify.notify("Title", "Message")  # must not raise
    assert "Title: Message" in capsys.readouterr().out


def test_telegram_not_configured_sends_nothing(monkeypatch):
    sent = _capture_urlopen(monkeypatch)
    assert notify.send_telegram("hi") is False
    assert notify.ping_health() is False
    assert sent == []


def test_notify_also_sends_telegram_when_configured(monkeypatch, no_real_secrets):
    no_real_secrets.write_text(json.dumps({"telegram_token": "T0K", "telegram_chat_id": "42"}))
    monkeypatch.setattr(subprocess, "run", lambda cmd, **kw: None)
    sent = _capture_urlopen(monkeypatch)
    notify.notify("GEX", "mode flipped")
    (url, data), = sent
    assert url == "https://api.telegram.org/botT0K/sendMessage"
    assert b"chat_id=42" in data and b"GEX%3A+mode+flipped" in data


def test_telegram_failure_never_raises(monkeypatch, no_real_secrets, capsys):
    no_real_secrets.write_text(json.dumps({"telegram_token": "T", "telegram_chat_id": "1"}))

    def down(*a, **k):
        raise OSError("no network")
    monkeypatch.setattr(urllib.request, "urlopen", down)
    assert notify.send_telegram("hi") is False
    assert "Telegram failed" in capsys.readouterr().out


def test_ping_health_ok_and_fail_urls(monkeypatch, no_real_secrets):
    no_real_secrets.write_text(json.dumps({"healthcheck_url": "https://hc-ping.com/abc/"}))
    sent = _capture_urlopen(monkeypatch)
    notify.ping_health()
    notify.ping_health(fail=True, body="log")
    assert [u for u, _ in sent] == ["https://hc-ping.com/abc", "https://hc-ping.com/abc/fail"]
