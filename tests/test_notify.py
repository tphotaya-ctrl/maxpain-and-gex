"""Tests for notify.py - no real popup is shown; subprocess.run is monkeypatched."""
import subprocess

import notify


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
    out = capsys.readouterr().out
    assert "Title: Message" in out
