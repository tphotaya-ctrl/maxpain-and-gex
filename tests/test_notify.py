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


def _no_msg(cmd, **kw):
    raise FileNotFoundError("msg.exe not found")  # Windows Home has no msg.exe


def test_notify_falls_back_to_powershell_popup_without_msg_exe(monkeypatch):
    popens = []
    monkeypatch.setattr(subprocess, "run", _no_msg)
    monkeypatch.setattr(subprocess, "Popen", lambda cmd, **kw: popens.append((cmd, kw["env"])))
    notify.notify("Title", "ข้อความ 'quoted'")
    assert len(popens) == 1
    cmd, env = popens[0]
    assert cmd[0] == "powershell"
    # text travels via env, never interpolated into the command line
    assert env["NOTIFY_TEXT"] == "ข้อความ 'quoted'" and env["NOTIFY_TITLE"] == "Title"
    assert "ข้อความ" not in " ".join(cmd)


def test_notify_falls_back_to_print_on_failure(monkeypatch, capsys):
    def boom(cmd, **kw):
        raise OSError("no powershell either")
    monkeypatch.setattr(subprocess, "run", _no_msg)
    monkeypatch.setattr(subprocess, "Popen", boom)
    notify.notify("Title", "Message")  # must not raise
    out = capsys.readouterr().out
    assert "Title: Message" in out
