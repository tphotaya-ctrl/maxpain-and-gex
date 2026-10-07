"""Tests for telegram_report.py - no network: the HTTP call is monkeypatched."""
import json
from datetime import date

import telegram_report as tr

MP = {"contract": "AB1 Week 1 - OCT 2026 (V26)", "trade_date": date(2026, 10, 2),
      "report": "FINAL", "price": 4162.3, "max_pain": 4200}
GEX = {"code": "G1TV6", "label": "Week 1 - OCT 2026", "trade_date": date(2026, 10, 2),
       "report": "FINAL", "price": 4162.3, "calls": 153, "puts": 209, "net": -56,
       "flip": None, "peak": 4175, "rows": [(4150, 1, 10), (4175, 21, 49), (4200, 31, 13)]}
XL = {"J5": "Negative GEX (วิ่ง)", "J9": 4200.0, "J10": 4100.0, "J11": 4175.0, "J18": 0.224,
      "J19": "พอใช้ได้", "J30": 4200.0, "J31": 4195.0, "J32": 4100.0, "J33": "", "J34": None}


def test_format_report_full():
    text = tr.format_report(MP, GEX, XL)
    assert text.splitlines()[0] == "GEX G1TV6 (Week 1 - OCT 2026) · ข้อมูล 2026-10-02 FINAL"
    assert "ราคา 4,162.3 · Max Pain 4,200 (+0.9%)" in text
    assert "NET GEX -56 → Negative GEX (วิ่ง) · Conviction 0.22 พอใช้ได้" in text
    assert "Call Wall 4,200 · Put Wall 4,100 · หนืดสุด 4,175" in text
    assert "Pins 4,200, 4,195, 4,100" in text  # blank J33/J34 left out


def test_format_report_without_excel_values_uses_python_totals():
    text = tr.format_report(MP, GEX, {})
    assert "NET GEX -56 → Negative GEX" in text
    assert "Call Wall" not in text and "หนืดสุด 4,175" in text and "Pins" not in text


def test_format_report_max_pain_only_when_gex_failed():
    text = tr.format_report(MP, None, {})
    assert "GEX ไม่ได้อัปเดต" in text and "Max Pain 4,200" in text


def test_send_is_noop_without_secrets(monkeypatch):
    calls = []
    monkeypatch.setattr(tr, "_call", lambda *a, **k: calls.append(a))
    assert tr.send_text("hi") is False and calls == []


def test_send_errors_never_raise(monkeypatch, tmp_path, capsys):
    secrets = tmp_path / "telegram.json"
    secrets.write_text(json.dumps({"token": "T", "chat_id": 1}))
    monkeypatch.setattr(tr, "SECRETS", secrets)

    def boom(*a, **k):
        raise OSError("network down")
    monkeypatch.setattr(tr, "_call", boom)
    assert tr.send_text("hi") is False
    assert "network down" in capsys.readouterr().out


def test_render_chart_writes_png(tmp_path):
    p = tr.render_chart(GEX["rows"], 4162.3, 4200, 4100, "t", tmp_path / "c.png")
    assert p.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"
