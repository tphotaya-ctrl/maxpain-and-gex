"""Tests for telegram_report.py - no network: the HTTP call is monkeypatched."""
import json
from datetime import date

import telegram_report as tr

MP = {"contract": "AB1 Week 1 - OCT 2026 (V26)", "trade_date": date(2026, 10, 2),
      "report": "FINAL", "price": 4162.3, "max_pain": 4200, "dte": 4, "zone": "กลาง",
      "call_total": 1000, "put_total": 1250,
      "strikes": [(4100, 10, 300), (4150, 50, 200), (4200, 400, 20), (4250, 300, 0), (4300, 120, 0)]}
GEX = {"code": "G1TV6", "label": "Week 1 - OCT 2026", "trade_date": date(2026, 10, 2),
       "report": "FINAL", "price": 4162.3, "calls": 153, "puts": 209, "net": -56,
       "flip": None, "peak": 4175,
       "rows": [(4100, 0, 39), (4150, 1, 10), (4155, 0, 0), (4160, 0, 0), (4165, 0, 0),
                (4170, 2, 3), (4175, 21, 49), (4180, 0, 0), (4200, 31, 13)]}
XL = {"J5": "Negative GEX (วิ่ง)", "J9": 4200.0, "J10": 4100.0, "J11": 4175.0, "J18": 0.224,
      "J19": "พอใช้ได้", "J27": "ไม่มีแรงดูด", "J30": 4200.0, "M30": "Call เด่น",
      "J31": 4195.0, "M31": "(mixed)", "J32": 4100.0, "M32": "Put เด่น", "J33": "", "J34": None}


def test_format_report_full():
    text = tr.format_report(MP, GEX, XL)
    assert text.splitlines()[0] == "📊 G1TV6 (Week 1 - OCT 2026) · ข้อมูล 2026-10-02 FINAL"
    assert "ราคา 4,162.3 · DTE 4" in text
    assert "NET GEX -56 → Negative GEX (วิ่ง) · Conviction 0.22 พอใช้ได้" in text
    assert "แรงดูด: ไม่มีแรงดูด" in text
    assert "Call Wall 4,200 (+37.7) · Put Wall 4,100 (-62.3)" in text
    assert "หนืดสุด 4,175 (+12.7)" in text
    # (mixed) side dropped, blank J33/J34 left out
    assert "Pins: 4,200 (+37.7) Call เด่น, 4,195 (+32.7), 4,100 (-62.3) Put เด่น" in text
    assert "Gamma หนาแน่นใกล้ราคา: เหนือ 4,175 (70), 4,200 (44) · ใต้ 4,150 (11), 4,100 (39)" in text
    assert "ช่องว่าง (gamma 0) ในกรอบ Wall: 4,155–4,165, 4,180" in text
    assert "ราคาอยู่ในกรอบ Put Wall–Call Wall (กว้าง 100)" in text
    assert "Max Pain 4,200 (+0.9%) · โซน กลาง" in text
    assert "OI รวม Call 1,000 / Put 1,250 · P/C 1.25" in text
    assert "Call OI สูงสุด: 4,200 (400), 4,250 (300), 4,300 (120)" in text
    assert "Put OI สูงสุด: 4,100 (300), 4,150 (200), 4,200 (20)" in text


def test_format_report_without_excel_values_uses_python_totals():
    text = tr.format_report(MP, GEX, {})
    assert "NET GEX -56 → Negative GEX" in text
    assert "อ่านค่าจาก Excel ไม่ได้" in text
    assert "Call Wall 4" not in text and "หนืดสุด 4,175" in text and "Pins:" not in text


def test_format_report_max_pain_only_when_gex_failed():
    text = tr.format_report(MP, None, {})
    assert "GEX ไม่ได้อัปเดต" in text and "Max Pain 4,200" in text and "— GEX —" not in text


def test_dense_strikes_ignores_zero_and_far_strikes():
    above, below = tr.dense_strikes(GEX["rows"], 4162.3, window=0.015)
    # ±1.5% of 4162.3 = ±62.4: top two by gross on each side, zero strikes never count
    assert above == [(4175, 70), (4200, 44)] and below == [(4150, 11), (4100, 39)]


def test_python_mode_matches_workbook_rule():
    # 2026-10-07 G1WV6: NET -10 over a near-balanced board - the sheet says "ไม่มีโหมด"
    rows = [(4185, 40, 140), (4190, 130, 30), (4200, 90, 10), (4180, 10, 80)]  # 10/350 < 0.05
    assert tr.python_mode(rows, -10) == "ไม่มีโหมด"
    assert tr.python_mode(GEX["rows"], -56) == "Negative GEX (วิ่ง)"


def test_reading_flags_weak_mode():
    assert "สัญญาณ GEX ไม่ชัด" in tr.reading("ไม่มีโหมด", 4187, 4200, 4185)[0]


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
    monkeypatch.setattr(tr.time, "sleep", lambda s: None)
    assert tr.send_text("hi") is False
    assert "network down" in capsys.readouterr().out


def test_render_chart_writes_png(tmp_path):
    p = tr.render_chart(GEX["rows"], 4162.3, 4200, 4100, "t", tmp_path / "c.png")
    assert p.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"


def test_card_html_orders_levels_and_marks_price():
    html = tr.card_html(MP, GEX, XL)
    # strikes top-down, the price row slotted between 4,175 (above) and 4,100 (below)
    order = [html.index(s) for s in ("4,200</td>", "4,195</td>", "4,175</td>", "ราคาปัจจุบัน", "4,100</td>")]
    assert order == sorted(order)
    assert "class='mode neg'" in html and "+37.7" in html and "-62.3" in html
    assert "Call Wall" in html and "Max Pain" in html and "Pin 1" in html
    assert "Call OI สูงสุด" in html and "P/C 1.25" in html
    assert "อ่านค่าจาก Excel ไม่ได้" not in html


def test_card_html_without_excel_warns_and_escapes():
    html = tr.card_html({**MP, "contract": "<b>x</b>"}, None, {})
    assert "GEX ไม่ได้อัปเดต" in html and "&lt;b&gt;x&lt;/b&gt;" in html
    html = tr.card_html(MP, GEX, {})
    assert "อ่านค่าจาก Excel ไม่ได้" in html and "Call Wall" not in html


def test_headline():
    assert tr.headline(MP, GEX, XL) == "G1TV6 · Negative GEX (วิ่ง) · Call Wall 4,200 / Put Wall 4,100 · Max Pain 4,200"


def test_send_album_posts_all_pictures_with_caption_on_first(monkeypatch, tmp_path):
    secrets = tmp_path / "telegram.json"
    secrets.write_text(json.dumps({"token": "T", "chat_id": 1}))
    monkeypatch.setattr(tr, "SECRETS", secrets)
    pics = []
    for i in range(3):
        p = tmp_path / f"{i}.png"; p.write_bytes(b"png"); pics.append(p)
    calls = []
    monkeypatch.setattr(tr, "_call", lambda token, method, data=None, files=None, timeout=20: calls.append((method, data, files)))
    assert tr.send_album(pics, "cap") is True
    method, data, files = calls[0]
    media = json.loads(data["media"])
    assert method == "sendMediaGroup" and len(media) == 3 and set(files) == {"p0", "p1", "p2"}
    assert media[0]["caption"] == "cap" and "caption" not in media[1]


def test_render_card_writes_png(tmp_path):
    pytest = __import__("pytest")
    try:
        p = tr.render_card(tr.card_html(MP, GEX, XL), tmp_path / "card.png")
    except Exception as e:  # no Chrome (e.g. CI)
        pytest.skip(f"Chrome unavailable: {e}")
    assert p.read_bytes()[:8] == b"\x89PNG\r\n\x1a\n"


def test_expected_trade_date_skips_weekend():
    assert tr.expected_trade_date(date(2026, 10, 9)) == date(2026, 10, 8)   # Fri -> Thu
    assert tr.expected_trade_date(date(2026, 10, 12)) == date(2026, 10, 9)  # Mon -> Fri
    assert tr.expected_trade_date(date(2026, 10, 11)) == date(2026, 10, 9)  # Sun -> Fri


def test_stale_note_and_card_banner():
    assert tr.stale_note(date(2026, 10, 8), date(2026, 10, 9)) == ""
    note = tr.stale_note(date(2026, 10, 7), date(2026, 10, 9))
    assert "2026-10-07" in note and "2026-10-08" in note
    assert "⏳" in tr.card_html(MP, GEX, XL, today=date(2026, 10, 6))   # MP/GEX hold 10-02
    assert "⏳" not in tr.card_html(MP, GEX, XL, today=date(2026, 10, 3))


def test_card_has_all_expiration_section():
    agg = {"codes": ["OG2V6", "OGX6"], "price": 4162.3, "mode": "Positive GEX (นิ่ง)", "net": 40,
           "calls": 140, "puts": 100, "conviction": 0.4, "call_wall": 4200, "put_wall": 4000,
           "flip": 4150, "peak": 4200}
    html = tr.card_html(MP, GEX, XL, agg=agg)
    assert "GEX รวม 2 สัญญา" in html and "OG2V6, OGX6" in html and "Gamma Flip" in html and "-12.3" in html
    assert "รวมทุกสัญญา: Positive GEX (นิ่ง)" in tr.headline(MP, GEX, XL, agg)
