"""Pure-logic test for publish_report.py's HTML rendering - no workbook, no git, no network."""
from datetime import datetime

from publish_report import render_html

MP_HEADER = ["วันที่ข้อมูล", "สัญญา", "DTE", "ราคา", "Max Pain", "ห่างจากราคา %",
             "ก้นหุบ (จำนวน strike)", "P/C Ratio", "ครอบคลุม CME %", "รายงาน", "บันทึกเมื่อ",
             "Futures ที่ใช้เป็นราคา"]
MP_ROW = [datetime(2026, 9, 25), "MW1 Week 4 - SEP 2026 (U26)", 3, 4321.2, 4330, 0.00204,
          3, 1.107, 1, "PRELIMINARY", datetime(2026, 9, 27, 16, 40, 48), "DEC 26"]

GEX_HEADER = ["วันที่ข้อมูล", "สัญญา", "ราคา", "รวม Call Gamma", "รวม Put Gamma", "NET GEX",
              "สถานะ", "Gamma Flip (ใกล้ราคา)", "Strike gamma สูงสุด", "บันทึกเมื่อ"]
GEX_ROW = [datetime(2026, 9, 25), "G4MU6", 4321.2, 488, 486, 2,
           "Positive GEX (นิ่ง)", 4475, 4325, datetime(2026, 9, 27, 16, 41, 35)]


def test_render_html_includes_latest_max_pain_and_gex_values():
    html = render_html(MP_HEADER, [MP_ROW], GEX_HEADER, [GEX_ROW], datetime(2026, 9, 27, 17, 0))
    assert "4330" in html                      # Max Pain strike
    assert "MW1 Week 4 - SEP 2026 (U26)" in html
    assert "4475" in html                      # Gamma Flip
    assert "Positive GEX" in html
    assert html.startswith("<!doctype html>") and html.rstrip().endswith("</html>")


def test_render_html_handles_empty_history():
    html = render_html(MP_HEADER, [], GEX_HEADER, [], datetime(2026, 9, 27))
    assert "<html" in html
    assert "<table>" in html  # empty tables still render, just with no <tr> rows
