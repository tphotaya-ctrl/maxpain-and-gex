"""daily_summary.format_summary / rule_lines - built from plain Log rows, no workbook."""
from datetime import datetime

from daily_summary import format_summary, rule_lines

MP = [datetime(2026, 9, 30), "BB1 Week 1 - OCT 2026 (V26)", 1, 4186.7, 4230, 0.01034,
      3, 1.1, 1, "PRELIMINARY", datetime(2026, 10, 1, 8, 31), "DEC 26"]
GEX = [datetime(2026, 9, 30), "G1RV6", 4186.7, 310, 469, -159, "Negative GEX (วิ่ง)", None, 4200,
       datetime(2026, 10, 1, 8, 32), 4230, 4175, 4225, 0.4087]


def test_summary_has_the_reading_and_armed_rules():
    text = format_summary(MP, GEX, 0.02, 0.05, change=7.0)
    assert "ข้อมูล 2026-09-30 (PRELIMINARY)" in text
    assert "ราคา 4,186.7 (DEC 26) เปลี่ยน +7.0" in text
    assert "Max Pain 4,230 (+1.03%, โซนกลาง)" in text
    assert "NET -159" in text and "ความชัดเจน 0.41" in text
    assert "Put Wall 4,175 - Call Wall 4,230 - Pin 4,225" in text
    # same day as the 1 Oct paper trade: negative mode, price up -> R2 long, stop at put wall
    assert "R2: Long ตามราคาที่ขึ้น +7.0, stop 4,175" in text
    assert "R3: Long ไปหา Max Pain 4,230" in text  # 1.03% >= 0.5%
    assert "R1: ไม่เข้า" in text


def test_summary_without_gex_says_so_and_disarms_gex_rules():
    text = format_summary(MP, None, 0.02, 0.05)
    assert "GEX: ไม่มีข้อมูลวันนี้" in text
    assert "R1: ไม่เข้า" in text and "R2: ไม่เข้า" in text


def test_r1_armed_in_positive_mode_shows_range_and_midpoint():
    r1, r2, r3 = rule_lines(4200, 4202, "Positive GEX (นิ่ง)", 4250, 4150, 5.0)
    assert r1.startswith("R1: กรอบ 4,150-4,250 - เปิดต่ำกว่า 4,200 ซื้อ (stop 4,150)")
    assert r2.startswith("R2: ไม่เข้า")
    assert r3.startswith("R3: ไม่เข้า")  # 2 points is < 0.5%


def test_r2_short_without_a_wall_above_has_no_stop():
    _, r2, _ = rule_lines(4200, None, "Negative GEX (วิ่ง)", None, None, -12.5)
    assert r2 == "R2: Short ตามราคาที่ลง -12.5, ไม่มี stop"
