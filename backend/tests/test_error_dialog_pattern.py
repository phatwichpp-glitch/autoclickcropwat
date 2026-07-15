"""
tests/test_error_dialog_pattern.py
====================================
ยืนยันจาก screenshot ผู้ใช้จริง (v0.5.28): CropWat โยน runtime error แบบ Delphi
ดิบๆ ("Access violation at address 00000000...") ผ่าน dialog ที่ title เป็นชื่อ
โปรแกรมเอง "FAO CROPWAT 8.0 for Windows" — เดิม title_re จับไม่ได้

v0.5.33 REGRESSION FIX — ยืนยันจากการ probe เครื่องจริง: การเพิ่ม "FAO CROPWAT"
เข้า title_re ไปจับโดนหน้าต่างซ่อนภายในของ Delphi (class TApplication, title
"FAO CROPWAT 8.0 for Windows", visible ตลอดเวลาที่โปรแกรมเปิด) ด้วย — ทำให้ทันที
ที่ error dialog จริงโผล่ จะมี 2 หน้าต่างตรง pattern พร้อมกัน เข้าเคส ambiguous
รายงาน error ได้แต่ไม่เคยกดปิด dialog จริง จน dialog ค้างบล็อกทุกอย่าง (ต้นเหตุ
จริงของ 30/30 ล้มเหลว) แก้ด้วยการกรอง "class ของหน้าต่าง" ควบคู่ title เสมอ

test นี้ล็อก: (1) title ยังจับ 3 รูปแบบที่เจอจริง (2) มี dialog_class_names ที่
"ไม่รวม TApplication" — กันไม่ให้ใครเผลอกลับไปจับ title อย่างเดียวอีก
"""

import re

import cropwat_controls as controls


def test_matches_known_error_dialog_title():
    assert re.match(controls.ERROR_DIALOG.title_re, "Error")


def test_matches_known_warning_dialog_title():
    assert re.match(controls.ERROR_DIALOG.title_re, "Warning")


def test_matches_fao_cropwat_crash_dialog_title():
    """v0.5.28 — ยืนยันจาก screenshot จริง: 'Access violation at address
    00000000...' โผล่ผ่าน native MessageBox ที่ title เป็นชื่อโปรแกรม"""
    assert re.match(controls.ERROR_DIALOG.title_re, "FAO CROPWAT 8.0 for Windows")


def test_does_not_match_unrelated_titles():
    assert not re.match(controls.ERROR_DIALOG.title_re, "Daily rain - untitled")
    assert not re.match(controls.ERROR_DIALOG.title_re, "Daily ETo Penman-Monteith - untitled")


def test_dialog_class_names_defined_and_excludes_tapplication():
    """v0.5.33 — หัวใจของ regression fix: ต้องกรองด้วย class ของ "dialog จริง"
    เท่านั้น (TMessageForm/#32770) — TApplication (หน้าต่างซ่อนภายในของ Delphi
    ที่ title ก็ตรง pattern "FAO CROPWAT") ต้องไม่อยู่ในรายชื่อ ไม่งั้นจะกลับไป
    เจอบั๊ก ambiguous เดิมที่ทำให้ error dialog ไม่เคยถูกกดปิด"""
    class_names = controls.ERROR_DIALOG.dialog_class_names
    assert "TMessageForm" in class_names
    assert "#32770" in class_names
    assert "TApplication" not in class_names
