"""
tests/test_error_dialog_pattern.py
====================================
ยืนยันจาก screenshot ผู้ใช้จริง (v0.5.28): CropWat โยน runtime error แบบ Delphi
ดิบๆ ("Access violation at address 00000000...") ผ่าน dialog ที่ title เป็นชื่อ
โปรแกรมเอง "FAO CROPWAT 8.0 for Windows" — ไม่ใช่ "Error"/"Warning" เลย เดิม
ERROR_DIALOG.title_re จับไม่ได้ ปล่อยให้ dialog นี้ค้างอยู่แบบไม่มีใครรู้จัก โดน
background watcher เหวี่ยงออกนอกจอไปแบบมองไม่เห็น (ยังเป็น modal บล็อก CropWat
อยู่) นี่คือสาเหตุจริงของปัญหา "CropWat ค้างสนิท" ที่เคยรายงานมาก่อนหน้านี้

test นี้ล็อกว่า pattern ต้องจับทั้ง 3 รูปแบบ title ที่เจอมาแล้วจริง กัน regression
กลับไปแคบเกินไปอีก"""

import re

import cropwat_controls as controls


def test_matches_known_error_dialog_title():
    assert re.search(controls.ERROR_DIALOG.title_re, "Error")


def test_matches_known_warning_dialog_title():
    assert re.search(controls.ERROR_DIALOG.title_re, "Warning")


def test_matches_fao_cropwat_crash_dialog_title():
    """v0.5.28 — ยืนยันจาก screenshot จริง: 'Access violation at address
    00000000. Read of address 00000000.' โผล่ผ่าน dialog title นี้"""
    assert re.search(controls.ERROR_DIALOG.title_re, "FAO CROPWAT 8.0 for Windows")


def test_does_not_match_unrelated_titles():
    assert not re.search(controls.ERROR_DIALOG.title_re, "Daily rain - untitled")
    assert not re.search(controls.ERROR_DIALOG.title_re, "Daily ETo Penman-Monteith - untitled")
