"""
tests/test_config_defaults.py
===============================
v0.7.0 — ล็อกการตัดสินใจเชิงสถาปัตยกรรม: โหมดเดสก์ท็อปซ่อนเป็น default และระบบ
เบื้องหลังเก่า (background_mode + watcher เหวี่ยงหน้าต่าง + click-shield ที่เป็น
ต้นเหตุ CropWat crash ตอนสลับไฟล์ข้ามเดือน/ปี — ยืนยันจากการวิเคราะห์เครื่องจริง)
ถูกถอดออกทั้งหมดแล้ว — กันการเผลอใส่กลับ/เปลี่ยน default โดยไม่รู้ตัว
"""

from config import Settings


def test_hidden_desktop_mode_is_default_on():
    assert Settings().hidden_desktop_mode is True


def test_legacy_config_with_background_mode_key_still_loads():
    """config.json เก่าของผู้ใช้ที่อัปเดตข้ามเวอร์ชันยังมี key 'background_mode'
    ค้างอยู่ — ต้องโหลดได้ปกติ (pydantic ข้าม key ที่ไม่รู้จัก) ไม่ crash ตอนเปิด"""
    s = Settings(input_dir="x", background_mode=True)
    assert s.input_dir == "x"
    assert not hasattr(s, "background_mode") or "background_mode" not in type(s).model_fields


def test_old_watcher_shield_machinery_removed():
    """ระบบเก่าที่ทำให้ crash ต้องไม่มีอยู่ใน engine อีก"""
    from automation.cropwat_engine import CropWatEngine

    assert not hasattr(CropWatEngine, "start_background_watcher")
    assert not hasattr(CropWatEngine, "_enter_protected_mode")
    assert not hasattr(CropWatEngine, "_exit_protected_mode")
