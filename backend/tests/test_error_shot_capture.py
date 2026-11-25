"""
tests/test_error_shot_capture.py
=================================
v0.11.3 — ล็อกพฤติกรรมการประกอบภาพ error (ไม่แตะ CropWat จริง: mock ส่วนถ่ายภาพ)

(1) หน้าต่างหลักปกติ → วางกล่อง error ทับตรงกลาง
(2) หน้าต่างหลักถูกย่อ / เล็กกว่ากล่อง / หาไม่เจอ → เซฟแค่กล่อง error ให้ข้อความครบ
(3) ถ่ายกล่องค้างเกินเวลา → ไม่ถ่ายหน้าต่างหลักซ้ำ (UI thread เดียวกันจะค้างอีก)
(4) error 2 กล่องในวินาทีเดียวกัน → ชื่อไฟล์ไม่ทับกัน
"""

from PIL import Image

from automation import cropwat_engine as eng_mod
from automation.cropwat_engine import CropWatEngine

DLG = (300, 120)


def _engine(tmp_path, monkeypatch, *, main_hwnd=111, main_size=(800, 600),
            iconic=False, dlg_timed_out=False):
    engine = CropWatEngine()
    engine.error_shots_dir = tmp_path
    engine.current_candidate_label = "1990_0910"
    calls: list[int] = []

    def fake_shot(hwnd, scale=None):
        calls.append(hwnd)
        if hwnd == 999:  # กล่อง error
            if dlg_timed_out:
                return None, True
            return Image.new("RGB", DLG, (255, 0, 0)), False
        return Image.new("RGB", main_size, (0, 0, 255)), False

    monkeypatch.setattr(engine, "_find_main_hwnd", lambda: main_hwnd)
    monkeypatch.setattr(engine, "_printwindow_image", fake_shot)
    monkeypatch.setattr(CropWatEngine, "_window_dpi_scale", staticmethod(lambda h: 1.0))
    monkeypatch.setattr(eng_mod.win32gui, "IsIconic", lambda h: iconic)
    return engine, calls


def test_dialog_pasted_at_center_of_main_window(tmp_path, monkeypatch):
    engine, _ = _engine(tmp_path, monkeypatch)
    img = Image.open(engine._capture_error_dialog(999, "Error"))
    assert img.size == (800, 600)
    assert img.getpixel((400, 300)) == (255, 0, 0)  # กลางภาพเป็นกล่อง error
    assert img.getpixel((5, 5)) == (0, 0, 255)  # ขอบเป็นหน้าต่างหลัก


def test_minimized_main_window_saves_dialog_only(tmp_path, monkeypatch):
    engine, calls = _engine(tmp_path, monkeypatch, iconic=True)
    img = Image.open(engine._capture_error_dialog(999, "Error"))
    assert img.size == DLG
    assert calls == [999]


def test_main_window_smaller_than_dialog_saves_dialog_only(tmp_path, monkeypatch):
    engine, _ = _engine(tmp_path, monkeypatch, main_size=(160, 28))
    img = Image.open(engine._capture_error_dialog(999, "Error"))
    assert img.size == DLG


def test_main_window_not_found_still_saves_dialog(tmp_path, monkeypatch):
    engine, _ = _engine(tmp_path, monkeypatch, main_hwnd=None)
    img = Image.open(engine._capture_error_dialog(999, "Error"))
    assert img.size == DLG


def test_dialog_timeout_skips_main_window_capture(tmp_path, monkeypatch):
    engine, calls = _engine(tmp_path, monkeypatch, dlg_timed_out=True)
    assert engine._capture_error_dialog(999, "Error") is None
    assert calls == [999]


def test_two_shots_same_second_do_not_overwrite(tmp_path, monkeypatch):
    engine, _ = _engine(tmp_path, monkeypatch)
    monkeypatch.setattr(eng_mod.time, "strftime", lambda fmt: "142233")
    first = engine._capture_error_dialog(999, "Error")
    second = engine._capture_error_dialog(999, "Error")
    assert first != second
    assert len(list(tmp_path.glob("*.png"))) == 2
