"""
tests/test_runner_active_state.py
===================================
runner.is_run_active() เดิมยึด _active_thread.is_alive() — บั๊กจริงที่ผู้ใช้เจอ:
CropWat ค้างสนิท (dialog ที่ automation จำไม่ได้โดนเหวี่ยงออกนอกจอไปแบบมองไม่เห็น)
ทำให้ background thread ค้างตลอดไปแต่ยัง "alive" เลยกดเริ่มรันใหม่ไม่ได้แม้ปิด-เปิด
CropWat ใหม่แล้ว ต้องปิด-เปิดโปรแกรมทั้งตัวถึงจะรีเซ็ตได้ — v1.0 เปลี่ยนไปยึด
run_state.overall_state แทน (ควบคุมชัดเจนผ่าน begin_run/end_run) ตัดขาดจากสุขภาพ
ของ thread เดิม เพื่อให้ปุ่ม "ปิด CropWat ฉุกเฉิน" (runner.force_reset) คืนสถานะ
IDLE ได้ทันทีโดยไม่ต้องรอ thread เก่า"""

import runner
from state import run_state


def setup_function():
    # run_state เป็น singleton โมดูลระดับบน — รีเซ็ตกลับ IDLE ก่อนแต่ละ test กันชน
    run_state.end_run()


def test_idle_by_default():
    assert runner.is_run_active() is False


def test_active_after_begin_run():
    run_state.begin_run()
    assert runner.is_run_active() is True
    run_state.end_run()


def test_force_reset_clears_active_state_without_needing_thread_to_die(monkeypatch):
    """จำลองสถานการณ์บั๊กจริง: begin_run() แล้วไม่มีใครเรียก end_run() เลย (เหมือน
    thread ค้างอยู่ในนั้นตลอดไป) force_reset() ต้องคืน IDLE ได้ทันทีโดยไม่ต้องพึ่ง
    thread เดิม — mock force_close_cropwat กันไม่ให้ test ไป taskkill ของจริง"""
    run_state.begin_run()
    assert runner.is_run_active() is True

    monkeypatch.setattr(runner, "force_close_cropwat", lambda: 1)
    killed = runner.force_reset()

    assert killed == 1
    assert runner.is_run_active() is False
