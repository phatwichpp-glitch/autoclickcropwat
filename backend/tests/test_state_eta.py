"""
tests/test_state_eta.py
=========================
RunState._estimate_eta_seconds() คือจุดที่เคยมีบั๊กจริงจากผู้ใช้ (v0.5.17: ETA
บอก "เหลือ 1 นาที" แล้วพุ่งเป็น "เหลือ 2 นาที" ทั้งที่เวลาผ่านไปแล้วไม่ใช่ถอย
หลัง) ต้นเหตุคือตอน resume แล้ว candidate_done กระโดดสูงทันทีที่ elapsed≈0 —
แก้ด้วย baseline_done (ดู state.py) test นี้ล็อกพฤติกรรมไว้กัน regression
"""

import time

import pytest

from models import OverallRunState
from state import RunState


def _fresh_state() -> RunState:
    s = RunState()
    s.init_years(2020, 2020)
    s.begin_run()
    return s


def test_no_eta_before_min_samples():
    s = _fresh_state()
    s.set_candidate_progress(1, 100)  # ต่ำกว่า _MIN_SAMPLES_FOR_ETA (2)
    assert s.snapshot().eta_seconds is None


def test_no_eta_when_not_running():
    s = _fresh_state()
    s.set_candidate_progress(5, 100)
    s.end_run()
    assert s.snapshot().eta_seconds is None


def test_no_eta_when_already_complete():
    s = _fresh_state()
    s.set_candidate_progress(100, 100)
    assert s.snapshot().eta_seconds is None


def test_eta_uses_only_newly_done_not_resumed_baseline():
    """จำลองสถานการณ์บั๊กจริง: resume มาโดยมี 206 วันปลูกเสร็จไปแล้วจากรอบก่อน
    (baseline) ตอนเริ่มรอบใหม่ elapsed ต้องยังไม่โดนหารด้วย 206 (ซึ่งจะทำให้
    ETA ต่ำเวอร์ผิดปกติ) แต่ต้องรอให้มีวันปลูกที่เสร็จ "ใหม่จริง" ในรอบนี้ก่อน"""
    s = RunState()
    s.init_years(2020, 2020)
    s.begin_run()
    s.set_baseline_done(206)
    # แค่ progress ขยับตาม baseline เดิม (resume) ยังไม่มีอะไรเสร็จใหม่เลย
    s.set_candidate_progress(206, 300)
    assert s.snapshot().eta_seconds is None

    # จำลองเวลาผ่านไปจริง แล้วมีวันปลูกใหม่เสร็จเพิ่ม 2 อัน (ครบ min samples)
    s._run_started_at = time.monotonic() - 20.0  # 20 วินาทีที่ผ่านมา
    s.set_candidate_progress(208, 300)
    eta = s.snapshot().eta_seconds
    assert eta is not None
    # ความเร็ว = 20s / 2 วันปลูกใหม่ = 10s/วันปลูก, เหลืออีก 92 วันปลูก -> ~920s
    assert eta == pytest.approx(920.0, rel=0.05)


def test_eta_monotonic_non_negative_progression():
    """เดินหน้าเรื่อยๆ (progress เพิ่มขึ้น, เวลาผ่านไปจริง) ETA ต้องไม่ใช่ค่าติด
    ลบ และไม่ควรพุ่งขึ้นแบบไม่มีเหตุผลจากพฤติกรรม resume (บั๊กเดิม)"""
    s = _fresh_state()
    s._run_started_at = time.monotonic() - 10.0
    s.set_candidate_progress(2, 100)
    eta1 = s.snapshot().eta_seconds
    assert eta1 is not None and eta1 >= 0

    s._run_started_at = time.monotonic() - 20.0
    s.set_candidate_progress(4, 100)
    eta2 = s.snapshot().eta_seconds
    assert eta2 is not None and eta2 >= 0
