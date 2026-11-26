"""
tests/test_month_switch_retry.py
=================================
v1.2.2 — "วันที่ 1 ชอบรันไม่ติด" (ผู้ใช้ 2026-11-26): วันแรกของเดือนคือวันเดียว
ที่ต้องสลับไฟล์ climate/rain และขั้นนั้นเคยไม่มีการลองใหม่เลย — ล็อกพฤติกรรม
(ไม่แตะ CropWat จริง: mock ขั้นเปิดไฟล์/คำนวณ)

(1) เปิดไฟล์พลาดชั่วคราว → เก็บกวาดแล้วลองใหม่จนผ่าน วันปลูกนั้นสำเร็จ
(2) พลาดครบทุกรอบ → วันปลูกนั้น fail พร้อมข้อความ error ตัวจริง
(3) FileNotFoundError (ไม่มีไฟล์จริง) → ไม่ลองซ้ำ
(4) _wait_for_file เรียกผ่าน instance ได้ (เดิม staticmethod เรียก self → NameError)
"""

from datetime import date
from pathlib import Path

import pytest

from automation.cropwat_engine import CandidateRunResult, CropWatEngine, PlantingDateTask


def _engine(monkeypatch, open_failures: int, exc_type=RuntimeError):
    engine = CropWatEngine(speed_multiplier=0)
    calls = {"open": 0, "reset": 0, "candidate": 0}

    def fake_open_climate(path):
        calls["open"] += 1
        if calls["open"] <= open_failures:
            raise exc_type("There are 2 elements that match the criteria TDayEToPMForm")

    def fake_candidate(year, task, export_dir, screenshot_dir=None):
        calls["candidate"] += 1
        return CandidateRunResult(year=year, planting_date=task.planting_date, ok=True)

    monkeypatch.setattr(engine, "open_climate_file", fake_open_climate)
    monkeypatch.setattr(engine, "open_rain_file", lambda path: None)
    monkeypatch.setattr(
        engine, "_reset_climate_rain_windows", lambda: calls.__setitem__("reset", calls["reset"] + 1)
    )
    monkeypatch.setattr(engine, "run_candidate_planting_date", fake_candidate)
    return engine, calls


def _run(engine):
    return engine.run_year(
        year=2020,
        tasks=[PlantingDateTask(planting_date=date(2020, 10, 1))],
        climate_file_for=lambda m: Path("clim.ped"),
        rain_file_for=lambda m: Path("rain.crd"),
        export_dir=Path("."),
    )


def test_transient_open_failure_is_retried(monkeypatch):
    engine, calls = _engine(monkeypatch, open_failures=2)
    result = _run(engine)
    assert result.ok
    assert calls["open"] == 3
    assert calls["reset"] == 2  # เก็บกวาดก่อนลองใหม่ทุกรอบ
    assert calls["candidate"] == 1


def test_persistent_open_failure_reports_real_error(monkeypatch):
    engine, calls = _engine(monkeypatch, open_failures=99)
    result = _run(engine)
    assert not result.ok
    assert calls["open"] == CropWatEngine._MAX_OPEN_ATTEMPTS
    assert calls["candidate"] == 0
    assert "TDayEToPMForm" in result.candidates[0].error_message


def test_missing_file_is_not_retried(monkeypatch):
    engine, calls = _engine(monkeypatch, open_failures=99, exc_type=FileNotFoundError)
    result = _run(engine)
    assert not result.ok
    assert calls["open"] == 1
    assert calls["reset"] == 0


def test_wait_for_file_polls_without_name_error(tmp_path):
    engine = CropWatEngine(speed_multiplier=0)
    assert engine._wait_for_file(tmp_path / "missing.txt", timeout=0.05) is False
    f = tmp_path / "x.txt"
    f.write_text("ok")
    assert engine._wait_for_file(f, timeout=0.05) is True


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
