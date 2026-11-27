"""
tests/test_v132_fixes.py
========================
v1.3.2 — ล็อกข้อที่แก้จากการรีวิว backend (2026-11-27):
- กดหยุดตอนไม่ได้รันต้องไม่ค้าง "กำลังหยุด"
- Force Close แล้วเริ่มรอบใหม่ทันที: รอบเก่าต้องไม่ฟื้นมาทำงานต่อ / ไม่ปิด CropWat ของผู้ใช้
- หน้าต่างที่พังติดกันด้วยสาเหตุฝั่ง CropWat ต้องถูกเปิดใหม่แล้วทำวันที่พังซ้ำ
- config.json เสียต้องไม่ทำให้โปรแกรมเปิดไม่ขึ้น, ช่วงปีถูกบีบ
- สคริปต์อัปเดตใช้ได้กับโฟลเดอร์ชื่อภาษาไทย
- API รับเฉพาะคำขอจากหน้าโปรแกรมเอง
"""

import threading
import time

import config
import desktop_session
import runner
import updater
from automation.cropwat_engine import CandidateRunResult, YearRunResult
from config import Settings
from models import OverallRunState, YearRunStatus
from state import RunState, run_state
from tests.test_parallel_windows import (  # noqa: F401 -- fake_env เป็น fixture
    _FakeEngine,
    _FakeSession,
    _settings,
    fake_env,
)


def setup_function():
    run_state.bind_thread(None)
    run_state.end_run()


# ---------------------------------------------------------------- สถานะ / หยุด

def test_stop_while_idle_does_nothing():
    s = RunState()
    s.request_stop()
    assert s.snapshot().overall_state == OverallRunState.IDLE
    assert s.is_stop_requested() is False


def test_stop_while_running_still_works():
    s = RunState()
    s.begin_run()
    s.request_stop()
    assert s.snapshot().overall_state == OverallRunState.STOPPING
    assert s.is_stop_requested() is True


def test_stale_thread_cannot_touch_new_run():
    s = RunState()
    old = s.begin_run()
    s.bind_thread(old)
    s.cancel_run()  # Force Close
    assert s.snapshot().overall_state == OverallRunState.IDLE
    assert s.is_stale() and s.is_stop_requested()

    result = {}

    def _new_run():
        result["id"] = s.begin_run()

    t = threading.Thread(target=_new_run)
    t.start()
    t.join()
    # thread เก่า (ยังผูกเลขรอบเดิม): ยังต้องหยุด และแก้สถานะ/จบรอบใหม่ไม่ได้
    assert s.is_stop_requested() is True
    s.set_year_status(1990, YearRunStatus.ERROR, error_message="จากรอบเก่า")
    s.end_run()
    snap = s.snapshot()
    assert snap.overall_state == OverallRunState.RUNNING
    assert all(y.year != 1990 for y in snap.years)
    s.bind_thread(None)


# ---------------------------------------------------------------- Force Close

def test_force_close_then_restart_old_run_does_not_resume(fake_env, monkeypatch):
    closed_user_cropwat = []
    monkeypatch.setattr(runner, "force_close_cropwat", lambda: closed_user_cropwat.append(1) or 1)
    _FakeEngine.task_seconds = 0.3
    s = _settings(fake_env, 1)

    run_state.init_years(1990, 1993)
    old = threading.Thread(target=runner._run_years, args=([1990, 1991, 1992, 1993], s))
    old.start()
    time.sleep(0.45)  # กำลังทำปี 1990 อยู่
    runner.force_reset()
    assert runner.is_run_active() is False
    # รันโหมดเดสก์ท็อปซ่อนอยู่ — ห้ามไปปิด CropWat บนจอหลักของผู้ใช้
    assert closed_user_cropwat == []

    run_state.init_years(2000, 2000)
    new = threading.Thread(target=runner._run_years, args=([2000], s))
    new.start()
    old.join(15)
    new.join(15)
    assert not old.is_alive() and not new.is_alive()

    years_done = {y for y, _d in _FakeEngine.dates}
    assert 2000 in years_done
    assert not years_done & {1991, 1992, 1993}, "รอบเก่ากลับมาทำงานต่อหลัง Force Close"
    snap = run_state.snapshot()
    assert snap.overall_state == OverallRunState.IDLE
    assert {y.year: y.status for y in snap.years} == {2000: YearRunStatus.DONE}


def test_force_close_when_not_running_hidden_still_closes_visible_cropwat(monkeypatch):
    monkeypatch.setattr(runner, "force_close_cropwat", lambda: 1)
    assert runner.force_reset() == 1


# ---------------------------------------------------------------- ตัวตัดวงจร

class _PoisonEngine:
    """CropWat ตัวแรก (pid ต่ำสุด) พังทุกวันปลูกด้วยสาเหตุ kind — ตัวที่เปิดใหม่ทำได้ปกติ"""

    kind = "StepTimeoutError"
    done: list = []
    poisoned_pid = None

    def __init__(self, background_mode=False, speed_multiplier=1.0):
        self.pause_check = None
        self.error_shots_dir = None

    def connect(self, process=None):
        self.pid = process
        if _PoisonEngine.poisoned_pid is None:
            _PoisonEngine.poisoned_pid = process

    def ensure_crop_soil_open(self, crop, soil):
        pass

    def run_year(self, year, tasks, climate_file_for, rain_file_for, export_dir,
                 screenshot_dir, on_candidate_done, should_stop):
        results = []
        for t in tasks:
            if should_stop():
                return YearRunResult(year=year, ok=False, candidates=results, stopped=True)
            time.sleep(0.01)
            if self.pid == _PoisonEngine.poisoned_pid:
                r = CandidateRunResult(year=year, planting_date=t.planting_date, ok=False,
                                       error_message="พัง", error_kind=_PoisonEngine.kind)
            else:
                r = CandidateRunResult(year=year, planting_date=t.planting_date, ok=True)
                _PoisonEngine.done.append((year, t.planting_date))
            results.append(r)
            on_candidate_done(r)
        failed = [c for c in results if not c.ok]
        return YearRunResult(year=year, ok=not failed, candidates=results,
                             error_message="มีวันปลูกพัง" if failed else None)


def _poison_run(fake_env, monkeypatch, kind):
    _PoisonEngine.kind = kind
    _PoisonEngine.done = []
    _PoisonEngine.poisoned_pid = None
    monkeypatch.setattr(runner, "CropWatEngine", _PoisonEngine)
    s = _settings(fake_env, 1, days=(1, 5, 10, 15, 20))
    run_state.init_years(1990, 1991)
    runner._run_years([1990, 1991], s)
    return {y.year: y.status for y in run_state.snapshot().years}


def test_window_failing_repeatedly_is_replaced_and_dates_redone(fake_env, monkeypatch):
    statuses = _poison_run(fake_env, monkeypatch, "StepTimeoutError")
    assert len(_FakeSession.instances) == 2  # เปิด CropWat ใหม่ 1 ครั้ง
    assert len(_PoisonEngine.done) == len(set(_PoisonEngine.done)) == 10  # 2 ปี × 5 วัน ครบ ไม่ซ้ำ
    assert statuses == {1990: YearRunStatus.DONE, 1991: YearRunStatus.DONE}


def test_data_errors_do_not_restart_cropwat(fake_env, monkeypatch):
    statuses = _poison_run(fake_env, monkeypatch, "CropWatReportedError")
    assert len(_FakeSession.instances) == 1
    assert statuses == {1990: YearRunStatus.ERROR, 1991: YearRunStatus.ERROR}


# ---------------------------------------------------------------- config

def test_corrupt_config_falls_back_to_defaults_and_keeps_copy(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DATA_DIR", tmp_path)
    monkeypatch.setattr(config, "CONFIG_PATH", tmp_path / "config.json")
    (tmp_path / "config.json").write_text('{"input_dir": "D:/x", "outp', encoding="utf-8")

    loaded = config.load_settings()
    assert loaded == Settings()
    kept = list(tmp_path.glob("config.corrupt-*.json"))
    assert len(kept) == 1 and "D:/x" in kept[0].read_text(encoding="utf-8")

    loaded.input_dir = "D:/data"
    config.save_settings(loaded)
    assert config.load_settings().input_dir == "D:/data"
    assert not (tmp_path / "config.json.tmp").exists()


def test_year_range_is_clamped():
    assert config.year_range(Settings(default_start_year=1981, default_end_year=1983)) == [1981, 1982, 1983]
    assert len(config.year_range(Settings(default_start_year=1981, default_end_year=20250))) == 2100 - 1981 + 1
    assert config.year_range(Settings(default_start_year=2000, default_end_year=1990)) == []


# ---------------------------------------------------------------- updater

def test_update_script_is_ascii_and_swaps_file_in_thai_folder(tmp_path, monkeypatch):
    updater._UPDATER_BAT.encode("ascii")
    assert 'start "" "%CWAR_TARGET_EXE%" --updated' in updater._UPDATER_BAT
    # ตัดบรรทัดเปิดโปรแกรมออกเฉพาะในเทสต์ (ไม่เปิดไฟล์ปลอมเป็นโปรแกรม)
    monkeypatch.setattr(
        updater, "_UPDATER_BAT",
        updater._UPDATER_BAT.replace('start "" "%CWAR_TARGET_EXE%" --updated\n', ""),
    )
    folder = tmp_path / "โปรแกรม CropWat ของฉัน"
    folder.mkdir()
    target = folder / "CropWatAutoRunner.exe"
    new = folder / "CropWatAutoRunner.update.exe"
    target.write_bytes(b"old")
    new.write_bytes(b"new version")

    proc = updater._spawn_swap_script(new, target)
    proc.wait(timeout=30)  # รหัสออกไม่มีความหมาย (สคริปต์ลบตัวเองเป็นคำสั่งสุดท้าย)
    assert target.read_bytes() == b"new version"
    assert not new.exists()
    assert not (folder / "cropwat_update.bat").exists()


# ---------------------------------------------------------------- API รับเฉพาะหน้าโปรแกรม

def test_only_same_origin_local_requests_allowed():
    import app

    ok = app._local_request_allowed
    assert ok({"host": "127.0.0.1:8000"})  # ไม่มี Origin (เช่น instance ที่สองของโปรแกรม)
    assert ok({"host": "127.0.0.1:8000", "origin": "http://127.0.0.1:8000"})
    assert ok({"host": "localhost:8765", "origin": "http://localhost:8765"})
    assert not ok({"host": "127.0.0.1:8000", "origin": "https://example.com"})
    assert not ok({"host": "127.0.0.1:8000", "origin": "http://127.0.0.1:5500"})
    assert not ok({"host": "evil.example:8000"})  # DNS rebinding
    assert not ok({"host": "127.0.0.1:8000", "origin": "null"})


def test_desktop_session_stop_not_called_from_force_reset(fake_env, monkeypatch):
    """force_reset ใช้ kill_process (ข้าม thread ได้) — stop() เป็นหน้าที่ของ thread เจ้าของ"""
    monkeypatch.setattr(runner, "force_close_cropwat", lambda: 0)
    _FakeEngine.task_seconds = 0.2
    s = _settings(fake_env, 1)
    run_state.init_years(1990, 1991)
    t = threading.Thread(target=runner._run_years, args=([1990, 1991], s))
    t.start()
    time.sleep(0.3)
    sess = _FakeSession.instances[0]
    assert runner.force_reset() == 1
    assert sess.killed.is_set()
    t.join(15)
    assert not t.is_alive() and sess.stopped
    assert desktop_session.HiddenDesktopSession is _FakeSession


# ---------------------------------------------------------------- ไฟล์ที่ใช้ของแต่ละวันปลูก

def test_run_records_files_used_per_planting_date(fake_env):
    s = _settings(fake_env, 2)
    run_state.init_years(1990, 1991)
    runner._run_years([1990, 1991], s)

    used = runner._load_files_used(s)
    assert set(used) == {(1990, "0401"), (1990, "0415"), (1991, "0401"), (1991, "0415")}
    e = used[(1991, "0415")]
    assert e["source"] == "run" and e["at"]
    assert e["climate"].endswith("1991_4.PED") and e["crop"].endswith("c.CRO")


def test_files_used_estimated_for_results_from_older_versions(tmp_path):
    """ผลที่รันก่อนมีการจด: ประเมินจากไฟล์ในโฟลเดอร์ต้นทางด้วยกติกาเดียวกับตอนรัน + บอกว่าประเมิน"""
    src = tmp_path / "in"
    (src / "Clim_111111").mkdir(parents=True)
    (src / "Rain_222222").mkdir()
    for name in ("111111clim_1990mar.PED", "111111clim_1990apr.PED"):
        (src / "Clim_111111" / name).write_text("x")
    (src / "Rain_222222" / "rain_222222_1990apr.CRD").write_text("x")
    (src / "MAIZE.CRO").write_text("x")
    (src / "sand.SOI").write_text("x")
    out = tmp_path / "out"
    (out / "txt").mkdir(parents=True)
    (out / "txt" / "1990_0415.txt").write_text("x")
    s = Settings(input_dir=str(src), output_dir=str(out))

    used = runner._load_files_used(s)
    e = used[(1990, "0415")]
    assert e["source"] == "estimate"
    assert e["climate"].endswith("111111clim_1990apr.PED")
    assert e["rain"].endswith("rain_222222_1990apr.CRD") and e["crop"].endswith("MAIZE.CRO")
