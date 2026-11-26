"""
tests/test_parallel_windows.py
==============================
v1.3.0 — ล็อกพฤติกรรม "รันหลายหน้าต่าง CropWat พร้อมกัน" (ไม่เปิด CropWat จริง:
แทน HiddenDesktopSession + CropWatEngine ด้วยของปลอม)

- ทุกวันปลูกถูกทำครบ ไม่ซ้ำ แบ่งกันหลายหน้าต่างจริง และนับ progress ครบ
- หน้าต่างไหนเปิด CropWat ไม่ขึ้น หน้าต่างอื่นทำแทนให้ครบ / เปิดไม่ขึ้นเลย = error พร้อมเหตุผล
- ไม่เปิดหน้าต่างเกินจำนวนปีที่ต้องทำ, ชื่อเดสก์ท็อปซ่อนไม่ซ้ำกัน
- โหมดอัตโนมัติใช้จำนวนที่ system_capacity แนะนำ
- ปรับจำนวนหน้าต่างระหว่างรัน: ลด = หน้าต่างท้ายวางงานกลางปี ส่งวันที่เหลือให้หน้าต่างอื่น
  (ไม่มีวันปลูกหาย/ไม่ทำซ้ำ), เพิ่ม = เปิดหน้าต่างใหม่มาช่วย
"""

import threading
import time

import pytest

import desktop_session
import runner
import system_capacity
from automation.cropwat_engine import CandidateRunResult, YearRunResult
from config import MonthSelection, Settings
from models import YearRunStatus
from state import run_state


class _FakeSession:
    instances: list = []
    fail_names: set = set()
    by_pid: dict = {}
    lock = threading.Lock()

    def __init__(self, exe, desktop_name=None):
        self.desktop_name = desktop_name
        with _FakeSession.lock:
            self.pid = 1000 + len(_FakeSession.instances)
            _FakeSession.instances.append(self)
            _FakeSession.by_pid[self.pid] = self
        self.stopped = False
        self.killed = threading.Event()

    def bind_and_launch(self):
        if self.desktop_name in _FakeSession.fail_names:
            raise desktop_session.DesktopSessionError(f"เปิดไม่ขึ้น ({self.desktop_name})")
        return self.pid

    def stop(self):
        self.stopped = True
        self.killed.set()

    def kill_process(self):
        self.killed.set()


class _FakeEngine:
    runs: list = []  # (thread name, year)
    dates: list = []  # (year, planting_date) ที่ทำสำเร็จ
    lock = threading.Lock()
    task_seconds = 0.05
    hang_on: set = set()  # {(year, day)} ที่ CropWat จะค้างจนกว่าถูกปิด (ครั้งแรกเท่านั้น)
    hung_log: list = []

    def __init__(self, background_mode=False, speed_multiplier=1.0):
        self.pause_check = None
        self.error_shots_dir = None

    def connect(self, process=None):
        self.pid = process

    def ensure_crop_soil_open(self, crop, soil):
        pass

    def run_year(self, year, tasks, climate_file_for, rain_file_for, export_dir,
                 screenshot_dir, on_candidate_done, should_stop):
        with _FakeEngine.lock:
            _FakeEngine.runs.append((threading.current_thread().name, year))
        results = []
        for t in tasks:
            if should_stop():
                return YearRunResult(year=year, ok=False, candidates=results, stopped=True)
            key = (year, t.planting_date.day)
            if key in _FakeEngine.hang_on:
                _FakeEngine.hang_on.discard(key)
                _FakeEngine.hung_log.append(key)
                _FakeSession.by_pid[self.pid].killed.wait(30)  # ค้างจนถูกปิด
                r = CandidateRunResult(year=year, planting_date=t.planting_date, ok=False,
                                       error_message="window handle is not valid")
                results.append(r)
                on_candidate_done(r)
                continue
            time.sleep(_FakeEngine.task_seconds)
            r = CandidateRunResult(year=year, planting_date=t.planting_date, ok=True)
            results.append(r)
            with _FakeEngine.lock:
                _FakeEngine.dates.append((year, t.planting_date))
            on_candidate_done(r)
        return YearRunResult(year=year, ok=True, candidates=results)


@pytest.fixture
def fake_env(monkeypatch, tmp_path):
    _FakeSession.instances = []
    _FakeSession.fail_names = set()
    _FakeEngine.runs = []
    _FakeEngine.dates = []
    _FakeEngine.task_seconds = 0.05
    _FakeEngine.hang_on = set()
    _FakeEngine.hung_log = []
    _FakeSession.by_pid = {}
    monkeypatch.setattr(desktop_session, "HiddenDesktopSession", _FakeSession)
    monkeypatch.setattr(runner, "CropWatEngine", _FakeEngine)
    monkeypatch.setattr(runner, "_WORKER_STAGGER_SECONDS", 0.02)
    monkeypatch.setattr(runner, "_finish_run", lambda years, settings: None)

    class _Idx:
        def resolve(self, year, month):
            return tmp_path / f"{year}_{month}.PED"

    def _fake_prepare(years, settings):
        total = sum(len(runner.planting_dates_for_year(settings, y)) for y in years)
        run_state.set_candidate_progress(0, total)
        return runner._Batch(
            crop_file=tmp_path / "c.CRO", soil_file=tmp_path / "s.SOI",
            climate_index=_Idx(), rain_index=_Idx(),
            txt_ok=set(), shots_ok=set(), total=total, done=0,
        )

    monkeypatch.setattr(runner, "_prepare_batch", _fake_prepare)
    return tmp_path


def _settings(tmp_path, windows, days=(1, 15)):
    return Settings(
        input_dir=str(tmp_path), output_dir=str(tmp_path / "out"),
        planting_calendar={4: MonthSelection(days=list(days), shot_days=[])},
        hidden_desktop_mode=True, auto_build_outputs=False, parallel_windows=windows,
    )


def _run(years, settings):
    run_state.init_years(years[0], years[-1])
    runner._run_years(years, settings)
    return {y.year: y for y in run_state.snapshot().years}


def _assert_all_dates_once(years, per_year):
    assert len(_FakeEngine.dates) == len(set(_FakeEngine.dates)), "มีวันปลูกถูกทำซ้ำ"
    assert {y for y, _d in _FakeEngine.dates} == set(years)
    assert len(_FakeEngine.dates) == len(years) * per_year, "มีวันปลูกหาย"


def test_all_years_done_once_across_windows(fake_env):
    years = list(range(1990, 2000))
    snap = _run(years, _settings(fake_env, 3))
    assert all(snap[y].status == YearRunStatus.DONE for y in years)
    _assert_all_dates_once(years, 2)
    assert len({t for t, _y in _FakeEngine.runs}) > 1  # แบ่งกันหลายหน้าต่างจริง
    s = run_state.snapshot()
    assert s.candidate_done == s.candidate_total == 20
    assert all(sess.stopped for sess in _FakeSession.instances)
    assert runner._active_hidden_sessions == []
    assert runner._worker_indices == set()
    assert s.windows_target == 0 and s.windows_active == 0  # จบรันแล้วล้างค่า


def test_desktop_names_unique(fake_env):
    _run(list(range(1990, 1996)), _settings(fake_env, 3))
    names = [s.desktop_name for s in _FakeSession.instances]
    assert len(names) == 3 and len(set(names)) == 3
    assert names[0] == desktop_session.desktop_name_for(0)


def test_one_window_fails_others_cover(fake_env):
    _FakeSession.fail_names = {desktop_session.desktop_name_for(1)}
    years = list(range(1990, 1998))
    snap = _run(years, _settings(fake_env, 3))
    assert all(snap[y].status == YearRunStatus.DONE for y in years)
    _assert_all_dates_once(years, 2)


def test_all_windows_fail_marks_years_error(fake_env):
    _FakeSession.fail_names = {desktop_session.desktop_name_for(i) for i in range(4)}
    years = list(range(1990, 1994))
    snap = _run(years, _settings(fake_env, 2))
    assert all(snap[y].status == YearRunStatus.ERROR for y in years)
    assert all("เปิดไม่ขึ้น" in (snap[y].error_message or "") for y in years)
    assert run_state.snapshot().overall_state.value == "idle"
    # ลองเปิดไม่เกินเพดานต่อรอบ ไม่วนเปิด-พังไม่จบ
    assert len(_FakeSession.instances) == runner._MAX_LAUNCH_FAILURES


def test_never_more_windows_than_needed(fake_env):
    _run([1990, 1991], _settings(fake_env, 4))
    assert len(_FakeSession.instances) <= 2


def test_single_window_keeps_year_order(fake_env):
    years = list(range(1990, 1994))
    snap = _run(years, _settings(fake_env, 1))
    assert all(snap[y].status == YearRunStatus.DONE for y in years)
    assert len(_FakeSession.instances) == 1
    assert [y for _t, y in _FakeEngine.runs] == years  # 1 หน้าต่าง = ตามลำดับปีเดิม


def test_auto_mode_uses_recommendation(fake_env, monkeypatch):
    monkeypatch.setattr(system_capacity, "recommend_windows", lambda: 2)
    seen = []
    orig = run_state.set_windows_target

    def _spy(n):
        seen.append(n)
        orig(n)

    monkeypatch.setattr(run_state, "set_windows_target", _spy)
    years = list(range(1990, 1998))
    _run(years, _settings(fake_env, 0))
    assert seen[0] == 2
    assert len({s.desktop_name for s in _FakeSession.instances}) == 2


def _start_async(years, settings):
    run_state.init_years(years[0], years[-1])
    t = threading.Thread(target=runner._run_years, args=(years, settings), daemon=True)
    t.start()
    return t


def _wait_until(cond, timeout=5.0):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if cond():
            return True
        time.sleep(0.01)
    return False


def test_live_decrease_hands_over_remaining_dates(fake_env):
    """ลด 3 → 1 กลางรัน: หน้าต่าง 2, 3 วางงานหลังวันปลูกที่ทำอยู่ แล้วส่งวันที่เหลือของปี
    นั้นกลับเข้าคิว — ทุกวันปลูกต้องทำครบ 1 ครั้งพอดี และสุดท้ายเหลือหน้าต่างเดียว"""
    _FakeEngine.task_seconds = 0.03
    years = list(range(1990, 1996))
    days = (1, 5, 10, 15, 20, 25)
    t = _start_async(years, _settings(fake_env, 3, days=days))
    assert _wait_until(lambda: run_state.snapshot().windows_active == 3)
    runner.set_windows_target(1)
    assert _wait_until(lambda: run_state.snapshot().windows_active == 1)
    t.join(timeout=10)
    assert not t.is_alive()
    snap = {y.year: y for y in run_state.snapshot().years}
    assert all(snap[y].status == YearRunStatus.DONE for y in years)
    _assert_all_dates_once(years, len(days))


def test_live_increase_opens_more_windows(fake_env):
    _FakeEngine.task_seconds = 0.03
    years = list(range(1990, 1998))
    t = _start_async(years, _settings(fake_env, 1, days=(1, 5, 10, 15)))
    assert _wait_until(lambda: run_state.snapshot().windows_active == 1)
    runner.set_windows_target(3)
    assert _wait_until(lambda: run_state.snapshot().windows_active == 3)
    t.join(timeout=10)
    _assert_all_dates_once(years, 4)
    assert len({s.desktop_name for s in _FakeSession.instances}) == 3


def test_set_windows_target_clamped():
    assert runner.set_windows_target(0) == 1
    assert runner.set_windows_target(99) == 4
    run_state.set_windows_target(0)


def test_parallel_windows_setting_clamped():
    assert Settings().parallel_windows == 0  # ค่าเริ่มต้น = อัตโนมัติ
    assert Settings(parallel_windows=-3).parallel_windows == 0
    assert Settings(parallel_windows=99).parallel_windows == 4
    assert Settings(parallel_windows="x").parallel_windows == 0


def test_desktop_name_for():
    base = desktop_session._DESKTOP_NAME
    assert desktop_session.desktop_name_for(0) == base
    assert desktop_session.desktop_name_for(2) == f"{base}-3"


@pytest.mark.parametrize("info,expected", [
    ({"threads": 32, "ram_total_mb": 32768, "ram_free_mb": 16000, "cpu_busy": 0.05}, 3),  # เพดานโหมดอัตโนมัติ
    ({"threads": 16, "ram_total_mb": 16384, "ram_free_mb": 8000, "cpu_busy": 0.1}, 2),
    ({"threads": 8, "ram_total_mb": 8192, "ram_free_mb": 4000, "cpu_busy": 0.1}, 1),
    ({"threads": 4, "ram_total_mb": 4096, "ram_free_mb": 1500, "cpu_busy": 0.2}, 1),
    ({"threads": 32, "ram_total_mb": 32768, "ram_free_mb": 16000, "cpu_busy": 0.9}, 1),  # เครื่องยุ่งมาก
    ({"threads": 32, "ram_total_mb": 8192, "ram_free_mb": 1900, "cpu_busy": 0.0}, 2),  # แรมน้อย
])
def test_capacity_recommendation(info, expected):
    r = system_capacity.recommend(info)
    assert r["recommended"] == expected
    assert 1 <= r["recommended"] <= r["max"]
    assert r["reason"]


def test_watchdog_skips_hung_date_and_finishes(fake_env, monkeypatch):
    """v1.3.1 — CropWat ค้างในวันปลูก 10/04/1991: watchdog ปิด CropWat → วันนั้นถูกข้าม
    (ปีนั้นขึ้นว่ามีปัญหา พร้อมเหตุผล) → วันที่เหลือของปีนั้นทำต่อบน CropWat ตัวใหม่ → งานจบ"""
    monkeypatch.setattr(runner, "_WATCHDOG_SECONDS", 0.4)
    _FakeEngine.task_seconds = 0.03
    _FakeEngine.hang_on = {(1991, 10)}
    years = [1990, 1991, 1992]
    days = (1, 5, 10, 15, 20)
    t = _start_async(years, _settings(fake_env, 2, days=days))
    t.join(timeout=20)
    assert not t.is_alive(), "ยังค้างอยู่ — watchdog ไม่ทำงาน"
    assert _FakeEngine.hung_log == [(1991, 10)]
    snap = {y.year: y for y in run_state.snapshot().years}
    assert snap[1990].status == YearRunStatus.DONE
    assert snap[1992].status == YearRunStatus.DONE
    assert snap[1991].status == YearRunStatus.ERROR
    assert "ไม่ตอบสนอง" in snap[1991].error_message and "10/04" in snap[1991].error_message
    ok_1991 = {d.day for y, d in _FakeEngine.dates if y == 1991}
    assert ok_1991 == {1, 5, 15, 20}  # วันที่เหลือได้ทำต่อครบ ไม่ทำซ้ำ
    assert len(_FakeEngine.dates) == len(set(_FakeEngine.dates))
    assert len(_FakeSession.instances) >= 3  # เปิด CropWat ใหม่แทนตัวที่ค้าง
    s = run_state.snapshot()
    assert s.candidate_done == s.candidate_total


def test_watchdog_quiet_when_progressing(fake_env, monkeypatch):
    """หน้าต่างที่ทำงานปกติ (มีวันปลูกเสร็จเรื่อยๆ) ต้องไม่โดน watchdog ปิด"""
    monkeypatch.setattr(runner, "_WATCHDOG_SECONDS", 0.5)
    _FakeEngine.task_seconds = 0.05
    years = list(range(1990, 1996))
    snap = _run(years, _settings(fake_env, 2, days=(1, 5, 10, 15, 20, 25)))
    assert all(snap[y].status == YearRunStatus.DONE for y in years)
    assert len(_FakeSession.instances) == 2


def test_watchdog_auto_mode_backs_off(fake_env, monkeypatch):
    """โหมดอัตโนมัติ: เจอหน้าต่างค้าง → ลดจำนวนหน้าต่างลง 1 (ตั้งเองไม่ลด)"""
    monkeypatch.setattr(runner, "_WATCHDOG_SECONDS", 0.4)
    monkeypatch.setattr(system_capacity, "recommend_windows", lambda: 3)
    _FakeEngine.task_seconds = 0.03
    _FakeEngine.hang_on = {(1990, 10)}
    seen = []
    orig = run_state.set_windows_target

    def _spy(n):
        seen.append(n)
        orig(n)

    monkeypatch.setattr(run_state, "set_windows_target", _spy)
    _run([1990, 1991, 1992, 1993], _settings(fake_env, 0, days=(1, 5, 10, 15, 20)))
    assert seen[:2] == [3, 2]
