"""
runner.py
=========
เชื่อม automation engine (automation/cropwat_engine.py) + file engine
(file_engine/paths.py) เข้ากับ state.py — รันอยู่ใน background thread แยกจาก
asyncio event loop ของ FastAPI เพราะ pywinauto เป็น blocking call (UI Automation
ไม่เหมาะกับ async โดยตรง)

"""

from __future__ import annotations

import logging
import re
import threading
from pathlib import Path

from automation.cropwat_engine import CropWatEngine, PlantingDateTask, force_close_cropwat
from automation.exceptions import CropWatAutomationError
from config import (
    SPEED_MULTIPLIERS,
    Settings,
    excel_path,
    load_settings,
    planting_dates_for_year,
    screenshot_dir,
    txt_dir,
)
from file_engine import paths as file_paths
from models import OverallRunState, YearRunStatus
from state import run_state

logger = logging.getLogger("runner")

_run_lock = threading.Lock()
_active_thread: threading.Thread | None = None

# ชื่อไฟล์ .txt ที่ export_results สร้าง: "{ปี}_{MMDD}.txt"
_TXT_RE = re.compile(r"^(?P<year>\d{4})_(?P<mmdd>\d{4})\.txt$")
_SHOT_RE = re.compile(r"^(?P<year>\d{4})_(?P<mmdd>\d{4})_schedule\.png$")




def _scan_valid_outputs(settings: Settings) -> tuple[set, set, list[str]]:
    """ตรวจไฟล์ output แบบละเอียด "อ่านเนื้อไฟล์จริง" ไม่ใช่แค่ดูว่ามีไฟล์:

    - .txt ต้อง parse ได้ครบทุกค่า (ใช้ parser ตัวเดียวกับตอนสร้าง Excel เป๊ะ —
      หัวไฟล์, Totals ครบ 6 ค่า, Yield reductions ครบ 4 stage) กันเคสไฟล์เขียน
      ค้างครึ่งทางจากการหยุด/เครื่องดับ/CropWat crash แล้วระบบหลงคิดว่าเสร็จ
    - ภาพต้องมีครบทั้งคู่ (ตาราง + กราฟ) และขนาดไม่เป็นศูนย์

    คืน (txt_ok, shots_ok, invalid_names) — ไฟล์ .txt ที่ไม่สมบูรณ์จะถูกรายงาน
    และ "ไม่นับว่าเสร็จ" → รอบรันถัดไปทำวันปลูกนั้นใหม่ทับให้เอง"""
    from file_engine.txt_parser import TxtParseError, parse_txt

    txt_ok: set[tuple[int, str]] = set()
    invalid: list[str] = []
    d = txt_dir(settings)
    if d.is_dir():
        for p in sorted(d.glob("*.txt")):
            m = _TXT_RE.match(p.name)
            if not m:
                continue
            try:
                parsed = parse_txt(p)
            except (TxtParseError, OSError) as exc:
                logger.warning("ไฟล์ output ไม่สมบูรณ์ (จะรันวันปลูกนี้ใหม่): %s — %s", p.name, exc)
                invalid.append(p.name)
                continue

            # ตรวจว่า "เนื้อไฟล์เป็นของวันปลูกตามชื่อไฟล์จริง" — บทเรียนสำคัญ
            # (ยืนยันจากไฟล์จริงทั้งเครื่องผู้ใช้และเพื่อน): เคยมีบั๊กที่ตั้งวันปลูก
            # ไม่ติดถึง model ของ CropWat ทุกไฟล์เลยเป็นวันปลูกเดียวกันหมดทั้งที่
            # ชื่อไฟล์ต่างกัน — เทียบวันปลูกในหัวไฟล์ (ที่ CropWat พิมพ์เอง) กับ
            # ชื่อไฟล์ ยอมรับทั้งลำดับ วัน/เดือน และ เดือน/วัน (กัน Region ต่างกัน)
            mmdd = m.group("mmdd")
            want_day, want_month = int(mmdd[2:]), int(mmdd[:2])
            pd_match = re.match(r"^\s*(\d{1,2})/(\d{1,2})\s*$", parsed.planting_date)
            if pd_match:
                a, b = int(pd_match.group(1)), int(pd_match.group(2))
                if (a, b) != (want_day, want_month) and (a, b) != (want_month, want_day):
                    logger.warning(
                        "ไฟล์ %s เนื้อในเป็นวันปลูก %s ไม่ตรงกับชื่อไฟล์ (จะรันใหม่)",
                        p.name, parsed.planting_date,
                    )
                    invalid.append(p.name)
                    continue
            txt_ok.add((int(m.group("year")), mmdd))

    shots_ok: set[tuple[int, str]] = set()
    sdir = screenshot_dir(settings)
    if sdir.is_dir():
        for p in sorted(sdir.glob("*_schedule.png")):
            m = _SHOT_RE.match(p.name)
            if not m:
                continue
            graph = p.with_name(p.name.replace("_schedule", "_graph"))
            try:
                if p.stat().st_size > 0 and graph.exists() and graph.stat().st_size > 0:
                    shots_ok.add((int(m.group("year")), m.group("mmdd")))
            except OSError:
                continue
    return txt_ok, shots_ok, invalid


def _candidate_complete(
    year: int, mmdd: str, need_shot: bool, txt_ok: set, shots_ok: set
) -> bool:
    """วันปลูกนับว่า "เสร็จสมบูรณ์" = .txt ผ่านการ parse + (ถ้าเป็นวันที่ต้อง 📷)
    ภาพครบทั้งคู่ด้วย — ขาดอย่างใดอย่างหนึ่ง = ทำใหม่ทั้งวันปลูกนั้น"""
    if (year, mmdd) not in txt_ok:
        return False
    if need_shot and (year, mmdd) not in shots_ok:
        return False
    return True


def _count_done_in_plan(
    settings: Settings, years: list[int], txt_ok: set, shots_ok: set
) -> int:
    """นับว่าในแผน (ปี × ปฏิทินวันปลูก) มีกี่วันปลูกที่เสร็จสมบูรณ์แล้ว — สำหรับตั้ง
    ค่าเริ่มต้นของ progress bar ตอน resume (ไม่ให้เริ่มจาก 0 ทั้งที่ทำไปเยอะแล้ว)"""
    count = 0
    for y in years:
        for d, shot in planting_dates_for_year(settings, y):
            if _candidate_complete(y, f"{d:%m%d}", shot, txt_ok, shots_ok):
                count += 1
    return count


def scan_output_progress(settings: Settings) -> dict:
    """รายงานความคืบหน้าไฟล์ output แบบละเอียด: ต่อปีเสร็จสมบูรณ์กี่/จากกี่วันปลูก
    + รวมทั้งหมด + จำนวนภาพ + จำนวนไฟล์ไม่สมบูรณ์ที่เจอ (จะถูกรันใหม่อัตโนมัติ)"""
    txt_ok, shots_ok, invalid = _scan_valid_outputs(settings)
    years = list(range(settings.default_start_year, settings.default_end_year + 1))

    per_year = []
    total_expected = 0
    total_done = 0
    for y in years:
        planned = planting_dates_for_year(settings, y)
        expected = len(planned)
        done = sum(
            1 for d, shot in planned
            if _candidate_complete(y, f"{d:%m%d}", shot, txt_ok, shots_ok)
        )
        total_expected += expected
        total_done += done
        if expected == 0 and done == 0:
            continue
        status = "done" if done >= expected and expected > 0 else ("partial" if done > 0 else "todo")
        per_year.append({"year": y, "done": done, "expected": expected, "status": status})

    shot_dir = screenshot_dir(settings)
    screenshot_count = len(list(shot_dir.glob("*.png"))) if shot_dir.is_dir() else 0

    return {
        "total_done": total_done,
        "total_expected": total_expected,
        "screenshot_count": screenshot_count,
        "invalid_count": len(invalid),
        "invalid_files": invalid[:20],
        "years": per_year,
    }


def is_run_active() -> bool:
    """v1.0 — เดิมเช็คจาก _active_thread.is_alive() ตรงๆ ซึ่งผิดเวลา thread ค้าง
    (เช่น เจอ dialog ที่ automation จำไม่ได้ ถูกเหวี่ยงออกนอกจอไปแบบมองไม่เห็น
    ทำให้ CropWat ค้างสนิท) — thread ยัง "alive" ตลอดไปทั้งที่ไม่ขยับต่อแล้ว ทำให้
    กดเริ่มรันใหม่ไม่ได้แม้จะปิด-เปิด CropWat ใหม่แล้วก็ตาม ต้องปิดโปรแกรมเราทั้งตัว
    ถึงจะรีเซ็ตได้ (บั๊กที่ผู้ใช้ยืนยันเจอจริง) — เปลี่ยนไปยึด run_state.overall_state
    (ที่ begin_run()/end_run() ควบคุมชัดเจน) เป็นความจริงแทน ตัดขาดจากสุขภาพของ
    thread เดิมไปเลย ให้ force_reset() (ปุ่ม "ปิด CropWat ฉุกเฉิน") คืนสถานะ IDLE ได้
    โดยไม่ต้องรอ thread เก่าตายจริงๆ"""
    return run_state.snapshot().overall_state != OverallRunState.IDLE


def start_run(years: list[int], settings: Settings) -> bool:
    """เริ่ม background thread รันปีที่ระบุ คืน False ถ้ามี run ทำงานอยู่แล้ว"""
    global _active_thread
    with _run_lock:
        if is_run_active():
            return False
        _active_thread = threading.Thread(
            target=_run_years, args=(years, settings), daemon=True
        )
        _active_thread.start()
        return True


def force_reset() -> int:
    """ทางออกฉุกเฉิน (v1.0): บังคับปิด CropWat + คืนสถานะฝั่งเราเป็น IDLE ทันที
    ไม่ต้องรอ thread เก่า (ที่อาจค้างอยู่ในนั้นตลอดไปเพราะรอ dialog ที่กด/ปิดเองไม่
    ได้) จบตามธรรมชาติ — thread เก่าถ้ายังค้างอยู่จริงจะกลายเป็น zombie เฉยๆ (ไม่มี
    ผลอะไรต่อ ไม่แตะ state ที่แชร์กันอันตราย เพราะ engine ของมันคุย CropWat ตัวเก่า
    ที่ตายไปแล้วเท่านั้น) คืนจำนวน process CropWat ที่ปิดสำเร็จ

    v0.7.0: ปิด CropWat บนเดสก์ท็อปซ่อนด้วย (find_windows มองไม่เห็นข้ามเดสก์ท็อป
    ต้องปิดผ่าน process handle ของ session ตรงๆ)"""
    killed = force_close_cropwat()
    session = _active_hidden_session
    if session is not None:
        try:
            session.stop()
            killed += 1
        except Exception:  # noqa: BLE001 -- session ปิดไปเองแล้ว = ปกติ
            pass
    run_state.request_stop()  # เผื่อ thread เก่ายังไม่ตาย ให้รู้ว่าควรหยุดถ้าตื่นมาเช็ค
    run_state.end_run()
    return killed


def start_default_run() -> bool:
    """เริ่มรันช่วงปี default จากค่าที่บันทึกไว้ทั้งหมด — ทางลัดสำหรับ global hotkey
    และปุ่มบน overlay ให้เริ่มรันได้โดยไม่ต้องเปิดหน้าเว็บเลย (ตั้งค่าครั้งเดียวพอ)"""
    if is_run_active():
        return False
    settings = load_settings()
    if not settings.input_dir or not settings.output_dir:
        logger.error("ยังไม่ได้ตั้งค่าโฟลเดอร์ต้นทาง/ผลลัพธ์ — เปิดหน้าตั้งค่า (⚙) ก่อน")
        return False
    run_state.init_years(settings.default_start_year, settings.default_end_year)
    years = list(range(settings.default_start_year, settings.default_end_year + 1))
    return start_run(years, settings)


def _resolve_single_station(root: Path, prefix: str, override: str) -> Path:
    """หาโฟลเดอร์สถานี — ใช้ override ถ้าตั้งไว้ในหน้าตั้งค่า ไม่งั้น auto-pick
    ตัวเดียวที่เจอ (error ชัดเจนถ้าเจอ 0 หรือมากกว่า 1 ตัว)"""
    if override:
        station_dir = Path(override)
        if not station_dir.is_dir():
            raise NotADirectoryError(f"ไม่พบโฟลเดอร์สถานีที่ตั้งไว้: {station_dir}")
        return station_dir
    candidates = file_paths.find_station_folders(root, prefix)
    if not candidates:
        raise FileNotFoundError(f"ไม่พบโฟลเดอร์ขึ้นต้นด้วย '{prefix}' ใน {root}")
    if len(candidates) > 1:
        names = ", ".join(p.name for p in candidates)
        raise ValueError(
            f"เจอโฟลเดอร์ '{prefix}*' มากกว่า 1 ตัว ({names}) กรุณาเลือกในหน้าตั้งค่า"
        )
    return candidates[0]


def _resolve_single_file(root: Path, ext: str, override: str) -> Path:
    if override:
        file_path = Path(override)
        if not file_path.is_file():
            raise FileNotFoundError(f"ไม่พบไฟล์ที่ตั้งไว้: {file_path}")
        return file_path
    candidates = file_paths.find_files_by_extension(root, ext)
    if not candidates:
        raise FileNotFoundError(f"ไม่พบไฟล์ '*{ext}' ใน {root}")
    if len(candidates) > 1:
        names = ", ".join(str(p) for p in candidates)
        raise ValueError(
            f"เจอไฟล์ '*{ext}' มากกว่า 1 ตัว ({names}) กรุณาเลือกในหน้าตั้งค่า"
        )
    return candidates[0]


def check_shift_year_coverage(settings: Settings) -> list[dict]:
    """v0.5.24 — ตามคำขอผู้ใช้: กติกา shift-year (StationIndex.resolve) "ก่อนหน้า
    เท่านั้น" ห้ามเดาใช้ไฟล์เดือนอนาคตแทนเด็ดขาด (แก้ไขจาก v0.5.23 ที่เคยลองทำ
    forward-fallback แล้วผิดหลักการ) — แต่นั่นแปลว่าบางวันปลูกอาจไม่มีไฟล์ climate/
    rain ให้ใช้เลยจริงๆ (เช่น สถานีมีข้อมูลเริ่มทีหลังเดือนที่ทดลองปลูก) ถ้าปล่อย
    ให้รันไปเจอทีละวันปลูกจะเสียเวลาไล่ error ทีละอัน — สแกนตรวจล่วงหน้า "ทุกปี ×
    ทุกเดือนที่จะใช้จริงตาม shift_year_per_candidate" ก่อนเริ่มรัน คืน list ของ
    ปัญหาที่เจอ ให้ frontend แจ้งเตือน + ขอความยินยอมจากผู้ใช้ก่อนเสมอ (ห้ามรันต่อ
    เงียบๆ โดยไม่ถาม — ผู้ใช้กำชับว่าไม่งั้นผลจะเพี้ยนจากที่ต้องการ)"""
    input_root = Path(settings.input_dir)
    try:
        climate_station = _resolve_single_station(input_root, "Clim_", settings.climate_station_dir)
        rain_station = _resolve_single_station(input_root, "Rain_", settings.rain_station_dir)
        climate_index = file_paths.index_climate_station(climate_station)
        rain_index = file_paths.index_rain_station(rain_station)
    except (FileNotFoundError, NotADirectoryError, ValueError) as exc:
        return [{"year": None, "month": None, "climate_error": str(exc), "rain_error": None}]

    years = list(range(settings.default_start_year, settings.default_end_year + 1))
    problems: list[dict] = []
    checked: set[tuple[int, int]] = set()
    for year in years:
        date_flags = planting_dates_for_year(settings, year)
        if not date_flags:
            continue
        # ต้องตรวจ "เดือนที่จะถูกใช้จริง" ให้ตรงกับที่ _run_years_inner จะทำจริง —
        # เปิดสวิตช์ = ทุกเดือนที่มีวันปลูกทดลอง, ปิดสวิตช์ = แค่เดือนแรกสุดของปี
        if settings.shift_year_per_candidate:
            months_to_check = sorted({d.month for d, _ in date_flags})
        else:
            months_to_check = [min(d.month for d, _ in date_flags)]
        for month in months_to_check:
            key = (year, month)
            if key in checked:
                continue
            checked.add(key)
            climate_err: str | None = None
            rain_err: str | None = None
            try:
                climate_index.resolve(year, month)
            except FileNotFoundError as exc:
                climate_err = str(exc)
            try:
                rain_index.resolve(year, month)
            except FileNotFoundError as exc:
                rain_err = str(exc)
            if climate_err or rain_err:
                problems.append({
                    "year": year, "month": month,
                    "climate_error": climate_err, "rain_error": rain_err,
                })
    return problems


def _run_years(years: list[int], settings: Settings) -> None:
    """v0.7.0 — เหลือ 2 โหมดที่พิสูจน์แล้วว่าเสถียรเท่านั้น:
    1. เดสก์ท็อปซ่อน (default): CropWat รันบนเดสก์ท็อปแยกที่มองไม่เห็น ไม่มีการ
       เหวี่ยงหน้าต่าง/shield ใดๆ ทั้งสิ้น (ระบบพวกนั้นคือต้นเหตุ crash — ถอดออก
       หมดแล้ว)
    2. โหมดคลาสสิก (ปิดสวิตช์): เห็น CropWat บนจอ ดึงหน้าต่างขึ้นมาระหว่างรัน
       (background_mode=False ดั้งเดิมที่พิสูจน์มานานว่าเสถียร) — สำหรับเครื่องที่
       โหมดเดสก์ท็อปซ่อนใช้ไม่ได้"""
    run_state.begin_run()

    if settings.hidden_desktop_mode:
        _run_years_hidden_desktop(years, settings)
        return

    # โหมดคลาสสิก: ผู้ใช้เปิด CropWat ค้างไว้เอง หน้าต่างถูกดึงขึ้นมาระหว่างรัน
    speed_multiplier = SPEED_MULTIPLIERS.get(settings.speed_preset, 1.0)
    engine = CropWatEngine(background_mode=False, speed_multiplier=speed_multiplier)
    try:
        engine.connect()
    except CropWatAutomationError as exc:
        logger.error("connect() ล้มเหลว หยุดการรันทั้งหมด: %s", exc)
        for year in years:
            run_state.set_year_status(year, YearRunStatus.ERROR, error_message=str(exc))
        run_state.end_run()
        return
    _run_years_inner(years, settings, engine)


# session เดสก์ท็อปซ่อนที่กำลังรันอยู่ — เก็บไว้ให้ force_reset ปิดได้ (CropWat บน
# เดสก์ท็อปซ่อนมองไม่เห็นจาก find_windows ของ thread บนเดสก์ท็อปปกติ จึงต้องปิด
# ผ่าน process handle ของ session ตรงๆ) และให้ปุ่ม "ดูเดสก์ท็อปซ่อน" เช็คว่ามีอยู่
_active_hidden_session = None


def _run_years_hidden_desktop(years: list[int], settings: Settings) -> None:
    """รันทั้ง batch บนเดสก์ท็อปซ่อน — เปิด CropWat เองในนั้น
    ต้องผูก thread + launch ก่อน connect เสมอ และปิด CropWat + เดสก์ท็อปตอนจบทุกกรณี"""
    global _active_hidden_session
    from desktop_session import DesktopSessionError, HiddenDesktopSession

    session = HiddenDesktopSession(settings.cropwat_exe_path)
    try:
        session.bind_and_launch()
    except DesktopSessionError as exc:
        logger.error("เปิดเดสก์ท็อปซ่อนไม่สำเร็จ หยุดการรันทั้งหมด: %s", exc)
        for year in years:
            run_state.set_year_status(year, YearRunStatus.ERROR, error_message=str(exc))
        run_state.end_run()
        return

    _active_hidden_session = session
    speed_multiplier = SPEED_MULTIPLIERS.get(settings.speed_preset, 1.0)
    # background_mode=True = สั่งเมนูแบบ message ล้วน (เหมาะกับเดสก์ท็อปที่ไม่มี
    # input จริง) — ไม่มี watcher/shield ในระบบอีกต่อไป (ถอดออกทั้งหมดใน v0.7.0)
    engine = CropWatEngine(background_mode=True, speed_multiplier=speed_multiplier)
    # v0.8.0 — pause_check เดิมใช้ park automation ตอนสลับจอไปดูสด (v0.7.x) แต่
    # ฟีเจอร์นั้นถูกถอดออกทั้งหมดแล้ว (ยังชนกับ "Cannot make a visible window
    # modal" ซ้ำๆ ไม่มีทางป้องกันได้ 100%) — ตอนนี้ผูก hook เดียวกัน (เช็คก่อน
    # "ทุกคำสั่งเมนู" ใน _invoke_menu) ให้บริการคำขอ "ถ่ายภาพหน้าจอตอนนี้" แทน
    # (PrintWindow ล้วนๆ ไม่สลับจอเลย ปลอดภัยกว่ามาก)
    import desktop_session

    def _service_peek() -> None:
        desktop_session.service_peek_request(engine, settings.output_dir)

    engine.pause_check = _service_peek
    try:
        try:
            # ล็อกเป้าที่ pid ที่เราเพิ่งเปิดเองตรงๆ (v0.10.1) — กัน
            # ElementAmbiguousError กรณีมี CropWat ค้างจาก session เก่า
            engine.connect(process=session.pid)
        except CropWatAutomationError as exc:
            logger.error("connect() เดสก์ท็อปซ่อนล้มเหลว หยุดการรัน: %s", exc)
            for year in years:
                run_state.set_year_status(year, YearRunStatus.ERROR, error_message=str(exc))
            run_state.end_run()
            return
        # _run_years_inner จัดการสถานะรายปี + end_run() เองครบทุกกรณี
        _run_years_inner(years, settings, engine)
    finally:
        # ปิด CropWat + เดสก์ท็อปซ่อนเสมอ ไม่ว่าจะจบปกติหรือ error
        _active_hidden_session = None
        session.stop()


def _run_years_inner(years: list[int], settings: Settings, engine: CropWatEngine) -> None:
    input_root = Path(settings.input_dir)

    # เตรียมทุกอย่างที่ "คงที่ตลอด batch" ครั้งเดียวก่อนเริ่มวนหลายปี: เช็ค crop/
    # soil เปิดอยู่ + resolve โฟลเดอร์สถานี + index ไฟล์ climate/rain ทั้งสถานี
    # (เดิม index ใหม่ทุกปีในลูป — สแกนโฟลเดอร์ทั้งต้นไม้ recursive ซ้ำ 45 รอบ
    # ทั้งที่ผลเหมือนเดิมทุกรอบ) — ถ้าขั้นนี้พังถือว่าทั้ง batch error หมดเลย
    try:
        crop_file = _resolve_single_file(input_root, ".cro", settings.crop_file)
        soil_file = _resolve_single_file(input_root, ".soi", settings.soil_file)
        engine.ensure_crop_soil_open(crop_file, soil_file)

        climate_station = _resolve_single_station(
            input_root, "Clim_", settings.climate_station_dir
        )
        rain_station = _resolve_single_station(
            input_root, "Rain_", settings.rain_station_dir
        )
        climate_index = file_paths.index_climate_station(climate_station)
        rain_index = file_paths.index_rain_station(rain_station)
    except Exception as exc:  # noqa: BLE001 -- ต้องแจ้งทุกปีว่า error เพราะเหตุนี้
        logger.error("เตรียมไฟล์/สถานีก่อนรันไม่สำเร็จ หยุดการรันทั้งหมด: %s", exc)
        for year in years:
            run_state.set_year_status(
                year, YearRunStatus.ERROR, error_message=f"เตรียมไฟล์/สถานีไม่สำเร็จ: {exc}"
            )
        try:
            import overlay

            overlay.notify("การรันล้มเหลวตั้งแต่เริ่ม", f"เตรียมไฟล์/สถานีไม่สำเร็จ: {str(exc)[:120]}")
        except Exception:  # noqa: BLE001
            pass
        run_state.end_run()
        return

    # resume: ตรวจไฟล์ output "แบบอ่านเนื้อไฟล์จริง" (parse ทุก .txt + เช็คภาพครบคู่)
    # เพื่อ "ทำต่อจากจุดที่ค้าง" — เฉพาะวันปลูกที่เสร็จสมบูรณ์จริงเท่านั้นที่ถูกข้าม
    # ไฟล์ครึ่งๆ กลางๆ จากการหยุด/crash จะถูกจับได้และรันใหม่ทับ
    txt_ok, shots_ok, invalid = _scan_valid_outputs(settings)
    if invalid:
        logger.warning("เจอไฟล์ output ไม่สมบูรณ์ %s ไฟล์ — จะรันวันปลูกเหล่านั้นใหม่", len(invalid))

    # progress ระดับ "วันปลูก" (ละเอียดกว่าระดับปี) ให้ bar เดินสม่ำเสมอไม่ดูค้าง —
    # นับรวมทุกปีก่อนเริ่ม โดยเริ่มนับจากที่ทำเสร็จไปแล้ว (resume) ไม่ใช่ 0
    total_candidates = sum(len(planting_dates_for_year(settings, y)) for y in years)
    completed = {"n": _count_done_in_plan(settings, years, txt_ok, shots_ok)}
    run_state.set_candidate_progress(completed["n"], total_candidates)
    # v0.5.17: บอก state ว่ามีกี่วันปลูก "เสร็จไปแล้วจากรอบก่อนหน้า" (resume) —
    # กัน ETA คำนวณความเร็วจากของเก่าที่ไม่ได้ทำอะไรในรอบนี้เลย (ดู
    # RunState._estimate_eta_seconds)
    run_state.set_baseline_done(completed["n"])

    for year in years:
        if run_state.is_stop_requested():
            run_state.set_year_status(year, YearRunStatus.QUEUED)
            continue

        try:
            date_flags = planting_dates_for_year(settings, year)
            if not date_flags:
                raise ValueError(
                    "ไม่มีวันปลูกที่ทดลองตั้งไว้เลย (ตั้งค่าปฏิทินในหน้า Dashboard ก่อน)"
                )
            # ข้ามเฉพาะวันปลูกที่ "เสร็จสมบูรณ์จริง" (.txt parse ผ่าน + ภาพครบถ้าต้องมี)
            remaining = [
                (d, shot) for d, shot in date_flags
                if not _candidate_complete(year, f"{d:%m%d}", shot, txt_ok, shots_ok)
            ]
            if not remaining:
                # ปีนี้ครบแล้ว — ข้ามทั้งปี ไม่ต้องเปิด CropWat/climate/rain เลย
                run_state.set_year_status(year, YearRunStatus.DONE)
                continue

            tasks = [
                PlantingDateTask(planting_date=d, capture_screenshot=shot)
                for d, shot in remaining
            ]

            # v0.5.21 — ตามคำขอผู้ใช้ (ค่า default ใหม่): resolve ไฟล์ climate/rain
            # ใหม่ตามเดือนจริงของ "แต่ละวันปลูก" (climate_file_for/rain_file_for
            # ถูกเรียกใหม่ใน engine.run_year ก่อนทุกวันปลูก) กติกา shift-year เดิม
            # (เลือกไฟล์เดือน <= เดือนปลูก มากที่สุด ถ้าไม่มีเลยในปีถอยไปปีก่อนหน้า)
            # อยู่ใน StationIndex.resolve() อยู่แล้ว แค่ต้องเรียกให้ตรงเดือนจริง
            # ไม่ใช่ครั้งเดียวจากเดือนแรกสุดของปีแบบเดิม (ปิดสวิตช์นี้เพื่อกลับไป
            # พฤติกรรมเดิม — ทั้งปีใช้ไฟล์เดือนแรกสุดไฟล์เดียว ยืนยันแล้วว่าไฟล์
            # climate/rain เป็นชุดข้อมูลต่อเนื่อง 12 เดือน ครอบคลุมทุกวันปลูกในปี
            # นั้นได้อยู่แล้ว แค่ไม่ตรงเดือนที่สุดเท่านั้น)
            if settings.shift_year_per_candidate:
                climate_file_for = lambda m, _y=year: climate_index.resolve(_y, m)  # noqa: E731
                rain_file_for = lambda m, _y=year: rain_index.resolve(_y, m)  # noqa: E731
            else:
                earliest_month = min(d.month for d, _ in date_flags)
                climate_file_fixed = climate_index.resolve(year, earliest_month)
                rain_file_fixed = rain_index.resolve(year, earliest_month)
                climate_file_for = lambda _m, _f=climate_file_fixed: _f  # noqa: E731
                rain_file_for = lambda _m, _f=rain_file_fixed: _f  # noqa: E731
        except (FileNotFoundError, NotADirectoryError, ValueError) as exc:
            run_state.set_year_status(year, YearRunStatus.ERROR, error_message=str(exc))
            continue

        run_state.set_current_year(year)
        run_state.set_year_status(year, YearRunStatus.RUNNING)

        def _on_candidate_done(_result, _c=completed):
            _c["n"] += 1
            run_state.set_candidate_progress(_c["n"], total_candidates)

        result = engine.run_year(
            year=year,
            tasks=tasks,
            climate_file_for=climate_file_for,
            rain_file_for=rain_file_for,
            export_dir=txt_dir(settings),
            screenshot_dir=screenshot_dir(settings),
            on_candidate_done=_on_candidate_done,
            should_stop=run_state.is_stop_requested,
        )

        if result.stopped:
            # ถูกสั่งหยุดกลางปี — คืนสถานะเป็น "รอคิว" ให้กด "รันปีที่ค้างใหม่"/
            # เริ่มใหม่ได้ (ไฟล์ .txt ของวันปลูกที่ทำเสร็จไปแล้วอยู่ครบ ไม่หายไปไหน
            # รันซ้ำก็แค่เขียนทับไฟล์เดิมด้วยผลเดียวกัน)
            run_state.set_year_status(year, YearRunStatus.QUEUED)
            continue

        if result.ok:
            # เก็บ path โฟลเดอร์ export ไว้ (มีหลายไฟล์ต่อปี — 1 ไฟล์ต่อ 1 วันปลูก
            # ที่ทดลอง — ไม่ใช่ไฟล์เดียวอีกต่อไป)
            exported_files = [str(c.exported_file) for c in result.candidates if c.exported_file]
            run_state.set_year_status(
                year,
                YearRunStatus.DONE,
                exported_file="; ".join(exported_files) if exported_files else None,
            )
        else:
            run_state.set_year_status(
                year, YearRunStatus.ERROR, error_message=result.error_message
            )

    _finish_run(years, settings)
    run_state.end_run()


def _finish_run(years: list[int], settings: Settings) -> None:
    """v0.7.1 (user-journey audit) — ปิดจ็อบให้จบในตัว: เดิมรันเสร็จแบบ "เงียบ
    สนิท" แล้วผู้ใช้ต้องกลับมากดสร้าง Excel + Word เองอีก 2 คลิกทุกครั้ง ทั้งที่
    เป้าหมายจริงของผู้ใช้คือ "ไฟล์ส่งงาน 2 ไฟล์" ไม่ใช่ .txt ดิบ — ถ้ารันจบครบ
    (ไม่ได้กดหยุด) ให้สร้าง Excel+Word ให้เองเลย (ปิดได้ผ่าน auto_build_outputs)
    แล้วยิง Windows notification สรุปผลเสมอ ให้รู้ทันทีแม้กำลังทำงานอื่นอยู่"""
    snap = run_state.snapshot()
    done_years = sum(1 for y in snap.years if y.status == YearRunStatus.DONE)
    error_years = sum(1 for y in snap.years if y.status == YearRunStatus.ERROR)
    stopped = run_state.is_stop_requested()

    built_note = ""
    if not stopped and error_years == 0 and settings.auto_build_outputs and settings.output_dir:
        try:
            build_excel(settings)
            build_word(settings)
            built_note = " — สร้าง Excel + Word ให้แล้ว เปิดโฟลเดอร์ผลลัพธ์ได้เลย"
            logger.info("auto-build: สร้าง Result.xlsx + Screenshots.docx หลังรันเสร็จแล้ว")
        except Exception:  # noqa: BLE001 -- auto-build เป็นของเสริม พังแล้วห้ามล้มการรันที่เสร็จแล้ว
            logger.exception("auto-build ไฟล์ผลลัพธ์ไม่สำเร็จ (กดสร้างเองจากหน้าเว็บได้)")
            built_note = " — สร้างไฟล์อัตโนมัติไม่สำเร็จ กดสร้างเองจากหน้าเว็บได้"

    try:
        import overlay

        if stopped:
            overlay.notify(
                "หยุดการรันแล้ว",
                f"ทำไปแล้ว {snap.candidate_done}/{snap.candidate_total} วันปลูก — "
                "กดเริ่มรันอีกครั้งเพื่อทำต่อจากจุดเดิมได้เลย",
            )
        elif error_years:
            overlay.notify(
                "รันเสร็จ แต่มีปัญหาบางปี",
                f"สำเร็จ {done_years} ปี มีปัญหา {error_years} ปี — เปิดโปรแกรมเพื่อดู"
                " รายละเอียดและกดรันปีที่มีปัญหาซ้ำ",
            )
        else:
            overlay.notify(
                "รันเสร็จครบแล้ว ✓",
                f"{snap.candidate_done}/{snap.candidate_total} วันปลูก สำเร็จทั้ง "
                f"{done_years} ปี{built_note}",
            )
    except Exception:  # noqa: BLE001
        logger.debug("แจ้งเตือนผลรันไม่สำเร็จ", exc_info=True)


def build_excel(settings: Settings) -> int:
    """เฟส 2 (แยกอิสระจากเฟส 1 โดยสิ้นเชิง): อ่านไฟล์ .txt ทั้งหมดที่มีอยู่จริง ณ
    ตอนนี้ → เขียนทับ sheet 'Result' ใหม่ทั้งแผ่น เรียกซ้ำได้ทุกเมื่อ คืนจำนวนปีที่
    เขียนสำเร็จ (import แบบ local กัน circular import ตอน backend startup)"""
    from file_engine.excel_writer import build_result_sheet

    return build_result_sheet(txt_dir(settings), excel_path(settings))


def build_word(settings: Settings) -> int:
    """สร้างไฟล์ Word รวมภาพ screenshot ทั้งหมด (โครงสร้างเหมือนไฟล์ตัวอย่างจริง
    ของผู้ใช้: บรรทัดวันที่ + ภาพตาราง + ภาพกราฟ ต่อ 1 วันปลูก) — เรียกซ้ำได้
    ทุกเมื่อเหมือน build_excel คืนจำนวนวันปลูกที่ใส่ลงเอกสาร"""
    from config import word_path
    from file_engine.word_writer import build_screenshot_doc

    return build_screenshot_doc(screenshot_dir(settings), word_path(settings))
