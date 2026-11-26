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
import queue
import re
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

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
    โดยไม่ต้องรอ thread เก่าตายจริงๆ

    v0.11.1 — บั๊ก "กดเริ่มใหม่ไม่ได้ ขึ้นว่ากำลังรันอยู่ทั้งที่ไม่มีอะไรรัน" ถูก
    แก้ที่ต้นเหตุแล้ว (finally ใน _run_years การันตี end_run ทุกกรณี) จึงยังยึด
    overall_state เป็นความจริงเดียวได้ตามดีไซน์ v1.0 (ตัดขาดจากสุขภาพ thread เพื่อ
    ให้ force_reset คืน IDLE ได้ทันที)"""
    return run_state.snapshot().overall_state != OverallRunState.IDLE


def start_run(years: list[int], settings: Settings) -> bool:
    """เริ่ม background thread รันปีที่ระบุ คืน False ถ้ามี run ทำงานอยู่แล้ว"""
    global _active_thread
    with _run_lock:
        if is_run_active():
            return False
        _snapshot_cropwat_options(settings)
        _active_thread = threading.Thread(
            target=_run_years, args=(years, settings), daemon=True
        )
        _active_thread.start()
        return True


CROPWAT_OPTIONS_SNAPSHOT = "CropWat_Options_used.ini"


def _snapshot_cropwat_options(settings: Settings) -> None:
    """v1.2.0 — เก็บสำเนา Cropwat.ini ที่ใช้รันรอบนี้ไว้ในโฟลเดอร์ผลลัพธ์ (ผลชุดนี้คำนวณด้วย
    Options อะไร — เปิดดูย้อนหลัง/ส่งต่อได้ และ sheet "หมายเหตุ" อ่านจากไฟล์นี้)"""
    import shutil

    import cropwat_options

    try:
        src = cropwat_options.ini_path()
        if settings.output_dir and src.is_file():
            Path(settings.output_dir).mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, Path(settings.output_dir) / CROPWAT_OPTIONS_SNAPSHOT)
    except Exception:  # noqa: BLE001 -- เก็บสำเนาไม่ได้ห้ามขวางการรัน
        logger.warning("เก็บสำเนา Cropwat.ini ไม่สำเร็จ", exc_info=True)


def force_reset() -> int:
    """ทางออกฉุกเฉิน (v1.0): บังคับปิด CropWat + คืนสถานะฝั่งเราเป็น IDLE ทันที
    ไม่ต้องรอ thread เก่า (ที่อาจค้างอยู่ในนั้นตลอดไปเพราะรอ dialog ที่กด/ปิดเองไม่
    ได้) จบตามธรรมชาติ — thread เก่าถ้ายังค้างอยู่จริงจะกลายเป็น zombie เฉยๆ (ไม่มี
    ผลอะไรต่อ ไม่แตะ state ที่แชร์กันอันตราย เพราะ engine ของมันคุย CropWat ตัวเก่า
    ที่ตายไปแล้วเท่านั้น) คืนจำนวน process CropWat ที่ปิดสำเร็จ

    v0.7.0: ปิด CropWat บนเดสก์ท็อปซ่อนด้วย (find_windows มองไม่เห็นข้ามเดสก์ท็อป
    ต้องปิดผ่าน process handle ของ session ตรงๆ)"""
    killed = force_close_cropwat()
    with _sessions_lock:
        sessions = list(_active_hidden_sessions)
    for session in sessions:  # v1.3.0: ปิดทุกหน้าต่างที่รันขนานกันอยู่
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
    # v0.11.1 — บั๊กจากผู้ใช้จริง: หลังงานจบ (โดยเฉพาะจบแบบมีปัญหา) กดเริ่มใหม่ไม่
    # ได้ ขึ้น "ระบบกำลังรันอยู่ ต้องหยุดก่อน" ทั้งที่ไม่มีอะไรรันแล้ว ต้องปิด
    # โปรแกรมทั้งตัวถึงจะหาย — สาเหตุ: ถ้ามี exception หลุดออกมา (เช่น _finish_run
    # ที่ auto-build Excel/Word พังหลังลูปปี แต่ก่อนถึง end_run) หรือ path ไหนที่
    # ลืมเรียก end_run, สถานะ overall_state จะค้างที่ RUNNING ตลอดไป (path
    # เดสก์ท็อปซ่อนยิ่งชัด — finally เดิมปิดแค่ session ไม่เคยแตะ end_run) — ครอบ
    # try/finally รอบนอกสุด "การันตี" คืนสถานะ IDLE ทุกกรณี ไม่ว่าจะจบยังไง
    try:
        if settings.hidden_desktop_mode:
            _run_years_hidden_desktop(years, settings)
            return

        # โหมดคลาสสิก: ผู้ใช้เปิด CropWat ค้างไว้เอง หน้าต่างถูกดึงขึ้นมาระหว่างรัน
        speed_multiplier = SPEED_MULTIPLIERS.get(settings.speed_preset, 1.0)
        engine = CropWatEngine(background_mode=False, speed_multiplier=speed_multiplier)
        engine.error_shots_dir = Path(settings.output_dir) / "_error_dialogs"
        try:
            engine.connect()
        except CropWatAutomationError as exc:
            logger.error("connect() ล้มเหลว หยุดการรันทั้งหมด: %s", exc)
            for year in years:
                run_state.set_year_status(year, YearRunStatus.ERROR, error_message=str(exc))
            return
        _run_years_inner(years, settings, engine)
    except Exception:  # noqa: BLE001 -- กันทุก exception ที่หลุดออกมาโดยไม่คาดคิด
        logger.exception("การรันล้มเหลวโดยไม่คาดคิด — คืนสถานะให้พร้อมเริ่มใหม่")
        try:
            import overlay

            overlay.notify(
                "การรันหยุดผิดปกติ",
                "เกิดข้อผิดพลาดที่ไม่คาดคิด — ระบบคืนสถานะให้พร้อมเริ่มใหม่แล้ว",
            )
        except Exception:  # noqa: BLE001
            pass
    finally:
        # ถ้ายังไม่ IDLE (มี exception หลุด หรือ path ไหนลืมปิด) ปิดให้เสมอ
        if run_state.snapshot().overall_state != OverallRunState.IDLE:
            run_state.end_run()


# session เดสก์ท็อปซ่อนที่กำลังรันอยู่ — เก็บไว้ให้ force_reset ปิดได้ (CropWat บน
# เดสก์ท็อปซ่อนมองไม่เห็นจาก find_windows ของ thread บนเดสก์ท็อปปกติ จึงต้องปิด
# ผ่าน process handle ของ session ตรงๆ) — v1.3.0: มีได้หลายตัวพร้อมกัน (1 ตัวต่อ 1
# หน้าต่าง CropWat ที่รันขนานกัน)
_active_hidden_sessions: list = []
_sessions_lock = threading.Lock()

# v1.3.0 — เว้นระยะเปิด CropWat แต่ละหน้าต่าง (วินาที) ไม่ให้เปิดพร้อมกันทีเดียว (ตัวคุม
# ใน _run_years_hidden_desktop เปิดทีละบาน):
# ตอนเปิด CropWat อ่าน Cropwat.ini + โหลดหน้าต่างหนักสุด เปิดชนกันหลายตัวเครื่องช้าๆ
# อาจหน่วงจน Welcome dialog/หน้าต่างหลักโผล่ช้าเกิน timeout
_WORKER_STAGGER_SECONDS = 2.5


@dataclass
class _Batch:
    """ของที่ "คงที่ตลอด batch" — เตรียมครั้งเดียวก่อนเริ่ม แชร์ให้ทุกหน้าต่างใช้ร่วม"""

    crop_file: Path
    soil_file: Path
    climate_index: object
    rain_index: object
    txt_ok: set
    shots_ok: set
    total: int
    done: int
    lock: threading.Lock = field(default_factory=threading.Lock)
    # v1.3.1 — วันปลูกที่ "ข้ามไป" เพราะหน้าต่างค้างจน watchdog ต้องปิด CropWat (ปี → ข้อความ)
    # ส่วนที่เหลือของปีนั้นย้ายไปทำต่ออีกหน้าต่าง — ตอนปีนั้นจบต้องยังขึ้นว่ามีปัญหา
    year_failures: dict = field(default_factory=dict)

    def candidate_done(self) -> None:
        # หลายหน้าต่างเรียกพร้อมกันได้ — นับ + ส่ง progress ในล็อกเดียวกัน กันเลขถอยหลัง
        # (min กับ total: วันปลูกที่ล้มเหลวแล้วถูกย้ายไปทำต่ออีกหน้าต่างตอนลดจำนวนหน้าต่าง
        # จะถูกนับซ้ำ — ไม่ให้ตัวเลขเกิน 100%)
        with self.lock:
            self.done = min(self.done + 1, self.total)
            run_state.set_candidate_progress(self.done, self.total)


@dataclass
class _YearPlan:
    year: int
    tasks: list
    climate_file_for: object
    rain_file_for: object


def _fail_all_years(years: list[int], message: str) -> None:
    for year in years:
        run_state.set_year_status(year, YearRunStatus.ERROR, error_message=message)


def _notify_safe(title: str, message: str) -> None:
    try:
        import overlay

        overlay.notify(title, message)
    except Exception:  # noqa: BLE001
        pass


def _prepare_batch(years: list[int], settings: Settings) -> Optional[_Batch]:
    """เตรียมทุกอย่างที่ "คงที่ตลอด batch" ครั้งเดียวก่อนเริ่มวนหลายปี: resolve ไฟล์
    crop/soil + โฟลเดอร์สถานี + index ไฟล์ climate/rain ทั้งสถานี (เดิม index ใหม่ทุก
    ปีในลูป — สแกนโฟลเดอร์ทั้งต้นไม้ recursive ซ้ำ 45 รอบ ทั้งที่ผลเหมือนเดิมทุกรอบ)
    + ตรวจไฟล์ผลลัพธ์เดิม (resume) — ถ้าขั้นนี้พังถือว่าทั้ง batch error หมด คืน None"""
    input_root = Path(settings.input_dir)
    try:
        crop_file = _resolve_single_file(input_root, ".cro", settings.crop_file)
        soil_file = _resolve_single_file(input_root, ".soi", settings.soil_file)
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
        _fail_all_years(years, f"เตรียมไฟล์/สถานีไม่สำเร็จ: {exc}")
        _notify_safe("การรันล้มเหลวตั้งแต่เริ่ม", f"เตรียมไฟล์/สถานีไม่สำเร็จ: {str(exc)[:120]}")
        return None

    # resume: ตรวจไฟล์ output "แบบอ่านเนื้อไฟล์จริง" (parse ทุก .txt + เช็คภาพครบคู่)
    # เพื่อ "ทำต่อจากจุดที่ค้าง" — เฉพาะวันปลูกที่เสร็จสมบูรณ์จริงเท่านั้นที่ถูกข้าม
    # ไฟล์ครึ่งๆ กลางๆ จากการหยุด/crash จะถูกจับได้และรันใหม่ทับ
    txt_ok, shots_ok, invalid = _scan_valid_outputs(settings)
    if invalid:
        logger.warning("เจอไฟล์ output ไม่สมบูรณ์ %s ไฟล์ — จะรันวันปลูกเหล่านั้นใหม่", len(invalid))

    # progress ระดับ "วันปลูก" (ละเอียดกว่าระดับปี) ให้ bar เดินสม่ำเสมอไม่ดูค้าง —
    # นับรวมทุกปีก่อนเริ่ม โดยเริ่มนับจากที่ทำเสร็จไปแล้ว (resume) ไม่ใช่ 0
    total = sum(len(planting_dates_for_year(settings, y)) for y in years)
    done = _count_done_in_plan(settings, years, txt_ok, shots_ok)
    run_state.set_candidate_progress(done, total)
    # v0.5.17: บอก state ว่ามีกี่วันปลูก "เสร็จไปแล้วจากรอบก่อนหน้า" (resume) —
    # กัน ETA คำนวณความเร็วจากของเก่าที่ไม่ได้ทำอะไรในรอบนี้เลย (ดู
    # RunState._estimate_eta_seconds)
    run_state.set_baseline_done(done)
    return _Batch(
        crop_file=crop_file, soil_file=soil_file,
        climate_index=climate_index, rain_index=rain_index,
        txt_ok=txt_ok, shots_ok=shots_ok, total=total, done=done,
    )


def _plan_year(year: int, settings: Settings, batch: _Batch) -> Optional[_YearPlan]:
    """สร้างรายการวันปลูกที่ยังต้องรันของปีนี้ — คืน None ถ้าไม่ต้องรัน (ครบแล้ว =
    ตั้งสถานะ DONE, ไฟล์/ปฏิทินมีปัญหา = ตั้งสถานะ ERROR ให้เรียบร้อยแล้ว)"""
    try:
        date_flags = planting_dates_for_year(settings, year)
        if not date_flags:
            raise ValueError(
                "ไม่มีวันปลูกที่ทดลองตั้งไว้เลย (ตั้งค่าปฏิทินในหน้า Dashboard ก่อน)"
            )
        # ข้ามเฉพาะวันปลูกที่ "เสร็จสมบูรณ์จริง" (.txt parse ผ่าน + ภาพครบถ้าต้องมี)
        remaining = [
            (d, shot) for d, shot in date_flags
            if not _candidate_complete(year, f"{d:%m%d}", shot, batch.txt_ok, batch.shots_ok)
        ]
        if not remaining:
            # ปีนี้ครบแล้ว — ข้ามทั้งปี ไม่ต้องเปิด CropWat/climate/rain เลย
            run_state.set_year_status(year, YearRunStatus.DONE)
            return None

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
        climate_index = batch.climate_index
        rain_index = batch.rain_index
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
        return None
    return _YearPlan(year, tasks, climate_file_for, rain_file_for)


def _execute_year(
    plan: _YearPlan,
    settings: Settings,
    engine: CropWatEngine,
    batch: _Batch,
    should_stop=None,
    on_progress=None,
):
    """รันวันปลูกที่เหลือของ 1 ปีบน engine (หน้าต่าง CropWat) ที่ให้มา + ตั้งสถานะปี
    คืน YearRunResult — should_stop (ถ้าไม่ระบุ = คำสั่งหยุดของผู้ใช้) ถูกเช็คก่อนทุก
    วันปลูก ใช้ให้หน้าต่างที่ถูก "ลดจำนวน" ระหว่างรันวางงานกลางปีได้ (ดู _hidden_worker)
    on_progress ถูกเรียกทุกครั้งที่จบ 1 วันปลูก (heartbeat ของ watchdog)"""
    year = plan.year
    run_state.set_current_year(year)
    run_state.set_year_status(year, YearRunStatus.RUNNING)

    def _done(_result) -> None:
        batch.candidate_done()
        if on_progress is not None:
            on_progress()

    result = engine.run_year(
        year=year,
        tasks=plan.tasks,
        climate_file_for=plan.climate_file_for,
        rain_file_for=plan.rain_file_for,
        export_dir=txt_dir(settings),
        screenshot_dir=screenshot_dir(settings),
        on_candidate_done=_done,
        should_stop=should_stop or run_state.is_stop_requested,
    )

    if result.stopped:
        # ถูกสั่งหยุดกลางปี — คืนสถานะเป็น "รอคิว" ให้กด "รันปีที่ค้างใหม่"/
        # เริ่มใหม่ได้ (ไฟล์ .txt ของวันปลูกที่ทำเสร็จไปแล้วอยู่ครบ ไม่หายไปไหน
        # รันซ้ำก็แค่เขียนทับไฟล์เดิมด้วยผลเดียวกัน)
        run_state.set_year_status(year, YearRunStatus.QUEUED)
        return result

    skipped = batch.year_failures.get(year) or []
    if result.ok and not skipped:
        # เก็บ path โฟลเดอร์ export ไว้ (มีหลายไฟล์ต่อปี — 1 ไฟล์ต่อ 1 วันปลูก
        # ที่ทดลอง — ไม่ใช่ไฟล์เดียวอีกต่อไป)
        exported_files = [str(c.exported_file) for c in result.candidates if c.exported_file]
        run_state.set_year_status(
            year,
            YearRunStatus.DONE,
            exported_file="; ".join(exported_files) if exported_files else None,
        )
    else:
        msgs = ([result.error_message] if result.error_message else []) + list(skipped)
        run_state.set_year_status(
            year, YearRunStatus.ERROR, error_message=" · ".join(msgs) or "มีวันปลูกที่ทำไม่สำเร็จ"
        )
    return result


# v1.3.0 — หน้าต่างที่ทำงานอยู่ (เลขลำดับ worker → ใช้ตั้งชื่อเดสก์ท็อปซ่อน) แชร์ระหว่าง
# ตัวคุม (_run_years_hidden_desktop) กับ worker ทุกตัว
_worker_indices: set[int] = set()
_workers_lock = threading.Lock()
# เปิดหน้าต่างไม่สำเร็จเกินกี่ครั้งต่อการรัน 1 รอบ แล้วเลิกพยายามเปิดเพิ่ม (กันวนเปิด-พังไม่จบ)
_MAX_LAUNCH_FAILURES = 3

# v1.3.1 — watchdog กันค้าง (ผู้ใช้: "เหมือนมันไม่ยอมข้ามก็เลยค้าง"): ทุกหน้าต่างอัปเดต
# heartbeat ทุกครั้งที่จบ 1 วันปลูก/เริ่มปีใหม่ ถ้าเงียบเกิน _WATCHDOG_SECONDS (ปกติ
# วันปลูกละ ~10-30 วิ รวมลองใหม่ทุกขั้นแล้วไม่เกิน ~2 นาที) ตัวคุมจะปิด CropWat ของหน้าต่าง
# นั้นทิ้ง → คำสั่งที่ค้างอยู่ error หลุดออกมา → วันปลูกที่ค้างถูกบันทึกว่ามีปัญหา (ข้ามไป
# เหมือนเวอร์ชันก่อนๆ) → วันที่เหลือของปีนั้นกลับเข้าคิว → ตัวคุมเปิด CropWat ใหม่มาทำต่อ
_WATCHDOG_SECONDS = 300.0
_worker_heartbeat: dict[int, float] = {}
_worker_sessions: dict[int, object] = {}
_worker_hung: set[int] = set()


def _beat(index: int) -> None:
    with _workers_lock:
        _worker_heartbeat[index] = time.monotonic()


def _is_hung(index: int) -> bool:
    with _workers_lock:
        return index in _worker_hung


def _watchdog_check(active: set[int], auto: bool = False) -> None:
    """เรียกจากลูปตัวคุม — ปิด CropWat ของหน้าต่างที่เงียบนานเกินกำหนด
    auto=True (โหมดอัตโนมัติ): หน้าต่างค้าง = สัญญาณว่าเครื่องรับไม่ไหว → ลดจำนวนหน้าต่าง
    ที่ตั้งไว้ลง 1 ทันที (ไม่ต่ำกว่า 1) แทนที่จะเปิดใหม่เท่าเดิมแล้วค้างซ้ำ"""
    now = time.monotonic()
    with _workers_lock:
        hung = [
            i for i in active
            if i not in _worker_hung and now - _worker_heartbeat.get(i, now) > _WATCHDOG_SECONDS
        ]
        for i in hung:
            _worker_hung.add(i)
        targets = [(i, _worker_sessions.get(i)) for i in hung]
    for i, session in targets:
        logger.warning(
            "หน้าต่างที่ %s ไม่มีความคืบหน้านานเกิน %.0f วินาที — ปิด CropWat ที่ค้างแล้วเปิดใหม่",
            i + 1, _WATCHDOG_SECONDS,
        )
        if session is not None:
            session.kill_process()
    if targets and auto:
        cur = run_state.windows_target()
        if cur > 1:
            run_state.set_windows_target(cur - 1)
            logger.warning("โหมดอัตโนมัติ: ลดจำนวนหน้าต่าง CropWat จาก %s เหลือ %s เพราะมีหน้าต่างค้าง", cur, cur - 1)


def _worker_should_retire(index: int) -> bool:
    """หน้าต่างนี้ควรปิดตัวเองไหม — เมื่อผู้ใช้ลดจำนวนหน้าต่างระหว่างรัน หน้าต่างที่ลำดับ
    สูงสุดจะถูกปิดก่อน (เทียบลำดับในกลุ่มที่ทำงานอยู่ ไม่ใช่เลข index ตรงๆ เผื่อมีช่องว่าง
    จากหน้าต่างที่เปิดไม่ขึ้น)"""
    target = run_state.windows_target()
    with _workers_lock:
        active = sorted(_worker_indices)
    if index not in active:
        return False
    return active.index(index) >= max(1, target)


def _run_years_hidden_desktop(years: list[int], settings: Settings) -> None:
    """รันทั้ง batch บนเดสก์ท็อปซ่อน — โปรแกรมเปิด CropWat เองในนั้น

    v1.3.0 — รันหลายหน้าต่างพร้อมกันได้: เตรียมแผนทุกปีครั้งเดียว ใส่คิว แล้ว "ตัวคุม"
    (ลูปในฟังก์ชันนี้) คอยเปิด worker thread ละ 1 หน้าต่าง CropWat (เดสก์ท็อปซ่อนแยกของ
    ใครของมัน) ให้ครบตามจำนวนที่ตั้งไว้ — จำนวนนั้นเริ่มจาก settings.parallel_windows
    (0 = อัตโนมัติตามกำลังเครื่อง ดู system_capacity) และผู้ใช้ปรับเพิ่ม/ลดได้ระหว่างรัน
    (run_state.set_windows_target ผ่าน POST /api/run/windows):
    - เพิ่ม: ตัวคุมเปิดหน้าต่างใหม่ทีละบาน เว้นระยะ _WORKER_STAGGER_SECONDS
    - ลด: หน้าต่างลำดับท้ายวางงานหลังวันปลูกที่ทำอยู่เสร็จ (ไม่ต้องรอจบปี) คืนวันปลูกที่
      เหลือของปีนั้นเข้าคิวให้หน้าต่างอื่นทำต่อ แล้วปิด CropWat ของตัวเอง
    1 ปีทำบนหน้าต่างเดียวตามลำดับเดิมเสมอ (ยกเว้นถูกลดจำนวนกลางปี ส่วนที่เหลือย้ายไปทำต่อ
    อีกหน้าต่างทีละวันปลูกเหมือนเดิม) ผลจึงเหมือนรันทีละหน้าต่าง — thread ใหม่ทุกหน้าต่าง
    ตรงกติกา SetThreadDesktop ที่ต้องเรียกบน thread ที่ยังไม่มีหน้าต่าง"""
    batch = _prepare_batch(years, settings)
    if batch is None:
        run_state.end_run()
        return

    plans: list[_YearPlan] = []
    for year in years:
        if run_state.is_stop_requested():
            run_state.set_year_status(year, YearRunStatus.QUEUED)
            continue
        plan = _plan_year(year, settings, batch)
        if plan is not None:
            plans.append(plan)

    if plans:
        work: queue.Queue = queue.Queue()
        for plan in plans:
            work.put(plan)
        if settings.parallel_windows > 0:
            target = settings.parallel_windows
        else:
            import system_capacity

            target = system_capacity.recommend_windows()
        run_state.set_windows_target(target)
        failures: list[str] = []
        logger.info("เริ่มรัน %s ปี ด้วย CropWat %s หน้าต่างพร้อมกัน", len(plans), target)

        threads: list[threading.Thread] = []
        last_spawn = -1e9
        with _workers_lock:
            _worker_heartbeat.clear()
            _worker_sessions.clear()
            _worker_hung.clear()
        while True:
            with _workers_lock:
                active = set(_worker_indices)
            run_state.set_windows_active(len(active))
            _watchdog_check(active, auto=settings.parallel_windows == 0)
            stop = run_state.is_stop_requested()
            if not active and (work.empty() or stop or len(failures) >= _MAX_LAUNCH_FAILURES):
                break
            want = max(1, run_state.windows_target())
            need_more = (
                not stop
                and not work.empty()
                and len(active) < want
                and len(failures) < _MAX_LAUNCH_FAILURES
            )
            wait_left = _WORKER_STAGGER_SECONDS - (time.monotonic() - last_spawn)
            if need_more and wait_left <= 0:
                index = next(i for i in range(len(active) + 1) if i not in active)
                with _workers_lock:
                    _worker_indices.add(index)
                    _worker_heartbeat[index] = time.monotonic()
                t = threading.Thread(
                    target=_hidden_worker,
                    args=(index, settings, batch, work, failures),
                    name=f"cropwat-window-{index + 1}",
                    daemon=True,
                )
                t.start()
                threads.append(t)
                last_spawn = time.monotonic()
                continue
            time.sleep(min(0.2, wait_left) if need_more else 0.2)
        for t in threads:
            t.join()
        run_state.set_windows_active(0)

        # ปีที่ยังค้างในคิว = ไม่มีหน้าต่างไหนได้ทำ (สั่งหยุด หรือเปิด CropWat ไม่ขึ้นเลย)
        reason = failures[0] if failures else "ไม่มีหน้าต่าง CropWat ที่พร้อมทำงาน"
        while True:
            try:
                plan = work.get_nowait()
            except queue.Empty:
                break
            if run_state.is_stop_requested():
                run_state.set_year_status(plan.year, YearRunStatus.QUEUED)
            else:
                run_state.set_year_status(plan.year, YearRunStatus.ERROR, error_message=reason)

    _finish_run(years, settings)
    run_state.end_run()


def _hidden_worker(
    index: int,
    settings: Settings,
    batch: _Batch,
    work: queue.Queue,
    failures: list[str],
) -> None:
    """1 หน้าต่าง CropWat: เปิดเดสก์ท็อปซ่อนของตัวเอง → เปิด CropWat → เปิด crop/soil
    → หยิบปีจากคิวมาทำจนหมด (หรือจนถูกลดจำนวนหน้าต่าง) → ปิด CropWat + เดสก์ท็อปเสมอ
    ไม่ว่าจบแบบไหน — ตัวคุมใส่ index ไว้ใน _worker_indices ก่อนเริ่ม thread แล้ว"""
    import desktop_session
    from desktop_session import DesktopSessionError, HiddenDesktopSession

    def _fail(message: str) -> None:
        failures.append(f"หน้าต่าง CropWat ที่ {index + 1}: {message}")

    def _retire_or_stop() -> bool:
        return run_state.is_stop_requested() or _worker_should_retire(index) or _is_hung(index)

    session = None
    try:
        if _retire_or_stop() or work.empty():
            return
        session = HiddenDesktopSession(
            settings.cropwat_exe_path, desktop_name=desktop_session.desktop_name_for(index)
        )
        try:
            session.bind_and_launch()
        except DesktopSessionError as exc:
            logger.error("หน้าต่างที่ %s: เปิดเดสก์ท็อปซ่อนไม่สำเร็จ: %s", index + 1, exc)
            _fail(str(exc))
            return
        with _sessions_lock:
            _active_hidden_sessions.append(session)
        with _workers_lock:
            _worker_sessions[index] = session
        _beat(index)

        speed_multiplier = SPEED_MULTIPLIERS.get(settings.speed_preset, 1.0)
        # background_mode=True = สั่งเมนูแบบ message ล้วน (เหมาะกับเดสก์ท็อปที่ไม่มี
        # input จริง) — ไม่มี watcher/shield ในระบบอีกต่อไป (ถอดออกทั้งหมดใน v0.7.0)
        engine = CropWatEngine(background_mode=True, speed_multiplier=speed_multiplier)
        # pause_check ถูกเรียกก่อน "ทุกคำสั่งเมนู" ใน _invoke_menu — ใช้บริการคำขอ
        # "ถ่ายภาพหน้าจอตอนนี้" (PrintWindow ล้วนๆ ไม่สลับจอเลย) หลายหน้าต่างก็มีแค่
        # ตัวแรกที่เห็นคำขอเป็นคนถ่าย (ดู desktop_session.service_peek_request)
        engine.pause_check = lambda: desktop_session.service_peek_request(
            engine, settings.output_dir
        )
        engine.error_shots_dir = Path(settings.output_dir) / "_error_dialogs"
        try:
            # ล็อกเป้าที่ pid ที่เราเพิ่งเปิดเองตรงๆ (v0.10.1) — กัน
            # ElementAmbiguousError กรณีมี CropWat ค้างจาก session เก่า
            engine.connect(process=session.pid)
        except CropWatAutomationError as exc:
            logger.error("หน้าต่างที่ %s: connect() ล้มเหลว: %s", index + 1, exc)
            _fail(str(exc))
            return
        try:
            engine.ensure_crop_soil_open(batch.crop_file, batch.soil_file)
        except Exception as exc:  # noqa: BLE001
            logger.error("หน้าต่างที่ %s: เปิดไฟล์ crop/soil ไม่สำเร็จ: %s", index + 1, exc)
            _fail(f"เตรียมไฟล์/สถานีไม่สำเร็จ: {exc}")
            return

        while not _retire_or_stop():
            try:
                plan = work.get_nowait()
            except queue.Empty:
                break
            _beat(index)
            try:
                result = _execute_year(
                    plan, settings, engine, batch,
                    should_stop=_retire_or_stop, on_progress=lambda: _beat(index),
                )
            except Exception as exc:  # noqa: BLE001 -- ปีเดียวพังห้ามล้มทั้งหน้าต่าง
                logger.exception("หน้าต่างที่ %s: ปี %s ล้มเหลวโดยไม่คาดคิด", index + 1, plan.year)
                run_state.set_year_status(plan.year, YearRunStatus.ERROR, error_message=str(exc))
                continue
            if _is_hung(index):
                # watchdog ปิด CropWat ของหน้าต่างนี้ไปแล้ว: วันปลูกที่ลงมือไปแล้วแต่ไม่สำเร็จ
                # (ตัวที่ค้าง) = ข้ามไปพร้อมบันทึกว่ามีปัญหา ส่วนที่ยังไม่ได้เริ่มกลับเข้าคิว
                # ให้ CropWat ตัวใหม่ทำต่อ แล้วหน้าต่างนี้จบตัวเอง (ตัวคุมเปิดตัวใหม่แทน)
                attempted = {c.planting_date for c in result.candidates}
                note = (
                    f"CropWat ไม่ตอบสนองนานเกิน {_WATCHDOG_SECONDS / 60:.0f} นาที — "
                    "ระบบปิดแล้วเปิดใหม่ให้อัตโนมัติ ข้ามวันนี้ไปก่อน (กด \"ประมวลผลซ้ำปีที่มีปัญหา\" ภายหลังได้)"
                )
                for c in result.candidates:
                    if not c.ok:
                        batch.year_failures.setdefault(plan.year, []).append(
                            f"{c.planting_date:%d/%m}: {note}"
                        )
                rest = [t for t in plan.tasks if t.planting_date not in attempted]
                if rest and not run_state.is_stop_requested():
                    work.put(_YearPlan(plan.year, rest, plan.climate_file_for, plan.rain_file_for))
                    run_state.set_year_status(plan.year, YearRunStatus.QUEUED)
                elif result.stopped:
                    skipped = batch.year_failures.get(plan.year)
                    if skipped:
                        run_state.set_year_status(
                            plan.year, YearRunStatus.ERROR, error_message=" · ".join(skipped)
                        )
                break
            if result.stopped and not run_state.is_stop_requested():
                # ถูกลดจำนวนหน้าต่างกลางปี — คืนวันปลูกที่ยังไม่สำเร็จเข้าคิวให้หน้าต่างอื่น
                done_ok = {c.planting_date for c in result.candidates if c.ok}
                rest = [t for t in plan.tasks if t.planting_date not in done_ok]
                if rest:
                    work.put(_YearPlan(plan.year, rest, plan.climate_file_for, plan.rain_file_for))
                    logger.info(
                        "หน้าต่างที่ %s: ลดจำนวนหน้าต่าง — ส่งปี %s ที่เหลือ %s วันปลูกให้หน้าต่างอื่น",
                        index + 1, plan.year, len(rest),
                    )
                else:
                    run_state.set_year_status(plan.year, YearRunStatus.DONE)
        if _worker_should_retire(index) and not run_state.is_stop_requested():
            logger.info("หน้าต่างที่ %s: ปิดตามจำนวนหน้าต่างที่ลดลง", index + 1)
    except Exception as exc:  # noqa: BLE001 -- กันทุกอย่าง ไม่ให้ worker ตายเงียบ
        logger.exception("หน้าต่างที่ %s ล้มเหลวโดยไม่คาดคิด", index + 1)
        _fail(str(exc))
    finally:
        if session is not None:
            with _sessions_lock:
                if session in _active_hidden_sessions:
                    _active_hidden_sessions.remove(session)
            session.stop()
        with _workers_lock:
            _worker_indices.discard(index)
            _worker_sessions.pop(index, None)
            _worker_heartbeat.pop(index, None)
            _worker_hung.discard(index)


def set_windows_target(n: int) -> int:
    """v1.3.0 — ผู้ใช้ปรับจำนวนหน้าต่างระหว่างรัน (POST /api/run/windows) คืนค่าที่ใช้จริง"""
    from config import MAX_PARALLEL_WINDOWS

    n = max(1, min(MAX_PARALLEL_WINDOWS, int(n)))
    run_state.set_windows_target(n)
    logger.info("ผู้ใช้ปรับจำนวนหน้าต่าง CropWat เป็น %s", n)
    return n


def _run_years_inner(years: list[int], settings: Settings, engine: CropWatEngine) -> None:
    """โหมดคลาสสิก (หน้าต่าง CropWat เดียวที่ผู้ใช้เปิดเองบนจอ) — ทำทีละปีตามลำดับ"""
    batch = _prepare_batch(years, settings)
    if batch is None:
        run_state.end_run()
        return
    try:
        engine.ensure_crop_soil_open(batch.crop_file, batch.soil_file)
    except Exception as exc:  # noqa: BLE001 -- ต้องแจ้งทุกปีว่า error เพราะเหตุนี้
        logger.error("เปิดไฟล์ crop/soil ไม่สำเร็จ หยุดการรันทั้งหมด: %s", exc)
        _fail_all_years(years, f"เตรียมไฟล์/สถานีไม่สำเร็จ: {exc}")
        _notify_safe("การรันล้มเหลวตั้งแต่เริ่ม", f"เตรียมไฟล์/สถานีไม่สำเร็จ: {str(exc)[:120]}")
        run_state.end_run()
        return

    for year in years:
        if run_state.is_stop_requested():
            run_state.set_year_status(year, YearRunStatus.QUEUED)
            continue
        plan = _plan_year(year, settings, batch)
        if plan is None:
            continue
        _execute_year(plan, settings, engine, batch)

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

    # v0.12.0 — สร้างไฟล์ให้แม้บางปีมีปัญหา (เดิมสร้างเฉพาะตอนผ่านครบทุกปี): ข้อมูล
    # จริงของผู้ใช้มีปีที่ climate ขาดเป็นปกติ (spec: "ขาดข้อมูล Sun เดือน 4") ถ้า
    # รอให้ไม่มี error เลย Excel จะไม่ถูกสร้างอัตโนมัติแทบทุกครั้ง — ปีที่มีปัญหา
    # ถูกบันทึกไว้ใน sheet "หมายเหตุ" แทน
    built_note = ""
    built = False
    if not stopped and settings.auto_build_outputs and settings.output_dir:
        try:
            build_excel(settings)
            built = True
            try:
                build_word(settings)
                built_note = " — สร้าง Excel + Word ให้แล้ว เปิดโฟลเดอร์ผลลัพธ์ได้เลย"
            except FileNotFoundError:
                # ไม่ได้เลือกวันถ่ายภาพไว้เลย = ไม่มีภาพให้ทำ Word — ไม่ใช่ความล้มเหลว
                built_note = " — สร้าง Excel ให้แล้ว เปิดโฟลเดอร์ผลลัพธ์ได้เลย"
            logger.info("auto-build: สร้างไฟล์ผลลัพธ์หลังรันเสร็จแล้ว")
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
                f"สำเร็จ {done_years} ปี มีปัญหา {error_years} ปี"
                + (" — สร้าง Excel แล้ว (ปีที่มีปัญหาอยู่ใน sheet หมายเหตุ)" if built else
                   " — เปิดโปรแกรมเพื่อดูรายละเอียดและกดรันปีที่มีปัญหาซ้ำ"),
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

    return build_result_sheet(txt_dir(settings), excel_path(settings), _collect_run_notes(settings))


# ชื่อภาพ error ที่ engine เซฟ: "{ปี}_{MMDD}_{HHMMSS}[_n].png" (ดู _unique_error_shot_path)
_ERROR_SHOT_RE = re.compile(r"^(?P<year>\d{4})_\d{4}_\d{6}(?:_\d+)?\.png$")


def _collect_run_notes(settings: Settings):
    """รวบรวมข้อมูลประกอบ sheet "หมายเหตุ" (v0.12.0): วันปลูกตามแผนของช่วงปีที่
    ตั้งไว้, ข้อความ error ของรอบล่าสุด (อยู่ใน memory — หายเมื่อปิดโปรแกรม) และ
    ภาพ error บนดิสก์ (ถาวร จึงยังโยงปีที่มีปัญหาได้แม้เปิดโปรแกรมใหม่)"""
    from file_engine.excel_writer import RunNotes

    notes = RunNotes()
    for y in range(settings.default_start_year, settings.default_end_year + 1):
        planned = [f"{d:%m%d}" for d, _ in planting_dates_for_year(settings, y)]
        if planned:
            notes.planned[y] = planned
    for ys in run_state.snapshot().years:
        if ys.status == YearRunStatus.ERROR and ys.error_message:
            notes.year_errors[ys.year] = ys.error_message
    try:
        import cropwat_options

        snap = Path(settings.output_dir) / CROPWAT_OPTIONS_SNAPSHOT
        notes.extra_meta = cropwat_options.summary(snap if snap.is_file() else None)
        if notes.extra_meta and not snap.is_file():
            notes.extra_meta.insert(0, ("CropWat Options", "ค่าปัจจุบันของเครื่อง (ไม่พบสำเนาตอนรัน)"))
    except Exception:  # noqa: BLE001
        logger.debug("อ่านค่า CropWat Options ไม่ได้", exc_info=True)
    shots_dir = Path(settings.output_dir) / "_error_dialogs"
    if shots_dir.is_dir():
        for p in sorted(shots_dir.glob("*.png")):
            m = _ERROR_SHOT_RE.match(p.name)
            if m:
                notes.error_shots.setdefault(int(m.group("year")), []).append(p.name)
    return notes


def build_word(settings: Settings) -> int:
    """สร้างไฟล์ Word รวมภาพ screenshot ทั้งหมด (โครงสร้างเหมือนไฟล์ตัวอย่างจริง
    ของผู้ใช้: บรรทัดวันที่ + ภาพตาราง + ภาพกราฟ ต่อ 1 วันปลูก) — เรียกซ้ำได้
    ทุกเมื่อเหมือน build_excel คืนจำนวนวันปลูกที่ใส่ลงเอกสาร"""
    from config import word_path
    from file_engine.word_writer import build_screenshot_doc

    return build_screenshot_doc(screenshot_dir(settings), word_path(settings))
