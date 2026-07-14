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
import threading
from pathlib import Path

from automation.cropwat_engine import CropWatEngine, PlantingDateTask
from automation.exceptions import CropWatAutomationError
from config import Settings, excel_path, planting_dates_for_year, screenshot_dir, txt_dir
from file_engine import paths as file_paths
from models import YearRunStatus
from state import run_state

logger = logging.getLogger("runner")

_run_lock = threading.Lock()
_active_thread: threading.Thread | None = None


def is_run_active() -> bool:
    with _run_lock:
        return _active_thread is not None and _active_thread.is_alive()


def start_run(years: list[int], settings: Settings) -> bool:
    """เริ่ม background thread รันปีที่ระบุ คืน False ถ้ามี run ทำงานอยู่แล้ว"""
    global _active_thread
    with _run_lock:
        if _active_thread is not None and _active_thread.is_alive():
            return False
        _active_thread = threading.Thread(
            target=_run_years, args=(years, settings), daemon=True
        )
        _active_thread.start()
        return True


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


def _run_years(years: list[int], settings: Settings) -> None:
    run_state.begin_run()
    engine = CropWatEngine()
    try:
        engine.connect()
    except CropWatAutomationError as exc:
        logger.error("connect() ล้มเหลว หยุดการรันทั้งหมด: %s", exc)
        for year in years:
            run_state.set_year_status(year, YearRunStatus.ERROR, error_message=str(exc))
        run_state.end_run()
        return

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
        run_state.end_run()
        return

    for year in years:
        if run_state.is_stop_requested():
            run_state.set_year_status(year, YearRunStatus.QUEUED)
            continue

        run_state.set_current_year(year)
        run_state.set_year_status(year, YearRunStatus.RUNNING)

        try:
            date_flags = planting_dates_for_year(settings, year)
            if not date_flags:
                raise ValueError(
                    "ไม่มีวันปลูกที่ทดลองตั้งไว้เลย (ตั้งค่าปฏิทินในหน้า Dashboard ก่อน)"
                )
            tasks = [
                PlantingDateTask(planting_date=d, capture_screenshot=shot)
                for d, shot in date_flags
            ]

            # ยืนยันจาก screenshot จริงของผู้ใช้แล้ว: ทั้งปีใช้ไฟล์ climate/rain
            # เดียวกันตลอด ไม่สลับไฟล์กลางทางแม้วันปลูกจะข้ามเดือน (ไฟล์เป็นชุด
            # ข้อมูลต่อเนื่อง 12 เดือน ครอบคลุมทุกวันปลูกที่ทดลองในปีนั้นอยู่แล้ว)
            # ใช้เดือนของวันปลูกแรก(เร็วสุด)ของปีเป็นตัวอ้างอิง resolve
            earliest_month = min(d.month for d, _ in date_flags)
            climate_file = climate_index.resolve(year, earliest_month)
            rain_file = rain_index.resolve(year, earliest_month)
        except (FileNotFoundError, NotADirectoryError, ValueError) as exc:
            run_state.set_year_status(year, YearRunStatus.ERROR, error_message=str(exc))
            continue

        result = engine.run_year(
            year=year,
            tasks=tasks,
            climate_file=climate_file,
            rain_file=rain_file,
            export_dir=txt_dir(settings),
            screenshot_dir=screenshot_dir(settings),
        )

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

    run_state.end_run()


def build_excel(settings: Settings) -> int:
    """เฟส 2 (แยกอิสระจากเฟส 1 โดยสิ้นเชิง): อ่านไฟล์ .txt ทั้งหมดที่มีอยู่จริง ณ
    ตอนนี้ → เขียนทับ sheet 'Result' ใหม่ทั้งแผ่น เรียกซ้ำได้ทุกเมื่อ คืนจำนวนปีที่
    เขียนสำเร็จ (import แบบ local กัน circular import ตอน backend startup)"""
    from file_engine.excel_writer import build_result_sheet

    return build_result_sheet(txt_dir(settings), excel_path(settings))
