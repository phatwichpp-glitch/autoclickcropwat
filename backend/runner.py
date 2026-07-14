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

from automation.cropwat_engine import CropWatEngine, PlantingDateTask
from automation.exceptions import CropWatAutomationError
from config import (
    Settings,
    excel_path,
    load_settings,
    planting_dates_for_year,
    screenshot_dir,
    txt_dir,
)
from file_engine import paths as file_paths
from models import YearRunStatus
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
                parse_txt(p)
            except (TxtParseError, OSError) as exc:
                logger.warning("ไฟล์ output ไม่สมบูรณ์ (จะรันวันปลูกนี้ใหม่): %s — %s", p.name, exc)
                invalid.append(p.name)
                continue
            txt_ok.add((int(m.group("year")), m.group("mmdd")))

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


def _run_years(years: list[int], settings: Settings) -> None:
    run_state.begin_run()
    engine = CropWatEngine(background_mode=settings.background_mode)
    try:
        engine.connect()
    except CropWatAutomationError as exc:
        logger.error("connect() ล้มเหลว หยุดการรันทั้งหมด: %s", exc)
        for year in years:
            run_state.set_year_status(year, YearRunStatus.ERROR, error_message=str(exc))
        run_state.end_run()
        return

    # โหมดเบื้องหลัง: เปิด watcher เฝ้าย้ายหน้าต่างชั่วคราว (Printing progress ฯลฯ)
    # ออกนอกจอตลอดการรัน — ต้อง stop เสมอตอนจบ (finally ด้านล่าง) ไม่งั้น dialog
    # ของผู้ใช้ที่กลับมาใช้ CropWat เองหลังรันจะโดนเหวี่ยงหนีจอไปด้วย
    engine.start_background_watcher()
    try:
        _run_years_inner(years, settings, engine)
    finally:
        engine.stop_background_watcher()


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

        run_state.set_current_year(year)
        run_state.set_year_status(year, YearRunStatus.RUNNING)

        def _on_candidate_done(_result, _c=completed):
            _c["n"] += 1
            run_state.set_candidate_progress(_c["n"], total_candidates)

        result = engine.run_year(
            year=year,
            tasks=tasks,
            climate_file=climate_file,
            rain_file=rain_file,
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

    run_state.end_run()


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
