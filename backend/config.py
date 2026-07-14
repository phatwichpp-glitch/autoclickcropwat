"""
config.py
=========
เก็บ path ไฟล์/โฟลเดอร์ต่างๆ ที่ระบบต้องรู้ (ตาม spec หัวข้อ "หน้าตั้งค่า")
ค่าที่ตั้งจะถูก persist ลงไฟล์ JSON (backend/data/config.json) เพื่อให้ตั้งครั้งเดียว
แล้วรอบถัดไปเปิดมาเจอค่าเดิม

ยืนยันกับผู้ใช้แล้ว (ดู mockup UX): ใช้แค่ "โฟลเดอร์ต้นทาง" เดียว (รวม climate/
rain/crop/soil ไว้ข้างใน ให้ระบบสแกนหาเอง — ดู file_engine/paths.py) กับ
"โฟลเดอร์ผลลัพธ์" เดียว (ระบบสร้าง subfolder ย่อยให้เองอัตโนมัติ):
  {output_dir}/txt/            ไฟล์ .txt ดิบจาก CropWat (เฟส 1)
  {output_dir}/screenshots/    ภาพหน้าจอ (เฉพาะวันปลูกที่เลือก capture)
  {output_dir}/Result.xlsx     ไฟล์สรุปรวม (เฟส 2)

ไม่มีไฟล์ "list วันปลูก" แยกต่างหากแบบที่ spec เดิมสมมติไว้ — ยืนยันแล้วว่าจริงๆ
คือ "วันปลูกที่ทดลอง" หลายสิบวันต่อปี (planting_calendar) ที่ผู้ใช้ปรับเองผ่าน
ปฏิทินในหน้า Dashboard ค่า default อ้างอิงจากรูปแบบจริงที่เจอใน 43 ปีของข้อมูล
ตัวอย่างผู้ใช้ (เม.ย.-ก.ค. ทุกวัน + capture ทุกวัน, ส.ค.-ก.ย. เฉพาะ 10/20/30 +
capture ทุกวันนั้น)
"""

from __future__ import annotations

import json
import sys
import threading
from datetime import date
from pathlib import Path

from pydantic import BaseModel, Field


def _app_base_dir() -> Path:
    """โฟลเดอร์ฐานสำหรับเก็บ config — ตอนรันแบบ dev (python) ใช้โฟลเดอร์ backend/
    ตามปกติ แต่ตอนถูก build เป็น .exe ด้วย PyInstaller (--onefile) ห้ามใช้
    Path(__file__) เพราะจะชี้ไปที่ temp extraction dir (sys._MEIPASS) ที่ถูกลบทิ้ง
    ทุกครั้งที่ปิดโปรแกรม — ต้องเก็บไว้ข้างๆ ตัว .exe เองแทนเพื่อให้ค่าที่ตั้งไว้อยู่ถาวร
    ข้ามการเปิดโปรแกรมแต่ละครั้ง"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


DATA_DIR = _app_base_dir() / "data"
CONFIG_PATH = DATA_DIR / "config.json"

_lock = threading.RLock()


class MonthSelection(BaseModel):
    """วันปลูกที่ทดลองในเดือนหนึ่งๆ (1-31) และ subset ที่ต้อง capture screenshot
    ด้วย (shot_days ต้องเป็น subset ของ days เสมอ)"""

    days: list[int] = Field(default_factory=list)
    shot_days: list[int] = Field(default_factory=list)


def _default_planting_calendar() -> dict[int, MonthSelection]:
    """ยืนยันจากไฟล์ .docx ตัวอย่างจริงของผู้ใช้ (1,204 screenshot ครอบคลุม 43 ปี
    1981-2023) ว่ารูปแบบนี้ใช้ซ้ำเหมือนกันทุกปี — ตั้งเป็นค่า default ให้ แต่ผู้ใช้
    แก้ไขได้อิสระทุกปีผ่านหน้า Dashboard (ไม่มีกฎตายตัวจริงจัง)"""
    dense_months = [4, 5, 6, 7]  # เม.ย.-ก.ค.: ทุกวัน 1,15
    sparse_months = [8, 9]  # ส.ค.-ก.ย.: เฉพาะ 10,20,30
    calendar: dict[int, MonthSelection] = {}
    for m in dense_months:
        calendar[m] = MonthSelection(days=[1, 15], shot_days=[1, 15])
    for m in sparse_months:
        calendar[m] = MonthSelection(days=[10, 20, 30], shot_days=[10, 20, 30])
    return calendar


class Settings(BaseModel):
    # โฟลเดอร์ต้นทางเดียวที่รวมทุกอย่าง (Clim_*/Rain_*/*.CRO/*.SOI) ไว้ข้างใน —
    # ระบบสแกนหาเองผ่าน GET /api/scan (ดู file_engine/paths.py)
    input_dir: str = ""

    # โฟลเดอร์ผลลัพธ์เดียว (ดู docstring ด้านบนสำหรับโครงสร้าง subfolder ที่สร้างให้)
    output_dir: str = ""

    # override การเลือกสถานี/ไฟล์ ถ้าสแกนแล้วเจอมากกว่า 1 ตัวเลือก (ว่าง = ให้ระบบ
    # auto-pick ตัวเดียวที่เจอ ถ้ามีแค่ตัวเดียว — ถ้าเจอหลายตัวและไม่ได้ตั้งค่าตรงนี้
    # ระบบจะ error บอกให้เข้ามาเลือก)
    climate_station_dir: str = ""
    rain_station_dir: str = ""
    crop_file: str = ""
    soil_file: str = ""

    # path ไฟล์ .exe ของ CropWat (สำรองไว้เผื่ออนาคต — ตอนนี้ backend ยังไม่เปิด
    # โปรแกรมให้อัตโนมัติ ผู้ใช้ต้องเปิดเองก่อนกด "เริ่มรันทั้งหมด")
    cropwat_exe_path: str = ""

    # ช่วงปี default ตาม spec
    default_start_year: int = 1981
    default_end_year: int = 2025

    # เดือน(1-12) -> วันปลูกที่ทดลอง + subset ที่ capture screenshot — ใช้ชุดเดียวกัน
    # ทุกปี (ปรับได้อิสระ ไม่ผูกกับปีใดปีหนึ่ง)
    planting_calendar: dict[int, MonthSelection] = Field(
        default_factory=_default_planting_calendar
    )

    # ตัวเลข "นาทีต่อขั้นตอนถ้าทำมือ" สำหรับคำนวณเวลาทำมือเทียบเท่าของงานทั้ง batch
    # (แสดงใน Dashboard) — ปรับได้ในหน้าตั้งค่าเพราะความเร็วมือแต่ละคนไม่เท่ากัน
    # ค่า default ประเมินจากลักษณะงานจริงใน workflow เดิมของผู้ใช้ (.docx 1,204 ภาพ):
    #  - ต่อปี: ไล่หาไฟล์ climate+rain ที่ถูกต้องตามกติกา shift-year ในโฟลเดอร์ซ้อน
    #    หลายชั้น + เปิดเข้า CropWat ทั้ง 2 ไฟล์
    #  - ต่อวันปลูก: ตั้งวันปลูก + คำนวณ CWR + Scheduling + เลือก table format +
    #    Print->ตั้งค่า->Save As ตั้งชื่อไฟล์ให้ถูก
    #  - ต่อ screenshot: เปิดหน้ากราฟ + capture 2 ภาพ + แปะจัดเรียงลงเอกสาร
    #  - Excel ต่อวันปลูก: คัดลอก/พิมพ์ 11 ค่าจากผลลัพธ์ลงช่องที่ถูกต้อง
    manual_minutes_per_year: float = 8.0
    manual_minutes_per_candidate: float = 5.0
    manual_minutes_per_screenshot: float = 3.0
    manual_minutes_excel_per_candidate: float = 3.0


def _ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def load_settings() -> Settings:
    with _lock:
        _ensure_data_dir()
        if not CONFIG_PATH.exists():
            settings = Settings()
            save_settings(settings)
            return settings
        raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        return Settings(**raw)


def save_settings(settings: Settings) -> None:
    with _lock:
        _ensure_data_dir()
        CONFIG_PATH.write_text(
            settings.model_dump_json(indent=2),
            encoding="utf-8",
        )


# --- helper path ที่คำนวณจาก output_dir เดียว (ไม่ใช่ field แยก) ---

def txt_dir(settings: Settings) -> Path:
    return Path(settings.output_dir) / "txt"


def screenshot_dir(settings: Settings) -> Path:
    return Path(settings.output_dir) / "screenshots"


def excel_path(settings: Settings) -> Path:
    return Path(settings.output_dir) / "Result.xlsx"


def planting_dates_for_year(settings: Settings, year: int) -> list[tuple[date, bool]]:
    """แปลง planting_calendar เป็น list ของ (วันที่, ต้อง capture screenshot ไหม)
    สำหรับปีที่กำหนด เรียงตามปฏิทิน — ข้ามวันที่ไม่มีจริงในปฏิทิน (เช่น 30 ก.พ.)"""
    results: list[tuple[date, bool]] = []
    for month in sorted(settings.planting_calendar):
        selection = settings.planting_calendar[month]
        shot_days = set(selection.shot_days)
        for day in sorted(selection.days):
            try:
                d = date(year, month, day)
            except ValueError:
                continue  # วันที่ไม่มีจริงในเดือนนั้นของปีนั้น (เช่น 30 ก.พ., 31 เม.ย.)
            results.append((d, day in shot_days))
    return results
