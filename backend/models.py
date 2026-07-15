"""
models.py
=========
Pydantic/dataclass model กลางที่ใช้ร่วมกันระหว่าง state, automation engine, และ API layer

สถานะรายปีมี 4 แบบตาม spec (หัวข้อ "หน้าจอ"):
- queued   = รอคิว
- running  = กำลังรัน
- done     = เสร็จ
- error    = มีปัญหา (ต้องมี error_message ประกอบ)
"""

from __future__ import annotations

from datetime import datetime, timezone
from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field


class YearRunStatus(str, Enum):
    QUEUED = "queued"
    RUNNING = "running"
    DONE = "done"
    ERROR = "error"


class YearStatus(BaseModel):
    year: int
    status: YearRunStatus = YearRunStatus.QUEUED
    error_message: Optional[str] = None
    exported_file: Optional[str] = None
    updated_at: str = Field(default_factory=lambda: datetime.now(timezone.utc).isoformat())

    def touch(self) -> None:
        self.updated_at = datetime.now(timezone.utc).isoformat()


class OverallRunState(str, Enum):
    IDLE = "idle"
    RUNNING = "running"
    STOPPING = "stopping"


class RunRequest(BaseModel):
    start_year: int
    end_year: int


class StateSnapshot(BaseModel):
    """สิ่งที่ส่งให้ frontend ทั้งผ่าน REST (GET /api/status) และ WebSocket"""

    overall_state: OverallRunState
    current_year: Optional[int] = None
    start_year: Optional[int] = None
    end_year: Optional[int] = None
    years: list[YearStatus] = Field(default_factory=list)
    # progress ละเอียดระดับ "วันปลูกที่ทดลอง" (ไม่ใช่แค่ระดับปี) — 1 ปีมีหลาย
    # วันปลูก การนับแค่ปีทำให้ bar กระโดดทีละก้าวใหญ่ๆ ดูเหมือนค้าง ทั้งที่งาน
    # กำลังเดินอยู่ — ใช้ขับทั้ง progress bar หน้าเว็บและ overlay ลอย
    candidate_done: int = 0
    candidate_total: int = 0
    # เวลาที่เหลือโดยประมาณ (วินาที) จากความเร็วเฉลี่ยจริงที่ทำได้ในรอบรันนี้ —
    # None = ยังประมาณไม่ได้ (ยังไม่เริ่ม/ยังไม่มีวันปลูกไหนเสร็จเลย) ดู
    # RunState._estimate_eta_seconds()
    eta_seconds: Optional[float] = None


class StationScan(BaseModel):
    """ผลสแกน 1 สถานี (climate หรือ rain) — years คือปีที่พบไฟล์อย่างน้อย 1 เดือน"""

    folder: str
    years: list[int] = Field(default_factory=list)
    missing_years: list[int] = Field(default_factory=list)


class ScanResult(BaseModel):
    """ผลสแกนโฟลเดอร์ต้นทาง (GET /api/scan) — ใช้แสดงหน้าตั้งค่าว่าเจออะไรบ้าง
    ก่อนเริ่มรันจริง"""

    climate_station_folders: list[str] = Field(default_factory=list)
    rain_station_folders: list[str] = Field(default_factory=list)
    crop_files: list[str] = Field(default_factory=list)
    soil_files: list[str] = Field(default_factory=list)
    climate: Optional[StationScan] = None
    rain: Optional[StationScan] = None
    errors: list[str] = Field(default_factory=list)
