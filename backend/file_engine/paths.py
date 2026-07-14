"""
file_engine/paths.py
=====================
สแกนโฟลเดอร์ข้อมูลต้นทาง (input_dir) หา climate/rain stations, crop, soil files
และ resolve path ไฟล์ climate/rain ที่ถูกต้องสำหรับปี+เดือนปลูกที่กำหนด ตามกติกา
shift-year ที่ระบุใน spec

ยืนยันจากโฟลเดอร์ข้อมูลจริงแล้ว (d:\\Cropwat\\autoclickcropwat\\cropwat\\): โครงสร้าง
sub-folder ข้างในสถานีหนึ่งๆ ไม่ consistent เลย — บางเดือนมี decade-range subfolder
ซ้อนอยู่ (เช่น "04_Apr/1981-1990/"), บางเดือนไม่มี (ไฟล์อยู่ตรงๆ ใน "05_May/"),
บางโฟลเดอร์ซ้อนชื่อตัวเองซ้ำ (เช่น "04_Rain_450006Apr/04_Rain_450006Apr/") — แต่
ชื่อไฟล์เองมีปี+เดือนครบเสมอไม่ว่าจะซ้อนลึกแค่ไหน เลยไม่ต้องสนใจโครงสร้าง folder
เลย ใช้วิธี scan ไฟล์แบบ recursive (rglob) ทั่วทั้ง station folder แล้ว parse
ปี/เดือนจากชื่อไฟล์โดยตรงแทน (ตรงกับที่ spec แนะนำ: list ไฟล์จริงแล้ว match ด้วย
regex แทนการ generate ชื่อไฟล์ขึ้นมาเอง)
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

MONTH_ABBR = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]

# ยืนยันแล้วจากไฟล์จริง: climate เป็นตัวเล็กทั้งหมด, rain ตัวใหญ่นำเดือน — ใช้
# re.IGNORECASE เผื่อความไม่ consistent เพิ่มเติมที่ยังไม่เจอ
CLIMATE_FILE_RE = re.compile(
    r"^(?P<station>\d+)clim_(?P<year>\d{4})(?P<mon>[a-z]{3})\.ped$", re.IGNORECASE
)
RAIN_FILE_RE = re.compile(
    r"^rain_(?P<station>\d+)_(?P<year>\d{4})(?P<mon>[a-z]{3})\.crd$", re.IGNORECASE
)

IGNORED_FILENAMES = {"desktop.ini"}


@dataclass(frozen=True)
class StationIndex:
    """(ปี, เดือน) -> path ไฟล์ ของสถานีหนึ่งๆ (climate หรือ rain) — สร้างจากการ
    scan ไฟล์จริงทั้งหมดในโฟลเดอร์สถานี ไม่ได้ผูกกับโครงสร้าง sub-folder เลย"""

    files_by_year_month: dict[tuple[int, int], Path] = field(default_factory=dict)

    def available_years(self) -> list[int]:
        return sorted({y for y, _m in self.files_by_year_month})

    def available_months(self, year: int) -> list[int]:
        return sorted(m for (y, m) in self.files_by_year_month if y == year)

    def resolve(self, planting_year: int, planting_month: int) -> Path:
        """กติกา shift-year ตาม spec: เลือกโฟลเดอร์ (ในที่นี้คือไฟล์) ที่เดือน
        เริ่มต้น <= planting_month และมากที่สุดในปีนั้น ถ้าไม่มีเลยในปีนั้น ให้ข้าม
        ไปหาเดือนเริ่มต้นมากที่สุดของปีก่อนหน้าแทน"""
        months_this_year = [m for m in self.available_months(planting_year) if m <= planting_month]
        if months_this_year:
            chosen_month = max(months_this_year)
            chosen_year = planting_year
        else:
            months_prev_year = self.available_months(planting_year - 1)
            if not months_prev_year:
                raise FileNotFoundError(
                    f"ไม่พบไฟล์เดือนเริ่มต้นที่ใช้ได้ทั้งปี {planting_year} "
                    f"และปี {planting_year - 1} (เดือนปลูก {planting_month})"
                )
            chosen_month = max(months_prev_year)
            chosen_year = planting_year - 1
        return self.files_by_year_month[(chosen_year, chosen_month)]


def _index_station(station_dir: Path, pattern: re.Pattern, glob_ext: str) -> StationIndex:
    files: dict[tuple[int, int], Path] = {}
    for path in station_dir.rglob(glob_ext):
        if path.name.lower() in IGNORED_FILENAMES:
            continue
        m = pattern.match(path.name)
        if not m:
            continue
        year = int(m.group("year"))
        month = MONTH_ABBR.index(m.group("mon").lower()) + 1
        files[(year, month)] = path
    return StationIndex(files_by_year_month=files)


def index_climate_station(station_dir: Path) -> StationIndex:
    return _index_station(station_dir, CLIMATE_FILE_RE, "*.ped")


def index_rain_station(station_dir: Path) -> StationIndex:
    return _index_station(station_dir, RAIN_FILE_RE, "*.crd")


def find_station_folders(root: Path, prefix: str) -> list[Path]:
    """หาโฟลเดอร์สถานีใน root ที่ชื่อขึ้นต้นด้วย prefix (เช่น 'Clim_' หรือ 'Rain_')"""
    if not root.is_dir():
        return []
    return sorted(
        p for p in root.iterdir() if p.is_dir() and p.name.lower().startswith(prefix.lower())
    )


def find_files_by_extension(root: Path, ext: str) -> list[Path]:
    """สแกน root ทั้งหมดแบบ recursive หาไฟล์นามสกุลที่กำหนด (ใช้หาไฟล์ crop/.CRO
    และ soil/.SOI ซึ่งอยู่ในโฟลเดอร์ชื่ออะไรก็ได้ ไม่มี prefix ตายตัวแบบ Clim_/Rain_)"""
    if not root.is_dir():
        return []
    ext = ext if ext.startswith(".") else f".{ext}"
    return sorted(
        p for p in root.rglob(f"*{ext}") if p.name.lower() not in IGNORED_FILENAMES
    )
