"""
file_engine/txt_parser.py
==========================
Parse ไฟล์ .txt ที่ CropWat print ออกมา (โหมด "Daily soil moisture balance", ASCII
file) ให้เป็นข้อมูลที่ใช้เติมลง Excel Result sheet ได้โดยตรง

ยืนยันจากไฟล์ตัวอย่างจริง (01043.TXT) แล้วว่า: การ "ทดลองปลูกวันนี้จะเกิดอะไรขึ้น"
1 ครั้ง (1 คอลัมน์ใน Result sheet) อ่านได้จาก "Totals:" section ของไฟล์เดียวกันนี้
ครบทุกอย่าง — ไม่ต้องแตะตารางรายวันเลย (Rain/Ks/Eta/... เป็นแค่รายละเอียดประกอบ
ไม่ใช่ 5 ค่าหลักที่ Excel ต้องการ):
  - "ค่า % yield reduction"                <- บรรทัด "Yield red.: X %" ในหัวไฟล์
  - "Actual water use by crop (mm)"         <- Totals: Actual water use by crop
  - "Potential water use by crop (mm)"      <- Totals: Potential water use by crop
  - "Total rainfall (mm)"                   <- Totals: Total rainfall
  - "Effective rainfall (mm)"               <- Totals: Effective rainfall
  - "Moist deficit at harvest (mm)"         <- Totals: Moist deficit at harvest
  - "Actual irrigation requirement (mm)"    <- Totals: Actual irrigation requirement
  - "Reductions in Etc (%) -Stage A/B/C/D"  <- Yield reductions: Reductions in ETc
    (4 ค่าแรกของแถวนี้ ตรงกับ Stage A,B,C,D ตามลำดับ — Season ไม่ต้องใช้เพราะ
    ซ้ำกับ Yield red. ในหัวไฟล์อยู่แล้ว)
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

_FLOAT_RE = r"-?\d+\.?\d*"


@dataclass(frozen=True)
class ParsedCandidateResult:
    """ผลลัพธ์ของ 1 ไฟล์ .txt = 1 วันปลูกที่ทดลอง = 1 คอลัมน์ใน Result sheet"""

    eto_station: str
    rain_station: str
    crop: str
    soil: str
    planting_date: str  # "dd/mm" ตามที่ CropWat พิมพ์ออกมา
    harvest_date: str
    yield_reduction_pct: float
    actual_water_use_mm: float
    potential_water_use_mm: float
    total_rainfall_mm: float
    effective_rainfall_mm: float
    moist_deficit_at_harvest_mm: float
    actual_irrigation_requirement_mm: float
    reduction_etc_stage_a: float
    reduction_etc_stage_b: float
    reduction_etc_stage_c: float
    reduction_etc_stage_d: float


class TxtParseError(Exception):
    """ไฟล์ .txt ไม่มีรูปแบบที่คาดไว้ (อาจเป็นไฟล์เพี้ยน หรือ CropWat print แบบอื่น)"""


def _search(pattern: str, text: str, *, flags: int = 0) -> re.Match:
    m = re.search(pattern, text, flags)
    if not m:
        raise TxtParseError(f"หา pattern ไม่เจอในไฟล์: {pattern!r}")
    return m


def _numbers_after_label(label: str, text: str, count: int) -> list[float]:
    """หาบรรทัดที่ขึ้นต้นด้วย label แล้วดึงตัวเลข float ตัวแรกๆ ตามจำนวน count
    ที่ตามหลังมา (ใช้กับตาราง Yield reductions ที่คอลัมน์เป็น A/B/C/D/Season)"""
    line_match = _search(rf"^\s*{re.escape(label)}\s+(.+)$", text, flags=re.MULTILINE)
    numbers = re.findall(_FLOAT_RE, line_match.group(1))
    if len(numbers) < count:
        raise TxtParseError(
            f"บรรทัด {label!r} มีตัวเลขไม่ครบ {count} ตัว (เจอ {len(numbers)}): {line_match.group(0)!r}"
        )
    return [float(n) for n in numbers[:count]]


def parse_txt(path: Path) -> ParsedCandidateResult:
    text = Path(path).read_text(encoding="cp1252", errors="replace")

    header = _search(
        r"ETo station:\s*(?P<eto>\S+)\s+Crop:\s*(?P<crop>.+?)\s{2,}Planting date:\s*(?P<planting>\d{1,2}/\d{1,2})",
        text,
    )
    header2 = _search(
        r"Rain station:\s*(?P<rain>\S+)\s+Soil:\s*(?P<soil>.+?)\s{2,}Harvest date:\s*(?P<harvest>\d{1,2}/\d{1,2})",
        text,
    )
    yield_red = _search(r"Yield red\.:\s*(?P<val>" + _FLOAT_RE + r")\s*%", text)

    def totals(label: str) -> float:
        m = _search(rf"{re.escape(label)}\s+(?P<val>{_FLOAT_RE})\s*mm", text)
        return float(m.group("val"))

    stage_a, stage_b, stage_c, stage_d = _numbers_after_label("Reductions in ETc", text, 4)

    return ParsedCandidateResult(
        eto_station=header.group("eto"),
        rain_station=header2.group("rain"),
        crop=header.group("crop").strip(),
        soil=header2.group("soil").strip(),
        planting_date=header.group("planting"),
        harvest_date=header2.group("harvest"),
        yield_reduction_pct=float(yield_red.group("val")),
        actual_water_use_mm=totals("Actual water use by crop"),
        potential_water_use_mm=totals("Potential water use by crop"),
        total_rainfall_mm=totals("Total rainfall"),
        effective_rainfall_mm=totals("Effective rainfall"),
        moist_deficit_at_harvest_mm=totals("Moist deficit at harvest"),
        actual_irrigation_requirement_mm=totals("Actual irrigation requirement"),
        reduction_etc_stage_a=stage_a,
        reduction_etc_stage_b=stage_b,
        reduction_etc_stage_c=stage_c,
        reduction_etc_stage_d=stage_d,
    )
