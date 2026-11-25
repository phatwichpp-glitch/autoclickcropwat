"""
file_engine/excel_writer.py
=============================
เฟส 2 ของระบบ (แยกจากเฟส 1 ที่รัน CropWat โดยสิ้นเชิง): สแกนโฟลเดอร์ .txt ที่มีอยู่
จริงทั้งหมด parse แล้วเขียนทับ sheet "Result" ใหม่ทั้งแผ่นใน Excel master — เรียกซ้ำ
ได้ทุกเมื่อโดยไม่ต้องรอเฟส 1 รันจนจบ (อ่านจากไฟล์ .txt ที่มีอยู่ ณ ตอนนั้นเท่านั้น
ไม่ได้ผูกกับ state การรันของเฟส 1 เลย)

v0.12.0 — สร้าง 3 sheet (ไม่แตะ sheet อื่นที่ผู้ใช้สร้างเองในไฟล์ปลายทาง):
  - "Result"   ตารางผลรายปี × วันปลูก (ด้านล่าง) + แถบสีบนแถว % yield reduction
  - "สรุป"     ค่าเฉลี่ย/ต่ำสุด/สูงสุดข้ามปีของแต่ละวันปลูก ไฮไลต์วันปลูกที่ดีที่สุด
               (yield reduction เฉลี่ยต่ำสุด) — ตอบคำถามวิจัย "ปลูกวันไหนดี" ได้ทันที
  - "หมายเหตุ" metadata (สถานี/พืช/ดิน) + บันทึกรายปี: วันปลูกที่มีผล/ที่ขาด +
               ข้อความ error + ชื่อภาพ error (ตาม spec: sheet หมายเหตุของไฟล์ตัวอย่าง)
  sheet "สรุป"/"หมายเหตุ" ที่ระบบสร้างมีหัวเครื่องหมาย (_AUTO_MARK) ถ้าในไฟล์มี sheet
  ชื่อเดียวกันที่ผู้ใช้เขียนเอง (ไม่มีเครื่องหมาย) จะไม่เขียนทับ — ไปเขียน "<ชื่อ> (auto)" แทน
  ("Result-1" ต้องใช้ข้อมูล ENSO จากภายนอก ยังไม่ทำ)

โครงสร้าง sheet ที่สร้าง (11 แถวต่อปี ยืนยันจากไฟล์ตัวอย่างจริง):
  - แถวหัวตาราง: คอลัมน์ A=ปี, B=ชื่อตัวแปร, C เป็นต้นไป=วันปลูกที่ทดลอง (union ของ
    ทุกวันปลูกที่เจอในไฟล์ .txt ทั้งหมด เรียงตามปฏิทิน mm/dd ใช้ชุดคอลัมน์เดียวกัน
    ทุกปี — ปีไหนไม่มีไฟล์สำหรับวันปลูกไหนก็เว้นว่างไว้)
  - ต่อปี 11 แถว: % yield reduction, Actual/Potential water use by crop,
    Total/Effective rainfall, Moist deficit at harvest, Actual irrigation
    requirement, Reductions in Etc Stage A-D
"""

from __future__ import annotations

import logging
import re
import statistics
import time
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from file_engine.txt_parser import ParsedCandidateResult, TxtParseError, parse_txt

logger = logging.getLogger("excel_writer")

METRIC_ROWS = [
    ("ค่า % yield reduction", lambda r: r.yield_reduction_pct),
    ("Actual water use by crop (mm)", lambda r: r.actual_water_use_mm),
    ("Potential water use by crop (mm)", lambda r: r.potential_water_use_mm),
    ("Total rainfall (mm)", lambda r: r.total_rainfall_mm),
    ("Effective rainfall (mm)", lambda r: r.effective_rainfall_mm),
    ("Moist deficit at harvest (mm)", lambda r: r.moist_deficit_at_harvest_mm),
    ("Actual irrigation requirement (mm)", lambda r: r.actual_irrigation_requirement_mm),
    ("Reductions in Etc (%) -Stage A", lambda r: r.reduction_etc_stage_a),
    ("Reductions in Etc (%) -Stage B", lambda r: r.reduction_etc_stage_b),
    ("Reductions in Etc (%) -Stage C", lambda r: r.reduction_etc_stage_c),
    ("Reductions in Etc (%) -Stage D", lambda r: r.reduction_etc_stage_d),
]

# ชื่อไฟล์ .txt ที่ automation engine สร้าง (ดู export_results ใน cropwat_engine.py):
# "{ปี}_{MMDD}.txt" เช่น "1981_0401.txt"
FILENAME_RE = re.compile(r"^(?P<year>\d{4})_(?P<mmdd>\d{4})\.txt$")

# หัว A1 ของ sheet ที่ระบบสร้างเอง — ใช้แยกจาก sheet ชื่อเดียวกันที่ผู้ใช้เขียนเอง
_AUTO_MARK = "สร้างอัตโนมัติโดย CropWat Auto-runner"

# คอลัมน์ของ sheet "สรุป": (หัวคอลัมน์, getter, วิธีสรุป)
SUMMARY_COLUMNS = [
    ("% yield reduction เฉลี่ย", lambda r: r.yield_reduction_pct, statistics.mean),
    ("% yield reduction ต่ำสุด", lambda r: r.yield_reduction_pct, min),
    ("% yield reduction สูงสุด", lambda r: r.yield_reduction_pct, max),
    ("Actual irrigation requirement เฉลี่ย (mm)", lambda r: r.actual_irrigation_requirement_mm, statistics.mean),
    ("Actual water use เฉลี่ย (mm)", lambda r: r.actual_water_use_mm, statistics.mean),
    ("Effective rainfall เฉลี่ย (mm)", lambda r: r.effective_rainfall_mm, statistics.mean),
]

_HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
_BEST_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
_WARN_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
_BOLD = Font(bold=True)
_MUTED = Font(italic=True, color="808080")


@dataclass
class RunNotes:
    """ข้อมูลประกอบ sheet "หมายเหตุ" ที่ไม่ได้อยู่ในไฟล์ .txt — runner รวบรวมให้
    planned: ปี -> วันปลูกตามแผน ["MMDD", ...] (ใช้หาวันปลูกที่ "ขาด")
    year_errors: ปี -> ข้อความ error ล่าสุดจากการรัน (มีเฉพาะรอบที่โปรแกรมยังเปิดอยู่)
    error_shots: ปี -> ชื่อไฟล์ภาพ error ใน _error_dialogs (อยู่บนดิสก์ถาวร)"""

    planned: dict[int, list[str]] = field(default_factory=dict)
    year_errors: dict[int, str] = field(default_factory=dict)
    error_shots: dict[int, list[str]] = field(default_factory=dict)


def _sort_key(mmdd: str) -> tuple[int, int]:
    return (int(mmdd[:2]), int(mmdd[2:]))


def _format_mmdd(mmdd: str) -> str:
    return f"{mmdd[2:]}/{mmdd[:2]}"


def _collect_results(
    txt_dir: Path, failures: dict[int, list[str]] | None = None
) -> dict[int, dict[str, ParsedCandidateResult]]:
    """สแกน txt_dir หาไฟล์ .txt ทั้งหมด parse ทีละไฟล์ — ไฟล์ที่ parse ไม่ได้แค่
    ข้าม + log warning ไป ไม่ทำให้ทั้งเฟส 2 ล้ม (เหมือนหลักการเดียวกับเฟส 1)
    failures (ถ้าส่งมา): เก็บชื่อไฟล์ที่ parse ไม่ได้แยกตามปี ไว้ลง sheet หมายเหตุ"""
    parsed: dict[int, dict[str, ParsedCandidateResult]] = {}
    for path in sorted(Path(txt_dir).glob("*.txt")):
        m = FILENAME_RE.match(path.name)
        if not m:
            logger.warning("ข้ามไฟล์ชื่อไม่ตรงรูปแบบที่คาด: %s", path.name)
            continue
        year = int(m.group("year"))
        mmdd = m.group("mmdd")
        try:
            result = parse_txt(path)
        except TxtParseError as exc:
            logger.warning("parse ไฟล์ %s ไม่สำเร็จ: %s", path.name, exc)
            if failures is not None:
                failures.setdefault(year, []).append(path.name)
            continue
        parsed.setdefault(year, {})[mmdd] = result
    return parsed


def _load_or_create_workbook(path: Path) -> Workbook:
    if path.exists():
        return load_workbook(path)
    wb = Workbook()
    wb.remove(wb.active)  # ลบ sheet เริ่มต้น "Sheet" ที่ openpyxl สร้างมาให้เฉยๆ
    return wb


def _replace_owned_sheet(wb: Workbook, name: str, index: int):
    """สร้าง sheet ใหม่ชื่อ name แทนของเดิม — เฉพาะของเดิมที่ระบบสร้างเอง (A1 มี
    _AUTO_MARK) ถ้าเป็น sheet ที่ผู้ใช้เขียนเอง ห้ามลบ ไปใช้ชื่อ "<name> (auto)" แทน"""
    if name in wb.sheetnames and _AUTO_MARK not in str(wb[name]["A1"].value or ""):
        name = f"{name} (auto)"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name, min(index, len(wb.sheetnames)))
    ws["A1"] = f"{name} — {_AUTO_MARK} {time.strftime('%d/%m/%Y %H:%M')} (สร้างใหม่ทุกครั้งที่กดสร้าง Excel ห้ามแก้ในแผ่นนี้)"
    ws["A1"].font = _MUTED
    return ws


def _write_result_sheet(wb: Workbook, parsed: dict, all_mmdd: list[str]) -> None:
    if "Result" in wb.sheetnames:
        del wb["Result"]
    ws = wb.create_sheet("Result", 0)

    ws.cell(row=1, column=1, value="ปี").font = _BOLD
    ws.cell(row=1, column=2, value="ตัวแปร").font = _BOLD
    for col_idx, mmdd in enumerate(all_mmdd, start=3):
        cell = ws.cell(row=1, column=col_idx, value=_format_mmdd(mmdd))
        cell.font = _BOLD
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    last_col = get_column_letter(2 + len(all_mmdd))
    yield_ranges: list[str] = []
    row_idx = 2
    for year in sorted(parsed):
        year_data = parsed[year]
        first_row_of_block = row_idx
        yield_ranges.append(f"C{row_idx}:{last_col}{row_idx}")
        for label, getter in METRIC_ROWS:
            ws.cell(row=row_idx, column=2, value=label)
            for col_idx, mmdd in enumerate(all_mmdd, start=3):
                result = year_data.get(mmdd)
                if result is not None:
                    ws.cell(row=row_idx, column=col_idx, value=getter(result))
            row_idx += 1
        ws.cell(row=first_row_of_block, column=1, value=year).font = _BOLD
        ws.merge_cells(
            start_row=first_row_of_block, start_column=1, end_row=row_idx - 1, end_column=1
        )
        ws.cell(row=first_row_of_block, column=1).alignment = Alignment(vertical="center")

    # แถบสีบนแถว % yield reduction ทุกปี (สเกลเดียวกันทั้งแผ่น เทียบข้ามปีได้):
    # เขียว = ผลผลิตลดน้อย, แดง = ลดมาก — ไม่เปลี่ยนค่าในเซลล์ แค่ช่วยให้อ่านเร็ว
    ws.conditional_formatting.add(
        " ".join(yield_ranges),
        ColorScaleRule(
            start_type="min", start_color="63BE7B",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max", end_color="F8696B",
        ),
    )

    ws.column_dimensions["B"].width = 32
    for col_idx in range(3, 3 + len(all_mmdd)):
        ws.column_dimensions[get_column_letter(col_idx)].width = 10
    ws.freeze_panes = "C2"


def _write_summary_sheet(wb: Workbook, parsed: dict, all_mmdd: list[str]) -> None:
    """sheet "สรุป": 1 แถวต่อวันปลูก — สถิติข้ามทุกปีที่มีผล ไฮไลต์วันปลูกที่
    % yield reduction เฉลี่ยต่ำสุด (เป็นค่าคงที่ ไม่ใช่สูตร — สร้างใหม่ทุกครั้งอยู่แล้ว)"""
    ws = _replace_owned_sheet(wb, "สรุป", 1)
    headers = ["วันปลูก", "จำนวนปีที่มีผล"] + [h for h, _, _ in SUMMARY_COLUMNS]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = _BOLD
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    rows: list[tuple[str, list]] = []
    for mmdd in all_mmdd:
        results = [parsed[y][mmdd] for y in sorted(parsed) if mmdd in parsed[y]]
        stats = [round(agg([getter(r) for r in results]), 2) for _, getter, agg in SUMMARY_COLUMNS]
        rows.append((mmdd, [len(results)] + stats))

    best = min(rows, key=lambda r: r[1][1])[0] if rows else None
    for row_idx, (mmdd, values) in enumerate(rows, start=4):
        ws.cell(row=row_idx, column=1, value=_format_mmdd(mmdd)).alignment = Alignment(horizontal="center")
        for col_idx, v in enumerate(values, start=2):
            ws.cell(row=row_idx, column=col_idx, value=v)
        if mmdd == best:
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col_idx).fill = _BEST_FILL
                ws.cell(row=row_idx, column=col_idx).font = _BOLD

    if best is not None:
        ws.cell(
            row=len(rows) + 5, column=1,
            value=f"แถวสีเขียว = วันปลูกที่ % yield reduction เฉลี่ยต่ำสุด ({_format_mmdd(best)}) "
                  "— ค่าเฉลี่ยคิดจากเฉพาะปีที่มีผล ดูคอลัมน์ \"จำนวนปีที่มีผล\" ประกอบ",
        ).font = _MUTED

    ws.row_dimensions[3].height = 45
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 12
    for col_idx in range(3, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16
    ws.freeze_panes = "B4"


def _write_notes_sheet(
    wb: Workbook, parsed: dict, notes: RunNotes, parse_failures: dict[int, list[str]]
) -> None:
    """sheet "หมายเหตุ": metadata + บันทึกปัญหารายปี (ตาม spec — sheet หมายเหตุของ
    ไฟล์ตัวอย่างจริงของผู้ใช้เก็บสถานี/ดิน/พืช + ปัญหาข้อมูลรายปี)"""
    ws = _replace_owned_sheet(wb, "หมายเหตุ", 2)
    all_results = [r for yd in parsed.values() for r in yd.values()]

    def distinct(getter) -> str:
        return ", ".join(sorted({getter(r) for r in all_results})) or "-"

    meta = [
        ("ETo station", distinct(lambda r: r.eto_station)),
        ("Rain station", distinct(lambda r: r.rain_station)),
        ("Crop", distinct(lambda r: r.crop)),
        ("Soil", distinct(lambda r: r.soil)),
        ("ปีที่มีผล", f"{min(parsed)}–{max(parsed)} ({len(parsed)} ปี)"),
        ("วันปลูกที่มีผลทั้งหมด", f"{len(all_results):,}"),
    ]
    for row_idx, (k, v) in enumerate(meta, start=3):
        ws.cell(row=row_idx, column=1, value=k).font = _BOLD
        ws.cell(row=row_idx, column=2, value=v)

    head_row = len(meta) + 4
    headers = ["ปี", "วันปลูกที่มีผล", "วันปลูกที่ขาด", "หมายเหตุ"]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=head_row, column=col_idx, value=h)
        cell.font = _BOLD
        cell.fill = _HEADER_FILL

    years = sorted(
        set(parsed) | set(notes.planned) | set(notes.year_errors)
        | set(notes.error_shots) | set(parse_failures)
    )
    row_idx = head_row + 1
    for year in years:
        have = set(parsed.get(year, {}))
        planned = notes.planned.get(year, [])
        missing = [m for m in planned if m not in have]
        remarks: list[str] = []
        if notes.year_errors.get(year):
            remarks.append(notes.year_errors[year])
        if parse_failures.get(year):
            remarks.append("ไฟล์ .txt อ่านไม่ได้: " + ", ".join(parse_failures[year]))
        if notes.error_shots.get(year):
            remarks.append("ภาพ error (โฟลเดอร์ _error_dialogs): " + ", ".join(notes.error_shots[year]))
        if not planned and not have and not remarks:
            continue
        ws.cell(row=row_idx, column=1, value=year)
        ws.cell(row=row_idx, column=2, value=f"{len(have)}" + (f"/{len(planned)}" if planned else ""))
        ws.cell(row=row_idx, column=3, value=", ".join(_format_mmdd(m) for m in missing) or "-")
        remark_cell = ws.cell(row=row_idx, column=4, value=" | ".join(remarks) or "-")
        remark_cell.alignment = Alignment(wrap_text=True, vertical="top")
        if missing or remarks:
            for col_idx in range(1, 5):
                ws.cell(row=row_idx, column=col_idx).fill = _WARN_FILL
        row_idx += 1

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 90
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)


def build_result_sheet(txt_dir: Path, output_xlsx: Path, notes: RunNotes | None = None) -> int:
    """สร้าง/เขียนทับ sheet 'Result' + 'สรุป' + 'หมายเหตุ' ใน output_xlsx จากไฟล์
    .txt ทั้งหมดใน txt_dir คืนค่าจำนวนปีที่เขียนสำเร็จ"""
    parse_failures: dict[int, list[str]] = {}
    parsed = _collect_results(txt_dir, parse_failures)
    if not parsed:
        raise FileNotFoundError(f"ไม่พบไฟล์ .txt ที่ parse ได้เลยในโฟลเดอร์: {txt_dir}")

    all_mmdd = sorted({mmdd for year_data in parsed.values() for mmdd in year_data}, key=_sort_key)

    wb = _load_or_create_workbook(Path(output_xlsx))
    _write_result_sheet(wb, parsed, all_mmdd)
    _write_summary_sheet(wb, parsed, all_mmdd)
    _write_notes_sheet(wb, parsed, notes or RunNotes(), parse_failures)
    wb.active = 0

    Path(output_xlsx).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_xlsx)
    return len(parsed)
