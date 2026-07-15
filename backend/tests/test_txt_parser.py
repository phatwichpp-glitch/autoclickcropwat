"""
tests/test_txt_parser.py
=========================
parse_txt() คือจุดที่ผลลัพธ์ตัวเลขทั้งหมดใน Excel ไหลผ่าน — ถ้า regex พลาดแม้แต่
จุดเดียว ค่าที่ได้จะผิดแบบเงียบๆ (parse ผ่านแต่ได้เลขคนละตัว) อันตรายกว่า parse
ไม่ผ่านเสียอีก จึงต้องมี test ล็อกรูปแบบไฟล์ที่ CropWat print จริงไว้

หมายเหตุ: fixture ข้างล่างเป็นข้อมูลสมมติที่เขียนขึ้นให้ตรงรูปแบบไฟล์จริงของ
CropWat (โครงสร้าง/ตำแหน่งบรรทัดยืนยันจากไฟล์ตัวอย่างจริงแล้ว) แต่ตัวเลขและชื่อ
สถานี/พืชเป็นค่าสมมติ ไม่ใช่ข้อมูลวิจัยจริงของผู้ใช้ (ไฟล์ข้อมูลจริงถูก
.gitignore กันไว้ไม่ให้เข้า repo อยู่แล้ว)
"""

from pathlib import Path

import pytest

from file_engine.txt_parser import TxtParseError, parse_txt

SAMPLE_TXT = """\

CROP IRRIGATION SCHEDULE

ETo station:  111111           Crop: TESTCROP  (Grain)           Planting date: 01/04
Rain station:  222222          Soil: test loam                Harvest date: 29/07

Yield red.:   12.8 %

Crop scheduling options
     Timing:        No predefined irrigation
     Application:   Refill to 100 % of field capacity
     Field eff.     70  %


Table format: Daily soil moisture balance

Date     Day   Stage  Rain    Ks     Eta   Depl  Net IrrDeficit Loss  Gr. Irr Flow
                       mm   fract. mm/day    %     mm     mm     mm     mm   l/s/ha

1 Apr     1    Init    0.0   1.00    1.5     4     0.0    1.5    0.0    0.0   0.00


Totals:

  Total gross irrigation           0.0  mm    Total rainfall                501.9  mm
  Total net irrigation             0.0  mm    Effective rainfall            358.3  mm
  Total irrigation losses          0.0  mm    Total rain loss               143.6  mm

  Actual water use by crop       358.5  mm    Moist deficit at harvest        0.2  mm
  Potential water use by crop    399.3  mm    Actual irrigation requirement  41.0  mm

  Efficiency irrigation schedule     -  %     Efficiency rain                71.4  %
  Deficiency irrigation schedule  10.2  %

Yield reductions:

  Stagelabel                        A         B         C         D       Season

  Reductions in ETc                0.0       0.0       16.1      15.5      10.2   %
  Yield response factor            0.40      0.40      1.30      0.50      1.25
  Yield reduction                  0.0       0.0       21.0      7.8       12.8   %
  Cumulative yield reduction       0.0       0.0       21.0      27.1             %

Cropwat 8.0 Béta                                  14/07/26 11:24:01 AM
"""


@pytest.fixture
def sample_file(tmp_path: Path) -> Path:
    p = tmp_path / "sample.txt"
    p.write_text(SAMPLE_TXT, encoding="cp1252")
    return p


def test_parses_header_fields(sample_file: Path):
    result = parse_txt(sample_file)
    assert result.eto_station == "111111"
    assert result.rain_station == "222222"
    assert result.crop == "TESTCROP  (Grain)"
    assert result.soil == "test loam"
    assert result.planting_date == "01/04"
    assert result.harvest_date == "29/07"


def test_parses_totals(sample_file: Path):
    result = parse_txt(sample_file)
    assert result.yield_reduction_pct == 12.8
    assert result.actual_water_use_mm == 358.5
    assert result.potential_water_use_mm == 399.3
    assert result.total_rainfall_mm == 501.9
    assert result.effective_rainfall_mm == 358.3
    assert result.moist_deficit_at_harvest_mm == 0.2
    assert result.actual_irrigation_requirement_mm == 41.0


def test_parses_yield_reduction_stages(sample_file: Path):
    result = parse_txt(sample_file)
    assert result.reduction_etc_stage_a == 0.0
    assert result.reduction_etc_stage_b == 0.0
    assert result.reduction_etc_stage_c == 16.1
    assert result.reduction_etc_stage_d == 15.5


def test_missing_pattern_raises_txt_parse_error(tmp_path: Path):
    p = tmp_path / "garbage.txt"
    p.write_text("this is not a cropwat output file at all", encoding="cp1252")
    with pytest.raises(TxtParseError):
        parse_txt(p)


def test_truncated_yield_table_raises(tmp_path: Path):
    # ตัดตาราง Reductions in ETc ให้เหลือตัวเลขไม่ครบ 4 ตัว (A/B/C/D) — ต้อง fail
    # ชัดเจน ไม่ใช่ parse ผ่านแบบเงียบๆ ด้วยค่าที่ผิด
    truncated = SAMPLE_TXT.replace(
        "  Reductions in ETc                0.0       0.0       16.1      15.5      10.2   %",
        "  Reductions in ETc                0.0       0.0   %",
    )
    p = tmp_path / "truncated.txt"
    p.write_text(truncated, encoding="cp1252")
    with pytest.raises(TxtParseError):
        parse_txt(p)
