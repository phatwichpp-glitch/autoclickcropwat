"""
tests/test_paths_shift_year.py
================================
StationIndex.resolve() คือจุดที่เคยแก้ผิดมาแล้วรอบหนึ่ง (v0.5.23 เผลอใส่
forward-fallback แล้วต้องแก้กลับใน v0.5.24) — ล็อกพฤติกรรมที่ถูกต้องไว้ด้วย
test กัน regression กลับไปเป็นแบบเดิมอีกโดยไม่มีใครรู้ตัว

กติกาที่ต้อง lock ไว้:
1. เดือนตรงเป๊ะ (ปี+เดือนเดียวกัน) ต้องถูกเลือกก่อนเสมอ
2. ถ้าไม่มีเดือนตรง ให้ใช้เดือน "ก่อนหน้า" ที่ใกล้ที่สุด ค้นได้ทุกปีย้อนหลัง
   (ไม่ใช่แค่ 1 ปีก่อนหน้า)
3. ห้ามใช้ไฟล์เดือน "ถัดไป" (อนาคต) เด็ดขาด ไม่ว่ากรณีใด
4. ถ้าหาเดือนก่อนหน้าไม่เจอเลย (สถานีเริ่มมีข้อมูลทีหลังเดือนที่ขอ) ต้อง raise
   FileNotFoundError ไม่ใช่เดาให้เงียบๆ
5. สถานีที่ไม่มีไฟล์เลย ต้อง raise FileNotFoundError ทันที
"""

from pathlib import Path

import pytest

from file_engine.paths import StationIndex


def _index(*year_months: tuple[int, int]) -> StationIndex:
    return StationIndex(
        files_by_year_month={ym: Path(f"fake_{ym[0]}_{ym[1]:02d}.PED") for ym in year_months}
    )


def test_exact_match_preferred():
    idx = _index((1981, 3), (1981, 4), (1981, 5))
    assert idx.resolve(1981, 4).name == "fake_1981_04.PED"


def test_falls_back_to_nearest_earlier_month_same_year():
    idx = _index((1981, 4), (1981, 8))
    # ขอเดือน 9 (กันยา) ไม่มี -> ต้องได้ 8 (สิงหา) ไม่ใช่ 4
    assert idx.resolve(1981, 9).name == "fake_1981_08.PED"


def test_searches_backward_across_multiple_years_not_just_one():
    # สถานีมีข้อมูลแค่ปี 1975 (ไม่ใช่แค่ 1 ปีก่อนหน้าปีที่ขอ) -> ต้องยังหาเจอ
    idx = _index((1975, 6),)
    result = idx.resolve(1981, 2)
    assert result.name == "fake_1975_06.PED"


def test_never_uses_a_future_month():
    # มีเดือนถัดไป (พ.ค.) แต่ไม่มีเดือนก่อนหน้าเลย -> ต้อง fail ไม่ใช่ไปเอาเดือน
    # ถัดไปมาใช้ (นี่คือบั๊กที่ v0.5.23 เคยทำผิดแล้วผู้ใช้ยืนยันว่าห้ามทำ)
    idx = _index((1981, 5),)
    with pytest.raises(FileNotFoundError):
        idx.resolve(1981, 2)


def test_empty_station_raises():
    idx = _index()
    with pytest.raises(FileNotFoundError):
        idx.resolve(1981, 4)


def test_picks_closest_backward_not_furthest():
    idx = _index((1979, 1), (1980, 6), (1981, 3))
    # ขอเดือน 1981/4 -> ใกล้สุดที่ก่อนหน้าคือ 1981/3 ไม่ใช่ 1980/6 หรือ 1979/1
    assert idx.resolve(1981, 4).name == "fake_1981_03.PED"
