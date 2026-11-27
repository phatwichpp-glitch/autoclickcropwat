"""
tests/test_discovery.py
=======================
v1.4.0 — ตัวค้นหาไฟล์ต้นทางแบบไม่ยึดชื่อโฟลเดอร์/ชื่อไฟล์ตายตัว (file_engine/discovery.py)
และตัวตรวจลำดับ วัน/เดือน จากวันเก็บเกี่ยว + อายุพืช (file_engine/txt_parser.py)
ไฟล์ในเทสต์เป็นของสมมติ (หัวไฟล์เลียนรูปแบบ CropWat 8.0 ไม่ใช่ข้อมูลวิจัยจริง)
"""

from pathlib import Path

import pytest

import runner
from config import Settings
from file_engine import discovery
from file_engine.txt_parser import crop_cycle_days, detect_date_order, planting_date_matches


def _climate(path: Path, station="111111", year=1990, body="  1.0\n") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        f"CROPWAT 8.0 Climate data\n 0   1\nThailand\n{station}\n  7.00\n  14.47  100.00\n{year}\n{body}",
        encoding="latin-1",
    )
    return path


def _rain(path: Path, station="222222", year=1990, body="  0.0\n") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        f"CROPWAT 8.0 Rain data\n{station}\n  2  1\n  0  0  0  0  0  0\n{year}\n{body}", encoding="latin-1"
    )
    return path


def _crop(path: Path, stages="20 30 40 30") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(f"CROPWAT 8.0 Crop data\nTESTCROP\n{stages} \n  0.30  1.20  0.35\n", encoding="latin-1")
    return path


# ---------------------------------------------------------------- อ่านชื่อไฟล์

@pytest.mark.parametrize("name,year,month", [
    ("425201clim_1981apr.PED", 1981, 4),
    ("Rain_450006_1981Apr.CRD", 1981, 4),
    ("station_1981-04.ped", 1981, 4),
    ("198104.ped", 1981, 4),
    ("04_1981.ped", 1981, 4),
    ("April 1981 daily.ped", 1981, 4),
    ("clim 2524 เม.ย..ped", 1981, 4),
    ("ฝน_เมษายน_2524.crd", 1981, 4),
    ("1981_sept.ped", 1981, 9),
])
def test_year_and_month_from_many_name_styles(name, year, month):
    info = discovery.parse_name(Path("C:/data") / name)
    assert info.years == [year] and info.months == [month]


def test_station_number_is_not_mistaken_for_year_or_month():
    info = discovery.parse_name(Path("C:/data/Rain_450006_1981Apr.CRD"))
    assert info.station == "450006" and info.years == [1981]


def test_month_and_year_from_folders_when_name_has_none(tmp_path):
    f = tmp_path / "ข้อมูล" / "1981" / "04_Apr" / "daily.ped"
    f.parent.mkdir(parents=True)
    f.write_text("x")
    info = discovery.parse_name(f, tmp_path)
    assert (info.folder_year, info.folder_month) == (1981, 4)
    # ช่วงปี "1981-1990" ไม่ใช่ปีเดียว — ห้ามเอามาเป็นปีของไฟล์
    g = tmp_path / "04_Apr" / "1981-1990" / "x.ped"
    g.parent.mkdir(parents=True)
    g.write_text("x")
    assert discovery.parse_name(g, tmp_path).folder_year is None


# ---------------------------------------------------------------- ค้นทั้งโฟลเดอร์

def test_finds_files_in_any_folder_layout(tmp_path):
    """ไม่มีโฟลเดอร์ Clim_/Rain_ เลย ชื่อไฟล์คนละแบบ อยู่ลึกคนละชั้น — ต้องหาเจอครบ"""
    _climate(tmp_path / "ภูมิอากาศ" / "ชุดที่ 1" / "1990-04.ped", year=1990)
    _climate(tmp_path / "ภูมิอากาศ" / "ชุดที่ 1" / "May_1990.ped", year=1990)
    _rain(tmp_path / "ฝนรายวัน" / "2533" / "เม.ย." / "data.crd", year=1990)
    _crop(tmp_path / "พืช" / "MAIZE.CRO")
    (tmp_path / "ดิน").mkdir()
    (tmp_path / "ดิน" / "loam.SOI").write_text("CROPWAT 8.0 Soil Data\nloam\n")

    found = discovery.discover(tmp_path)
    assert set(found.stations["climate"]) == {"111111"}
    assert set(found.stations["climate"]["111111"]) == {(1990, 4), (1990, 5)}
    assert set(found.stations["rain"]["222222"]) == {(1990, 4)}
    assert [p.name for p in found.crops] == ["MAIZE.CRO"] and [p.name for p in found.soils] == ["loam.SOI"]
    assert found.unreadable == [] and found.warnings == []
    assert found.index("climate").resolve(1990, 5).name == "May_1990.ped"


def test_output_folder_inside_input_is_ignored(tmp_path):
    _climate(tmp_path / "in" / "c_1990apr.ped")
    _climate(tmp_path / "in" / "output" / "copy_1991apr.ped", year=1991)
    found = discovery.discover(tmp_path / "in", exclude=(tmp_path / "in" / "output",))
    assert set(found.stations["climate"]["111111"]) == {(1990, 4)}


def test_unreadable_files_are_reported_with_reason_not_guessed(tmp_path):
    _climate(tmp_path / "a" / "mystery.ped", year=1990)             # ไม่มีเดือนที่ไหนเลย
    _climate(tmp_path / "a" / "1990apr_1990may.ped", year=1990)     # 2 เดือนในชื่อ
    _rain(tmp_path / "a" / "wrongext_1990apr.ped", year=1990)       # เนื้อเป็น rain แต่นามสกุล .ped
    (tmp_path / "a" / "noyear_apr.ped").write_text("not a cropwat file")  # ไม่มีปีที่ไหนเลย
    found = discovery.discover(tmp_path)
    reasons = {u["path"].split("\\")[-1].split("/")[-1]: u["reason"] for u in found.unreadable}
    assert set(reasons) == {"mystery.ped", "1990apr_1990may.ped", "wrongext_1990apr.ped", "noyear_apr.ped"}
    assert "เดือน" in reasons["mystery.ped"] and "มากกว่า 1" in reasons["1990apr_1990may.ped"]
    assert "Rain" in reasons["wrongext_1990apr.ped"] and "ปี" in reasons["noyear_apr.ped"]
    assert found.stations["climate"] == {}
    with pytest.raises(FileNotFoundError, match="อ่านปี/เดือนไม่ออก"):
        found.pick_station("climate")


def test_name_year_wins_but_header_mismatch_is_flagged(tmp_path):
    _climate(tmp_path / "c_1994apr.ped", year=1984)
    found = discovery.discover(tmp_path)
    assert set(found.stations["climate"]["111111"]) == {(1994, 4)}
    assert len(found.warnings) == 1 and "1984" in found.warnings[0]


def test_duplicates_identical_are_silent_different_are_flagged(tmp_path):
    _climate(tmp_path / "main" / "c_1990apr.ped")
    _climate(tmp_path / "copy of main" / "c_1990apr.ped")                    # เหมือนกันทุกไบต์
    _climate(tmp_path / "main" / "c_1990may.ped", body="  1.0\n")
    _climate(tmp_path / "backup" / "c_1990may.ped", body="  9.9\n")          # ต่างกัน อยู่ในโฟลเดอร์สำรอง
    found = discovery.discover(tmp_path)
    assert found.duplicate_copies == 1
    assert len(found.warnings) == 1 and "ไม่เหมือนกัน" in found.warnings[0]
    assert found.stations["climate"]["111111"][(1990, 5)].path.parent.name == "main"


def test_multiple_stations_must_be_chosen_by_id_or_folder(tmp_path):
    _climate(tmp_path / "mix" / "a_1990apr.ped", station="111111")
    _climate(tmp_path / "mix" / "b_1990apr.ped", station="333333")
    _climate(tmp_path / "north" / "n_1990apr.ped", station="555555")
    found = discovery.discover(tmp_path)
    with pytest.raises(ValueError, match="มากกว่า 1 สถานี"):
        found.pick_station("climate")
    assert found.pick_station("climate", "333333") == "333333"
    assert found.pick_station("climate", str(tmp_path / "north")) == "555555"
    with pytest.raises(ValueError, match="รหัสสถานี"):
        found.pick_station("climate", str(tmp_path / "mix"))
    with pytest.raises(FileNotFoundError):
        found.pick_station("climate", "999999")


def test_station_falls_back_to_name_then_top_folder(tmp_path):
    for name in ("777777_1990apr.ped", "plain_1990may.ped"):
        p = tmp_path / "สถานีเหนือ" / name
        p.parent.mkdir(exist_ok=True)
        p.write_text("old file without CropWat 8.0 header")
    found = discovery.discover(tmp_path)
    assert set(found.stations["climate"]) == {"777777", "สถานีเหนือ"}


def test_runner_uses_discovery_and_crop_choice_by_name(tmp_path):
    _climate(tmp_path / "x" / "c_1990apr.ped")
    _rain(tmp_path / "y" / "r_1990apr.crd")
    _crop(tmp_path / "crops" / "MAIZE.CRO")
    _crop(tmp_path / "crops" / "RICE.CRO")
    (tmp_path / "s.SOI").write_text("CROPWAT 8.0 Soil Data\n")
    s = Settings(input_dir=str(tmp_path), output_dir=str(tmp_path / "out"))
    climate_index, rain_index = runner._station_indexes(s)
    assert climate_index.resolve(1990, 6).name == "c_1990apr.ped"
    with pytest.raises(ValueError, match="มากกว่า 1 ไฟล์"):
        runner._crop_soil_files(s)
    s.crop_file = "rice.cro"
    assert runner._crop_soil_files(s)[0].name == "RICE.CRO"


# ---------------------------------------------------------------- ลำดับ วัน/เดือน

def test_crop_cycle_days_from_cro(tmp_path):
    assert crop_cycle_days(_crop(tmp_path / "m.CRO")) == 120
    assert crop_cycle_days(_crop(tmp_path / "odd.CRO", stages="20 30 40")) is None
    assert crop_cycle_days(tmp_path / "missing.CRO") is None


def test_date_order_detected_from_harvest_date():
    # ของจริงจากข้อมูลตัวอย่าง: ปลูก 01/04 เก็บเกี่ยว 29/07 พืช 120 วัน = วัน/เดือน
    assert detect_date_order("01/04", "29/07", 1981, 120) == "dmy"
    # เครื่องที่เขียน เดือน/วัน: ปลูก 5 เม.ย. → "04/05", เก็บเกี่ยว 2 ส.ค. → "08/02"
    assert detect_date_order("04/05", "08/02", 1981, 120) == "mdy"
    assert detect_date_order("01/04", "29/07", 1981, None) is None
    assert detect_date_order("01/04", "15/09", 1981, 120) is None  # ไม่เข้าทางไหนเลย = ไม่ตัดสิน


def test_swapped_day_month_is_caught_when_evidence_exists():
    # สั่ง 4 พ.ค. (ไฟล์ชื่อ 0504) แต่ CropWat คำนวณ 5 เม.ย. แล้วเขียน "04/05" + เก็บเกี่ยว "08/02"
    assert planting_date_matches("0504", "04/05") is True                       # ดูแค่ข้อความ: แยกไม่ออก
    assert planting_date_matches("0504", "04/05", "08/02", 1981, 120) is False  # มีหลักฐาน: จับได้
    # สั่ง 4 พ.ค. ได้ 4 พ.ค. จริง (วัน/เดือน): เก็บเกี่ยว 31 ส.ค.
    assert planting_date_matches("0504", "04/05", "31/08", 1981, 120) is True
    # ไม่มีอายุพืช = ผ่อนปรนเหมือนเดิม
    assert planting_date_matches("0504", "04/05", "08/02", 1981, None) is True


# ---------------------------------------------------------------- โครงสร้างจริงจากผู้ใช้งาน (2026-11-27)

def test_real_user_layout_month_only_in_folder_name(tmp_path):
    """Rain_381017/381017/01_Rain381017Jan/01/Rain_381017_1981.CRD — เดือนอยู่ในชื่อโฟลเดอร์
    ไม่ใช่ชื่อไฟล์ และโฟลเดอร์ย่อย "01", "02" เป็นลำดับกลุ่มปี ไม่ใช่เดือน (ตัวค้นหาเดิมหาไม่เจอ
    เพราะชื่อไฟล์ไม่มีเดือน)"""
    base = tmp_path / "RUN" / "Rain_381017" / "381017"
    _rain(base / "01_Rain381017Jan" / "01" / "Rain_381017_1981.CRD", station="381017", year=1981)
    _rain(base / "01_Rain381017Jan" / "02" / "Rain_381017_1991.CRD", station="381017", year=1991)
    _rain(base / "05_Rain381017May" / "01" / "Rain_381017_1981.CRD", station="381017", year=1981, body="  5.0\n")
    _rain(base / "10_Rain381017Oct" / "03" / "Rain_381017_2001.CRD", station="381017", year=2001)

    found = discovery.discover(tmp_path / "RUN")
    assert found.unreadable == [] and found.warnings == []
    files = found.stations["rain"]["381017"]
    # กลุ่มปี "02" ต้องไม่ถูกอ่านเป็นเดือน ก.พ. — เดือนมาจากชื่อโฟลเดอร์ที่มีชื่อเดือน (Jan)
    assert set(files) == {(1981, 1), (1991, 1), (1981, 5), (2001, 10)}
    index = found.index("rain")
    # ไม่มีโฟลเดอร์ มิ.ย.–ก.ย.: ปลูก ก.ค. 1981 ใช้ไฟล์ พ.ค. 1981 (เดือนก่อนหน้าที่ใกล้ที่สุด)
    assert index.resolve(1981, 7).parent.parent.name == "05_Rain381017May"


def test_bare_number_folder_is_month_only_when_no_month_name_anywhere(tmp_path):
    _climate(tmp_path / "04" / "c_1990.ped")
    found = discovery.discover(tmp_path)
    assert set(found.stations["climate"]["111111"]) == {(1990, 4)}


def test_files_status_reports_current_and_finished(tmp_path, monkeypatch):
    out = tmp_path / "out"
    out.mkdir()
    (out / runner.RUN_FILES_LOG).write_text(
        '{"year": 1981, "mmdd": "0401", "climate": "C:/in/c_1981apr.PED", "rain": "C:/in/r_1981apr.CRD", '
        '"crop": "C:/in/M.CRO", "soil": "C:/in/s.SOI", "at": "2026-11-27 10:00:00"}\n'
        '{"year": 1981, "mmdd": "0415", "climate": "C:/in/c_1981apr.PED", "rain": "C:/in/r_1981apr.CRD", '
        '"crop": "C:/in/M.CRO", "soil": "C:/in/s.SOI", "at": "2026-11-27 10:00:09"}\n'
        '{"year": 1981, "mm',  # บรรทัดเขียนค้าง
        encoding="utf-8",
    )
    s = Settings(input_dir=str(tmp_path), output_dir=str(out))
    status = runner.files_status(s)
    assert status["current"] == [] and status["recent_total"] == 2
    assert [e["date"] for e in status["recent"]] == ["15/04", "01/04"]  # ใหม่สุดก่อน
    assert status["recent"][0]["climate"].endswith("c_1981apr.PED")
