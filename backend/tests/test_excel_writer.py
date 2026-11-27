"""
tests/test_excel_writer.py
===========================
v0.12.0 — ล็อกโครงสร้างไฟล์ Result.xlsx: sheet Result / สรุป / หมายเหตุ
ใช้ไฟล์ .txt สมมติจาก test_txt_parser (ตัวเลขสมมติ ไม่ใช่ข้อมูลวิจัยจริง)
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

from file_engine.excel_writer import _AUTO_MARK, RunNotes, build_result_sheet
from tests.test_txt_parser import SAMPLE_TXT


def _write_txt(txt_dir: Path, year: int, mmdd: str, yield_red: float, planting: str | None = None) -> None:
    """planting = วันปลูกในเนื้อไฟล์ "dd/mm" (ค่าเริ่มต้น = ตรงกับชื่อไฟล์)"""
    text = SAMPLE_TXT.replace("Yield red.:   12.8 %", f"Yield red.:   {yield_red} %")
    text = text.replace("Planting date: 01/04", f"Planting date: {planting or f'{mmdd[2:]}/{mmdd[:2]}'}")
    (txt_dir / f"{year}_{mmdd}.txt").write_text(text, encoding="cp1252")


def test_file_with_other_planting_date_is_not_used(tmp_path):
    """v1.3.2 — ไฟล์ที่ชื่อเป็น 15/05 แต่เนื้อในเป็นวันปลูก 01/04 (ตั้งวันปลูกไม่ติด) ห้ามถูก
    ลงคอลัมน์ 15/05 และต้องถูกบันทึกไว้ใน sheet หมายเหตุ"""
    txt_dir = _fixture(tmp_path)
    _write_txt(txt_dir, 1981, "0515", 99.0, planting="01/04")
    out = tmp_path / "Result.xlsx"
    build_result_sheet(txt_dir, out)
    wb = load_workbook(out)
    headers = [c.value for c in wb["Result"][1]]
    assert "15/05" not in headers
    notes_text = " ".join(str(c.value) for row in wb["หมายเหตุ"].iter_rows() for c in row if c.value)
    assert "1981_0515.txt" in notes_text


def _fixture(tmp_path: Path) -> Path:
    txt_dir = tmp_path / "txt"
    txt_dir.mkdir()
    # 2 ปี × 2 วันปลูก — 15/04 ให้ yield reduction ต่ำกว่า (ต้องเป็นวันที่ดีที่สุด)
    _write_txt(txt_dir, 1981, "0401", 20.0)
    _write_txt(txt_dir, 1981, "0415", 5.0)
    _write_txt(txt_dir, 1982, "0401", 30.0)
    _write_txt(txt_dir, 1982, "0415", 7.0)
    return txt_dir


def test_creates_three_sheets_result_first(tmp_path):
    out = tmp_path / "Result.xlsx"
    assert build_result_sheet(_fixture(tmp_path), out) == 2
    wb = load_workbook(out)
    assert wb.sheetnames[:3] == ["Result", "สรุป", "หมายเหตุ"]


def test_result_sheet_values(tmp_path):
    out = tmp_path / "Result.xlsx"
    build_result_sheet(_fixture(tmp_path), out)
    ws = load_workbook(out)["Result"]
    assert ws["C1"].value == "01/04" and ws["D1"].value == "15/04"
    assert ws["A2"].value == 1981 and ws["C2"].value == 20.0 and ws["D2"].value == 5.0
    assert ws["A13"].value == 1982 and ws["C13"].value == 30.0  # บล็อกละ 11 แถว
    assert ws.conditional_formatting  # แถบสีบนแถว yield reduction


def test_summary_stats_and_best_date(tmp_path):
    out = tmp_path / "Result.xlsx"
    build_result_sheet(_fixture(tmp_path), out)
    ws = load_workbook(out)["สรุป"]
    rows = {ws.cell(r, 1).value: [ws.cell(r, c).value for c in range(2, 6)] for r in (4, 5)}
    assert rows["01/04"] == [2, 25.0, 20.0, 30.0]  # ปี, เฉลี่ย, ต่ำสุด, สูงสุด
    assert rows["15/04"] == [2, 6.0, 5.0, 7.0]
    assert ws.cell(5, 1).font.bold and not ws.cell(4, 1).font.bold  # ไฮไลต์ 15/04


def test_notes_lists_missing_dates_errors_and_shots(tmp_path):
    out = tmp_path / "Result.xlsx"
    notes = RunNotes(
        planned={1981: ["0401", "0415", "0501"], 1982: ["0401", "0415"], 1983: ["0401"]},
        year_errors={1983: "Rain data are not OK"},
        error_shots={1983: ["1983_0401_101010.png"]},
    )
    build_result_sheet(_fixture(tmp_path), out, notes)
    ws = load_workbook(out)["หมายเหตุ"]
    by_year = {
        row[0]: row[1:] for row in ws.iter_rows(values_only=True) if isinstance(row[0], int)
    }
    assert by_year[1981][:2] == ("2/3", "01/05")
    assert by_year[1982][1] == "-"
    assert by_year[1983][0] == "0/1"
    assert "Rain data are not OK" in by_year[1983][2]
    assert "1983_0401_101010.png" in by_year[1983][2]
    values = [c for row in ws.iter_rows(values_only=True) for c in row if c]
    assert "111111" in values and "test loam" in values  # metadata จากไฟล์ .txt


def test_unparseable_txt_is_reported_in_notes(tmp_path):
    txt_dir = _fixture(tmp_path)
    (txt_dir / "1982_0501.txt").write_text("garbage", encoding="cp1252")
    out = tmp_path / "Result.xlsx"
    build_result_sheet(txt_dir, out)
    ws = load_workbook(out)["หมายเหตุ"]
    remarks = [row[3] for row in ws.iter_rows(values_only=True) if row[0] == 1982]
    assert "1982_0501.txt" in remarks[0]


def test_user_written_sheets_are_never_overwritten(tmp_path):
    out = tmp_path / "Result.xlsx"
    wb = Workbook()
    wb.active.title = "หมายเหตุ"
    wb["หมายเหตุ"]["A1"] = "บันทึกของฉันเอง"
    wb.create_sheet("กราฟของฉัน")["A1"] = 42
    wb.save(out)

    txt_dir = _fixture(tmp_path)
    build_result_sheet(txt_dir, out)
    build_result_sheet(txt_dir, out)  # สร้างซ้ำต้องไม่งอกแผ่นใหม่เรื่อยๆ
    wb = load_workbook(out)
    assert wb["หมายเหตุ"]["A1"].value == "บันทึกของฉันเอง"
    assert wb["กราฟของฉัน"]["A1"].value == 42
    assert _AUTO_MARK in wb["หมายเหตุ (auto)"]["A1"].value
    assert sorted(wb.sheetnames) == sorted(
        ["Result", "สรุป", "หมายเหตุ (auto)", "หมายเหตุ", "กราฟของฉัน", "ไฟล์ที่ใช้"]
    )


def test_files_used_sheet_lists_files_and_flags_fallback(tmp_path):
    """v1.3.2 — sheet "ไฟล์ที่ใช้": 1 แถวต่อวันปลูกที่มีผล บอกไฟล์ + ที่มา และไฮไลต์แถวที่
    ไฟล์ไม่ใช่ของเดือนที่ปลูก"""
    out = tmp_path / "Result.xlsx"
    notes = RunNotes(files_used={
        (1981, "0401"): {
            "climate": "D:/in/Clim_1/04_Apr/111111clim_1981apr.PED",
            "rain": "D:/in/Rain_2/rain_222222_1981apr.CRD",
            "crop": "D:/in/MAIZE.CRO", "soil": "D:/in/sand.SOI",
            "at": "2026-11-27 10:00:00", "source": "run",
        },
        (1981, "0415"): {
            "climate": "D:/in/Clim_1/111111clim_1980dec.PED",
            "rain": "D:/in/Rain_2/rain_222222_1981apr.CRD",
            "crop": "", "soil": "", "at": "", "source": "estimate",
        },
    })
    build_result_sheet(_fixture(tmp_path), out, notes)
    wb = load_workbook(out)
    assert wb.sheetnames[:4] == ["Result", "สรุป", "หมายเหตุ", "ไฟล์ที่ใช้"]
    ws = wb["ไฟล์ที่ใช้"]
    rows = {(r[0].value, r[1].value): r for r in ws.iter_rows(min_row=4) if r[0].value}
    assert len(rows) == 4  # 2 ปี × 2 วันปลูกที่มีผล

    r = rows[(1981, "01/04")]
    assert [c.value for c in r[2:10]] == [
        "111111clim_1981apr.PED", "04/1981", "rain_222222_1981apr.CRD", "04/1981",
        "MAIZE.CRO", "sand.SOI", "บันทึกตอนรัน", "2026-11-27 10:00:00",
    ]
    assert r[10].value == "-" and r[11].value.endswith("04_Apr")

    r = rows[(1981, "15/04")]
    assert r[8].value == "ประเมินย้อนหลัง"
    assert "12/1980" in r[10].value and r[0].fill.start_color.rgb.endswith("FCE4D6")

    assert rows[(1982, "01/04")][8].value == "ไม่มีบันทึก"
