"""
tests/test_cropwat_options.py
==============================
v1.2.0 — เขียน Cropwat.ini (ใช้ไฟล์จำลองใน tmp เท่านั้น ไม่แตะไฟล์จริงของเครื่อง)
"""

import pytest

import cropwat_options as co

SAMPLE = (
    "[File Locations]\r\nClimate/ETo Directory=D:\\data\r\n"
    "[Climate/ETo options]\r\nshowgeneralinfo=0\r\nnewtempnotaverage=1\r\nhumidityasperc=1\r\n"
    "[Rain options]\r\nnewraintype=3\r\nermethod=2\r\nerfixperc=80\r\nraina=0.5\r\nrainb=-5\r\n"
    "[Crop scheduling options]\r\ntimopt=8\r\napplopt=2\r\nappleff=70\r\n"
    "[Land preparation options]\r\nuserdefpuddlepercrate=3\r\n"
)


@pytest.fixture
def ini(tmp_path, monkeypatch):
    p = tmp_path / "Cropwat.ini"
    p.write_bytes(SAMPLE.encode("latin-1"))
    monkeypatch.setattr(co, "ini_path", lambda: p)
    monkeypatch.setattr(co, "cropwat_running", lambda: False)
    return p


def test_read_values(ini):
    r = co.read_options()
    assert r["values"]["ermethod"] == 2 and r["values"]["raina"] == 0.5 and r["values"]["rainb"] == -5
    assert r["values"]["ricetimopt"] is None  # ไม่มีในไฟล์


def test_save_changes_only_target_lines(ini):
    co.save_options({"ermethod": 3, "raina": 0.6, "timopt": 2, "appleff": 65})
    text = ini.read_bytes().decode("latin-1")
    assert "ermethod=3\r\n" in text and "raina=0.6\r\n" in text and "timopt=2\r\n" in text and "appleff=65\r\n" in text
    # บรรทัดอื่นคงเดิมทุกตัวอักษร (รวม CRLF, คีย์ที่ไม่รู้จัก, File Locations)
    for line in ("Climate/ETo Directory=D:\\data", "showgeneralinfo=0", "newraintype=3", "rainb=-5", "userdefpuddlepercrate=3"):
        assert line + "\r\n" in text
    assert text.count("\n") == text.count("\r\n")


def test_missing_key_and_section_are_appended(ini):
    co.save_options({"rainx": 40, "ricetimopt": 9, "ricecontinuous": 8})
    r = co.read_options()["values"]
    assert r["rainx"] == 40 and r["ricetimopt"] == 9 and r["ricecontinuous"] == 8
    text = ini.read_bytes().decode("latin-1")
    rain = text.index("[Rain options]"); crop = text.index("[Crop scheduling options]")
    assert rain < text.index("rainx=40") < crop  # ต่อท้ายในส่วนของมันเอง


def test_backup_and_restore(ini):
    original = ini.read_bytes()
    co.save_options({"ermethod": 5})
    assert ini.with_name("Cropwat.ini.autorunner-original").read_bytes() == original
    co.restore_backup()
    assert ini.read_bytes() == original
    co.restore_backup()  # ย้อนกลับการย้อนกลับ
    assert co.read_options()["values"]["ermethod"] == 5


@pytest.mark.parametrize("changes", [
    {"ermethod": 9}, {"appleff": 0}, {"appleff": 101}, {"appleff": 70.5}, {"erfixperc": "abc"}, {"bogus": 1},
])
def test_invalid_values_rejected_and_file_untouched(ini, changes):
    before = ini.read_bytes()
    with pytest.raises(co.OptionsError):
        co.save_options(changes)
    assert ini.read_bytes() == before


def test_float_formatting(ini):
    co.save_options({"userdefpuddlepercrate": 3.0})
    assert "userdefpuddlepercrate=3\r\n" in ini.read_bytes().decode("latin-1")
    co.save_options({"userdefpuddlepercrate": 2.5})
    assert "userdefpuddlepercrate=2.5\r\n" in ini.read_bytes().decode("latin-1")


def test_import_skips_file_locations(ini, tmp_path):
    prof = tmp_path / "profile.ini"
    prof.write_text("[File Locations]\nClimate/ETo Directory=E:\\other\n[Rain options]\nermethod=1\nerfixperc=70\n", encoding="latin-1")
    co.import_profile(str(prof))
    text = ini.read_bytes().decode("latin-1")
    assert "Climate/ETo Directory=D:\\data" in text and "ermethod=1" in text and "erfixperc=70" in text


def test_import_rejects_non_profile(ini, tmp_path):
    bad = tmp_path / "x.ini"
    bad.write_text("[Foo]\na=1\n", encoding="latin-1")
    with pytest.raises(co.OptionsError):
        co.import_profile(str(bad))


def test_summary_labels(ini):
    rows = dict(co.summary())
    assert rows["CropWat: Effective rainfall"] == "Dependable rain (FAO/AGLW formula)"
    assert rows["CropWat: Non-rice timing"] == "No irrigation (rainfed)"
    assert rows["CropWat: Non-rice irrigation efficiency"] == "70 %"


def test_schema_codes_are_unique_and_keys_unique():
    keys = [f["key"] for s in co.SECTIONS for f in s["fields"]]
    assert len(keys) == len(set(keys))
    for s in co.SECTIONS:
        for f in s["fields"]:
            if f["type"] == "choice":
                codes = [c[0] for c in f["choices"]]
                assert len(codes) == len(set(codes))
            assert isinstance(f["unit"], str), f["key"]  # เงื่อนไข when ต้องไม่หลุดไปอยู่ในช่องหน่วย
            for dep in f.get("when", {}):
                assert dep in co.FIELDS


def test_dependent_choices_have_conditions():
    for key in ("applopt", "riceapplopt", "ricetfixedany12", "riceafixedany12"):
        assert co.FIELDS[key][1].get("when"), key
