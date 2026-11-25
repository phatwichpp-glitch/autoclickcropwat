"""
cropwat_options.py
==================
v1.2.0 — ตั้งค่าหน้า "CROPWAT options" จากโปรแกรมเราโดยเขียน Cropwat.ini ตรงๆ

CropWat 8.0 เก็บค่าในหน้า Settings → Options ไว้ที่ %ProgramData%\\CROPWAT\\Cropwat.ini
(ทั้งเครื่อง) และอ่านไฟล์นี้ทุกครั้งที่เปิดโปรแกรม — รวมถึงตัวที่โปรแกรมเราเปิดเองบน
เดสก์ท็อปซ่อน การเขียนไฟล์นี้จึงเท่ากับกด "Save as default" ในหน้า Options โดยไม่ต้อง
ไปคลิกหน้าจอ CropWat (ซึ่งเสี่ยงค้าง)

ตาราง SECTIONS ข้างล่าง "ไม่ได้เดา" — จับคู่จาก CropWat 8.0 จริง (2026-11-25): เปิด CropWat
ตัวทดสอบบนเดสก์ท็อปซ่อน เลือกทุกตัวเลือก/ใส่ค่าทดสอบไม่ซ้ำกันทุกช่อง กด Save as default
แล้วดูว่าคีย์ไหนใน ini เปลี่ยน (สำรองและคืนไฟล์เดิมทุกไบต์หลังทดสอบ):
  - รหัสตัวเลือก = ลำดับใน dropdown + 1 (เช่น timopt=8 = "No irrigation (rainfed)")
  - คีย์ที่ไม่ได้ถูกบันทึกโดย Save as default (เช่นวิธี ETo "calculated from temperature
    data") ไม่อยู่ในตารางนี้ — แก้จากโปรแกรมนี้ไม่ได้
ตัวเลือก "user defined" (ตาราง interval/ความลึกที่ผู้ใช้กำหนดเอง) เก็บตารางไว้ที่อื่น
เลือกได้แต่ต้องไปกรอกตารางใน CropWat เอง

การเขียน: แก้ทีละบรรทัดในไฟล์เดิม (คงลำดับ/คีย์อื่น/บรรทัดที่ไม่รู้จักไว้ทั้งหมด) +
สำรองไฟล์ก่อนเขียนทุกครั้ง (Cropwat.ini.autorunner-bak = ก่อนบันทึกล่าสุด,
Cropwat.ini.autorunner-original = ก่อนโปรแกรมนี้แก้ครั้งแรก)
"""

from __future__ import annotations

import configparser
import os
import shutil
import subprocess
import time
from pathlib import Path
from typing import Any, Optional



def ini_path() -> Path:
    base = os.environ.get("PROGRAMDATA") or r"C:\ProgramData"
    return Path(base) / "CROPWAT" / "Cropwat.ini"


def _c(code: int, label: str) -> list:
    return [code, label]


_STAGES = [("1", "Initial stage"), ("2", "Development stage"), ("3", "Mid season"), ("4", "Late season")]

NONRICE_TIMING = [
    _c(1, "Irrigate at user defined intervals"),
    _c(2, "Irrigate at critical depletion"),
    _c(3, "Irrigate below or above critical depletion"),
    _c(4, "Irrigate at fixed interval per stage"),
    _c(5, "Irrigate at fixed depletion"),
    _c(6, "Irrigate at given ETcrop reduction per stage"),
    _c(7, "Irrigate at given yield reduction"),
    _c(8, "No irrigation (rainfed)"),
]
NONRICE_APPL = [
    _c(1, "User defined application depth"),
    _c(2, "Refill soil to field capacity"),
    _c(3, "Refill soil below / above field capacity"),
    _c(4, "Fixed application depth"),
]
RICE_TIMING = [
    _c(1, "Irrigate at user defined intervals"),
    _c(2, "Irrigate at fixed waterdepth"),
    _c(3, "Irrigate at fixed % desaturation"),
    _c(4, "Irrigate at fixed % of critical depletion"),
    _c(5, "Irrigate at fixed WD/desaturation/critical depletion per stage"),
    _c(6, "Irrigate at fixed interval per stage"),
    _c(7, "Irrigate at given ETcrop reduction"),
    _c(8, "Irrigate at given yield reduction"),
    _c(9, "Irrigate continuously"),
    _c(10, "No irrigation (rainfed)"),
]
RICE_APPL = [
    _c(1, "User defined application depth"),
    _c(2, "Refill to fixed waterdepth"),
    _c(3, "Refill to or below saturation"),
    _c(4, "Refill to or below field capacity"),
    _c(5, "Refill to fixed WD/saturation/FC"),
    _c(6, "Fixed application depth"),
]


def _f(key, label, kind="int", unit="", when=None, choices=None, minv=None, maxv=None, note=""):
    d = {"key": key, "label": label, "type": kind, "unit": unit}
    if when:
        d["when"] = when
    if choices:
        d["choices"] = choices
    if minv is not None:
        d["min"] = minv
    if maxv is not None:
        d["max"] = maxv
    if note:
        d["note"] = note
    return d


def _per_stage(prefix, unit, when, suffix="", minv=0, maxv=1000):
    return [_f(f"{prefix}{n}{suffix}", label, "int", unit, when, minv=minv, maxv=maxv) for n, label in _STAGES]


SECTIONS: list[dict] = [
    {
        "id": "rain", "ini": "Rain options", "title": "Effective rainfall",
        "desc": "วิธีคำนวณฝนใช้การได้ (มีผลกับผลคำนวณโดยตรง)",
        "fields": [
            _f("ermethod", "วิธีคำนวณ", "choice", choices=[
                _c(1, "Fixed percentage"), _c(2, "Dependable rain (FAO/AGLW formula)"),
                _c(3, "Empirical formula"), _c(4, "USDA soil conservation service"),
                _c(5, "Rainfall not considered (effective rainfall = 0)"),
            ]),
            _f("erfixperc", "เปอร์เซ็นต์คงที่", "int", "%", {"ermethod": [1]}, minv=0, maxv=100),
            _f("raina", "Peff = a × P + b (เมื่อ P ≤ x) — a", "float", "", {"ermethod": [3]}),
            _f("rainb", "— b", "int", "mm", {"ermethod": [3]}, minv=-1000, maxv=1000),
            _f("rainc", "Peff = c × P + d (เมื่อ P > x) — c", "float", "", {"ermethod": [3]}),
            _f("raind", "— d", "int", "mm", {"ermethod": [3]}, minv=-1000, maxv=1000),
            _f("rainx", "ค่า x (จุดแบ่ง)", "int", "mm", {"ermethod": [3]}, minv=0, maxv=1000),
        ],
    },
    {
        "id": "crop", "ini": "Crop scheduling options", "title": "Non-rice crop scheduling",
        "desc": "การให้น้ำของพืชที่ไม่ใช่ข้าว (เช่น ข้าวโพด)",
        "fields": [
            _f("timopt", "Irrigation timing", "choice", choices=NONRICE_TIMING,
               note="ตัวเลือก user defined ต้องกรอกตารางเองในโปรแกรม CropWat"),
            _f("deplram", "Irrigation at … % of critical depletion", "int", "%", {"timopt": [3]}, minv=0, maxv=1000),
            *_per_stage("timfixint", "days", {"timopt": [4]}, minv=1, maxv=365),
            _f("timdepl", "Irrigation at … mm depletion", "int", "mm", {"timopt": [5]}, minv=0, maxv=1000),
            *_per_stage("useretred", "%", {"timopt": [6]}, maxv=100),
            _f("yielddefi", "Irrigation at … % yield reduction", "int", "%", {"timopt": [7]}, minv=0, maxv=100),
            _f("applopt", "Irrigation application", "choice", "", {"timopt": [1, 2, 3, 4, 5, 6, 7]}, choices=NONRICE_APPL,
               note="ไม่ใช้เมื่อเลือก No irrigation (rainfed)"),
            _f("fieldcapcorr", "Refill soil moisture content to … % of field capacity", "int", "%",
               {"timopt": [1, 2, 3, 4, 5, 6, 7], "applopt": [3]}, minv=0, maxv=1000),
            _f("fixirridept", "Fixed irrigation depth", "int", "mm",
               {"timopt": [1, 2, 3, 4, 5, 6, 7], "applopt": [4]}, minv=0, maxv=1000),
            _f("appleff", "Irrigation efficiency", "int", "%", minv=1, maxv=100),
        ],
    },
    {
        "id": "rice", "ini": "Rice scheduling options", "title": "Rice scheduling",
        "desc": "การให้น้ำของข้าว",
        "fields": [
            _f("ricetimopt", "Irrigation timing", "choice", choices=RICE_TIMING,
               note="ตัวเลือก user defined ต้องกรอกตารางเองในโปรแกรม CropWat"),
            _f("ricetwd", "Irrigation at … mm waterdepth", "int", "mm", {"ricetimopt": [2]}, minv=0, maxv=1000),
            _f("ricetdesat", "Irrigation at … % desaturation", "int", "%", {"ricetimopt": [3]}, minv=0, maxv=100),
            _f("ricetdepl", "Irrigation at … % of critical depletion", "int", "%", {"ricetimopt": [4]}, minv=0, maxv=1000),
            *[x for n, label in _STAGES for x in (
                _f(f"ricetfixedany{n}1", label, "int", "", {"ricetimopt": [5]}, minv=0, maxv=1000),
                _f(f"ricetfixedany{n}2", f"{label} — หน่วย", "choice", "", {"ricetimopt": [5]},
                   choices=[_c(1, "mm WD"), _c(2, "% desat."), _c(3, "% crit.depl.")]),
            )],
            *_per_stage("ricetfixedint", "days", {"ricetimopt": [6]}, minv=1, maxv=365),
            *_per_stage("ricetetred", "%", {"ricetimopt": [7]}, maxv=100),
            _f("ricetyieldred", "Irrigation at … % yield reduction", "int", "%", {"ricetimopt": [8]}, minv=0, maxv=100),
            _f("ricecontinuous", "Continuous supply of", "int", "mm/day", {"ricetimopt": [9]}, minv=0, maxv=1000),
            _f("riceapplopt", "Irrigation application", "choice", "", {"ricetimopt": [1, 2, 3, 4, 5, 6, 7, 8]},
               choices=RICE_APPL, note="ไม่ใช้เมื่อเลือก Irrigate continuously หรือ No irrigation"),
            _f("riceawd", "Refill waterdepth to", "int", "mm",
               {"ricetimopt": [1, 2, 3, 4, 5, 6, 7, 8], "riceapplopt": [2]}, minv=0, maxv=1000),
            _f("riceasat", "Refill to … % saturation", "int", "%",
               {"ricetimopt": [1, 2, 3, 4, 5, 6, 7, 8], "riceapplopt": [3]}, minv=0, maxv=1000),
            _f("riceafc", "Refill to … % of field capacity", "int", "%",
               {"ricetimopt": [1, 2, 3, 4, 5, 6, 7, 8], "riceapplopt": [4]}, minv=0, maxv=1000),
            *[x for n, label in _STAGES for x in (
                _f(f"riceafixedany{n}1", label, "int", "",
                   {"ricetimopt": [1, 2, 3, 4, 5, 6, 7, 8], "riceapplopt": [5]}, minv=0, maxv=1000),
                _f(f"riceafixedany{n}2", f"{label} — หน่วย", "choice", "",
                   {"ricetimopt": [1, 2, 3, 4, 5, 6, 7, 8], "riceapplopt": [5]},
                   choices=[_c(1, "mm WD"), _c(2, "% sat."), _c(3, "% of FC")]),
            )],
            _f("riceafixeddepth", "Fixed application depth", "int", "mm",
               {"ricetimopt": [1, 2, 3, 4, 5, 6, 7, 8], "riceapplopt": [6]}, minv=0, maxv=1000),
            _f("riceappleff", "Irrigation efficiency", "int", "%", minv=1, maxv=100),
        ],
    },
    {
        "id": "landprep", "ini": "Land preparation options", "title": "Land preparation (rice)",
        "desc": "การเตรียมแปลงนาข้าว (ก่อนทำเทือก / ทำเทือก)",
        "fields": [
            _f("maxpercopt", "Maximum percolation rate after puddling", "choice",
               choices=[_c(1, "FAO formula (max. percolation of non-puddled soil ^ 0.33)"), _c(2, "User defined")]),
            _f("userdefpuddlepercrate", "User defined", "float", "mm/day", {"maxpercopt": [2]}, minv=0, maxv=1000),
            _f("changeinpercopt", "Daily decrease in maximum percolation rate during puddling", "choice",
               choices=[_c(1, "FAO formula"), _c(2, "Linear decrease")]),
            _f("soakbelowpuddledepth", "Soaking day 1: fill soil to saturation up to puddle depth +", "int", "cm", minv=0, maxv=1000),
            _f("prepuddlingtimopt", "Pre-puddling — irrigation timing", "choice",
               choices=[_c(1, "Irrigate at fixed % desaturation"), _c(2, "Irrigate at fixed % depletion of FC")]),
            _f("prepuddletdesat", "Irrigate at … % desaturation", "int", "%", {"prepuddlingtimopt": [1]}, minv=0, maxv=100),
            _f("prepuddletdepl", "Irrigate at … % depletion of field capacity", "int", "%", {"prepuddlingtimopt": [2]}, minv=0, maxv=100),
            _f("prepuddlingapplopt", "Pre-puddling — irrigation application", "choice",
               choices=[_c(1, "Fixed application depth"), _c(2, "Refill to fixed % saturation")]),
            _f("prepuddleafixeddepth", "Fixed application depth", "int", "mm", {"prepuddlingapplopt": [1]}, minv=0, maxv=1000),
            _f("prepuddleasat", "Refill soil moisture content to … % saturation", "int", "%", {"prepuddlingapplopt": [2]}, minv=0, maxv=1000),
            _f("puddlingtimopt", "Puddling — irrigation timing", "choice",
               choices=[_c(1, "Irrigate at fixed mm waterdepth"), _c(2, "Irrigate at fixed % desaturation")]),
            _f("puddletwd", "Irrigation at … mm waterdepth", "int", "mm", {"puddlingtimopt": [1]}, minv=0, maxv=1000),
            _f("puddletdesat", "Irrigation at … % desaturation", "int", "%", {"puddlingtimopt": [2]}, minv=0, maxv=100),
            _f("puddlingapplopt", "Puddling — irrigation application", "choice",
               choices=[_c(1, "Refill to fixed water depth"), _c(2, "Refill to fixed % saturation"), _c(3, "Fixed application depth")]),
            _f("puddleawd", "Refill waterdepth to", "int", "mm", {"puddlingapplopt": [1]}, minv=0, maxv=1000),
            _f("puddleasat", "Refill soil moisture content to … % saturation", "int", "%", {"puddlingapplopt": [2]}, minv=0, maxv=1000),
            _f("puddleafixeddepth", "Fixed application depth", "int", "mm", {"puddlingapplopt": [3]}, minv=0, maxv=1000),
        ],
    },
    {
        "id": "climate", "ini": "Climate/ETo options", "title": "Climate / ETo",
        "desc": "หน่วยและรูปแบบข้อมูลภูมิอากาศ — CropWat ระบุว่ามีผลกับข้อมูลที่สร้างใหม่ ไม่ใช่ไฟล์ที่มีอยู่แล้ว",
        "fields": [
            _f("newtempnotaverage", "Temperature", "choice",
               choices=[_c(1, "Minimum / maximum temperatures"), _c(0, "Average temperatures")]),
            _f("humidityasperc", "Humidity", "choice",
               choices=[_c(1, "Relative humidity in %"), _c(0, "Vapour pressure in kPa")]),
            _f("windaskmday", "Windspeed", "choice", choices=[_c(1, "Kilometers per day"), _c(0, "Meters per sec")]),
            _f("sunshineunit", "Sunshine", "choice",
               choices=[_c(1, "Hours sunshine"), _c(2, "% of daylength"), _c(3, "Fraction of daylength")]),
            _f("etoasmmday", "ETo", "choice", choices=[_c(1, "mm per day"), _c(0, "mm per period")]),
        ],
    },
]

FIELDS: dict[str, tuple[dict, dict]] = {f["key"]: (s, f) for s in SECTIONS for f in s["fields"]}


class OptionsError(Exception):
    """ข้อผิดพลาดที่อธิบายให้ผู้ใช้เข้าใจได้"""


def _read_raw(path: Path) -> dict[str, dict[str, str]]:
    cp = configparser.RawConfigParser(strict=False)
    cp.optionxform = str
    cp.read(path, encoding="latin-1")
    return {s: dict(cp.items(s)) for s in cp.sections()}


def _parse_value(field: dict, raw: Optional[str]) -> Any:
    if raw is None:
        return None
    try:
        if field["type"] == "float":
            return float(raw)
        return int(float(raw))
    except ValueError:
        return None


def read_options() -> dict:
    """ค่าปัจจุบันทั้งหมด + ตาราง (schema) สำหรับหน้าเว็บ"""
    path = ini_path()
    exists = path.is_file()
    raw = _read_raw(path) if exists else {}
    values = {}
    for key, (sec, field) in FIELDS.items():
        values[key] = _parse_value(field, raw.get(sec["ini"], {}).get(key))
    return {
        "path": str(path),
        "exists": exists,
        "values": values,
        "sections": SECTIONS,
        "has_backup": _backup_path(path).is_file(),
        "cropwat_running": cropwat_running(),
    }


def _format(field: dict, value: Any) -> str:
    if field["type"] == "float":
        v = float(value)
        return str(int(v)) if v.is_integer() else repr(round(v, 4))
    return str(int(value))


def validate(changes: dict) -> dict[str, str]:
    """ตรวจและแปลงค่าเป็นข้อความที่จะเขียน — raise OptionsError ถ้าผิด"""
    out: dict[str, str] = {}
    for key, value in changes.items():
        if key not in FIELDS:
            raise OptionsError(f"ไม่รู้จักค่า {key}")
        _sec, field = FIELDS[key]
        if field["type"] == "choice":
            codes = [c[0] for c in field["choices"]]
            try:
                code = int(value)
            except (TypeError, ValueError):
                raise OptionsError(f"{field['label']}: ค่าไม่ถูกต้อง") from None
            if code not in codes:
                raise OptionsError(f"{field['label']}: ไม่มีตัวเลือกนี้")
            out[key] = str(code)
            continue
        try:
            num = float(value)
        except (TypeError, ValueError):
            raise OptionsError(f"{field['label']}: ต้องเป็นตัวเลข") from None
        if field["type"] == "int" and not num.is_integer():
            raise OptionsError(f"{field['label']}: ต้องเป็นจำนวนเต็ม")
        if "min" in field and num < field["min"]:
            raise OptionsError(f"{field['label']}: ต้องไม่น้อยกว่า {field['min']}")
        if "max" in field and num > field["max"]:
            raise OptionsError(f"{field['label']}: ต้องไม่เกิน {field['max']}")
        out[key] = _format(field, num)
    return out


def _backup_path(path: Path) -> Path:
    return path.with_name(path.name + ".autorunner-bak")


def _original_path(path: Path) -> Path:
    return path.with_name(path.name + ".autorunner-original")


def _apply_lines(text: str, updates: dict[str, dict[str, str]]) -> str:
    """แก้ค่า key=value ทีละบรรทัดในส่วน [section] ที่ระบุ — บรรทัดอื่นคงเดิมทุกตัวอักษร
    คีย์ที่ยังไม่มีจะต่อท้ายส่วนนั้น ส่วนที่ยังไม่มีจะต่อท้ายไฟล์"""
    nl = "\r\n" if "\r\n" in text or not text else "\n"
    lines = text.splitlines()
    pending = {s: dict(kv) for s, kv in updates.items()}
    out: list[str] = []
    current: Optional[str] = None

    def flush(section):
        if section and pending.get(section):
            # ต่อท้ายส่วน (ก่อนบรรทัดว่างท้ายส่วน ถ้ามี)
            insert_at = len(out)
            while insert_at > 0 and out[insert_at - 1].strip() == "":
                insert_at -= 1
            for k, v in pending.pop(section).items():
                out.insert(insert_at, f"{k}={v}")
                insert_at += 1

    for line in lines:
        s = line.strip()
        if s.startswith("[") and s.endswith("]"):
            flush(current)
            current = s[1:-1]
            out.append(line)
            continue
        if current in pending and "=" in line and not s.startswith((";", "#")):
            k = line.split("=", 1)[0].strip()
            if k in pending[current]:
                out.append(f"{k}={pending[current].pop(k)}")
                continue
        out.append(line)
    flush(current)
    for section, kv in pending.items():
        if kv:
            out.append(f"[{section}]")
            out.extend(f"{k}={v}" for k, v in kv.items())
    return nl.join(out) + nl


def _write(path: Path, updates: dict[str, dict[str, str]]) -> None:
    if not path.is_file():
        raise OptionsError(
            f"ไม่พบไฟล์ค่าตั้งของ CropWat ({path}) — เปิดโปรแกรม CropWat 8.0 สักครั้งก่อน แล้วลองใหม่"
        )
    text = path.read_bytes().decode("latin-1")
    new_text = _apply_lines(text, updates)
    if new_text == text:
        return
    try:
        if not _original_path(path).exists():
            shutil.copy2(path, _original_path(path))
        shutil.copy2(path, _backup_path(path))
        path.write_bytes(new_text.encode("latin-1"))
    except PermissionError as exc:
        raise OptionsError(
            f"เขียนไฟล์ {path} ไม่ได้ (ไม่มีสิทธิ์) — ลองเปิดโปรแกรมนี้แบบ Run as administrator "
            "หรือให้ผู้ดูแลเครื่องตั้งสิทธิ์โฟลเดอร์ C:\\ProgramData\\CROPWAT ให้แก้ไขได้"
        ) from exc


def save_options(changes: dict) -> dict:
    checked = validate(changes)
    updates: dict[str, dict[str, str]] = {}
    for key, value in checked.items():
        sec, _field = FIELDS[key]
        updates.setdefault(sec["ini"], {})[key] = value
    _write(ini_path(), updates)
    return read_options()


def restore_backup() -> dict:
    path = ini_path()
    bak = _backup_path(path)
    if not bak.is_file():
        raise OptionsError("ยังไม่มีค่าก่อนบันทึกให้ย้อนกลับ")
    try:
        current = path.read_bytes()
        shutil.copy2(bak, path)
        bak.write_bytes(current)  # สลับกัน — กดซ้ำได้ (ย้อนกลับการย้อนกลับ)
    except PermissionError as exc:
        raise OptionsError(f"เขียนไฟล์ {path} ไม่ได้ (ไม่มีสิทธิ์)") from exc
    return read_options()


def export_profile(folder: str) -> str:
    path = ini_path()
    if not path.is_file():
        raise OptionsError("ไม่พบไฟล์ค่าตั้งของ CropWat ในเครื่องนี้")
    target_dir = Path(folder)
    if not target_dir.is_dir():
        raise OptionsError(f"ไม่พบโฟลเดอร์: {folder}")
    target = target_dir / f"CropWat_Options_{time.strftime('%Y%m%d-%H%M')}.ini"
    shutil.copy2(path, target)
    return str(target)


def import_profile(file_path: str) -> dict:
    """นำเข้าเฉพาะค่าที่รู้จักในตาราง (ไม่นำ [File Locations] ที่เป็น path ของอีกเครื่องมา)"""
    src = Path(file_path)
    if not src.is_file():
        raise OptionsError(f"ไม่พบไฟล์: {file_path}")
    try:
        raw = _read_raw(src)
    except configparser.Error as exc:
        raise OptionsError(f"อ่านไฟล์โปรไฟล์ไม่ได้: {exc}") from exc
    changes = {}
    for key, (sec, _field) in FIELDS.items():
        v = raw.get(sec["ini"], {}).get(key)
        if v is not None:
            changes[key] = v
    if not changes:
        raise OptionsError("ไฟล์นี้ไม่ใช่โปรไฟล์ค่าตั้ง CropWat (ไม่พบค่า Options ที่รู้จัก)")
    return save_options(changes)


def cropwat_running() -> bool:
    """มี CropWat เปิดอยู่ไหม (ตัวที่เปิดอยู่จะยังใช้ค่าเดิมจนกว่าจะปิดแล้วเปิดใหม่)"""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq cropwat.exe", "/NH"],
            capture_output=True, text=True, timeout=5, creationflags=0x08000000,
        ).stdout
        return "cropwat.exe" in out.lower()
    except Exception:  # noqa: BLE001
        return False


def summary(path: Optional[Path] = None) -> list[tuple[str, str]]:
    """ค่าหลักที่มีผลกับผลคำนวณ เป็นข้อความอ่านง่าย — ใช้ใน Excel sheet "หมายเหตุ" """
    p = path or ini_path()
    if not p.is_file():
        return []
    raw = _read_raw(p)

    def label(key):
        sec, field = FIELDS[key]
        v = _parse_value(field, raw.get(sec["ini"], {}).get(key))
        if v is None:
            return "-"
        if field["type"] == "choice":
            return next((c[1] for c in field["choices"] if c[0] == v), str(v))
        return f"{_format(field, v)} {field['unit']}".strip()

    rows = [
        ("CropWat: Effective rainfall", label("ermethod")),
        ("CropWat: Non-rice timing", label("timopt")),
        ("CropWat: Non-rice application", label("applopt")),
        ("CropWat: Non-rice irrigation efficiency", label("appleff")),
        ("CropWat: Rice timing", label("ricetimopt")),
        ("CropWat: Rice application", label("riceapplopt")),
    ]
    rm = _parse_value(FIELDS["ermethod"][1], raw.get("Rain options", {}).get("ermethod"))
    if rm == 1:
        rows.insert(1, ("CropWat: Fixed percentage", label("erfixperc")))
    elif rm == 3:
        rows.insert(1, ("CropWat: Empirical formula",
                        f"a={label('raina')}, b={label('rainb')}, c={label('rainc')}, d={label('raind')}, x={label('rainx')}"))
    return rows
