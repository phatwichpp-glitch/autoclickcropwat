"""
file_engine/discovery.py
=========================
v1.4.0 — ค้นหาไฟล์ข้อมูลต้นทางแบบ "ไม่ยึดชื่อโฟลเดอร์/รูปแบบชื่อไฟล์ตายตัว"

เดิม (file_engine/paths.py): โฟลเดอร์สถานีต้องชื่อขึ้นต้น Clim_/Rain_ และอยู่ชั้นแรกของ
โฟลเดอร์ต้นทาง, ชื่อไฟล์ต้องเป็น "<สถานี>clim_<ปี><เดือน>.ped" / "rain_<สถานี>_<ปี><เดือน>.crd"
เป๊ะ — ผิดจากนี้ = หาไม่เจอ โดยไม่บอกว่าไฟล์ไหนถูกข้ามเพราะอะไร

ตอนนี้: เดินทุกโฟลเดอร์ย่อยของโฟลเดอร์ต้นทาง เก็บทุกไฟล์ .PED / .CRD / .CRO / .SOI แล้วหา
"ชนิด / สถานี / ปี / เดือน" ของแต่ละไฟล์จากหลักฐานหลายทางประกอบกัน:

  1. เนื้อไฟล์ — ไฟล์ของ CropWat 8.0 ขึ้นต้นด้วยบรรทัดบอกชนิด ("CROPWAT 8.0 Climate data" /
     "Rain data" / "Crop data" / "Soil Data") และมีรหัสสถานี + ปีอยู่ในหัวไฟล์ (ยืนยันจากไฟล์
     จริงใน cropwat/: climate = สถานีบรรทัดที่ 4 ปีบรรทัดที่ 7, rain = สถานีบรรทัดที่ 2 ปีบรรทัดที่ 5)
     — "เดือน" ไม่มีในเนื้อไฟล์ (ทุกไฟล์เป็นข้อมูลเต็มปี เดือนในชื่อคือเดือนปลูกที่ไฟล์นี้เตรียมไว้ให้)
  2. ชื่อไฟล์ — ปี ค.ศ. หรือ พ.ศ. (แปลงให้), เดือนเป็นชื่ออังกฤษย่อ/เต็ม ชื่อไทยย่อ/เต็ม
     หรือตัวเลขติดกับปี (1981-04, 198104, 04_1981)
  3. ชื่อโฟลเดอร์ที่ครอบอยู่ (ใกล้สุดก่อน) — ใช้เมื่อชื่อไฟล์ไม่มีเดือน/ปี (เช่น โฟลเดอร์ "04_Apr")

หลักการ: "ไม่เดาเงียบๆ" — ไฟล์ที่หลักฐานไม่พอหรือขัดกันเองจะไม่ถูกใช้ และถูกรายงานพร้อมเหตุผล
(unreadable) ส่วนที่ใช้ได้แต่มีข้อสังเกต (ชื่อบอกปีหนึ่ง เนื้อไฟล์อีกปีหนึ่ง / มีไฟล์ซ้ำเนื้อหา
ไม่เหมือนกัน) รายงานเป็น warnings ให้ผู้ใช้เห็นในหน้าตั้งค่าก่อนเริ่มรัน
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from file_engine.paths import StationIndex

KIND_BY_EXT = {".ped": "climate", ".crd": "rain", ".cro": "crop", ".soi": "soil"}
KIND_LABEL = {"climate": "Climate/ETo", "rain": "Rain", "crop": "Crop", "soil": "Soil"}
_SIGNATURE_WORD = {"climate": "climate", "rain": "rain", "crop": "crop", "soil": "soil"}

_EN_FULL = ["january", "february", "march", "april", "may", "june", "july", "august",
            "september", "october", "november", "december"]
_EN_ABBR = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
_TH_FULL = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม",
            "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
_TH_ABBR = ["มค", "กพ", "มีค", "เมย", "พค", "มิย", "กค", "สค", "กย", "ตค", "พย", "ธค"]

_EN_WORDS = {w: i + 1 for i, w in enumerate(_EN_FULL)}
_EN_WORDS.update({w: i + 1 for i, w in enumerate(_EN_ABBR)})
_EN_WORDS["sept"] = 9
_EN_MONTH_RE = re.compile(
    r"(?<![a-z])(" + "|".join(sorted(_EN_WORDS, key=len, reverse=True)) + r")(?![a-z])"
)
_TH_WORDS = {w: i + 1 for i, w in enumerate(_TH_FULL)}
_TH_WORDS.update({w: i + 1 for i, w in enumerate(_TH_ABBR)})
_TH_MONTH_RE = re.compile(
    r"(?<![ก-๎])(" + "|".join(sorted(_TH_WORDS, key=len, reverse=True)) + r")(?![ก-๎])"
)

# ปี ค.ศ. 1800-2199 หรือ พ.ศ. 2400-2699 (แปลงเป็น ค.ศ.) — ต้องไม่ติดกับตัวเลขอื่น
# (รหัสสถานี 6 หลักอย่าง 450006 จึงไม่ถูกมองเป็นปี)
_YEAR_RE = re.compile(r"(?<!\d)(1[89]\d\d|2[0-6]\d\d)(?!\d)")
# ปี+เดือนเขียนติดกัน 6 หลัก (198104) — ต้องไม่ติดกับตัวเลขอื่นเช่นกัน
_YEARMONTH_RE = re.compile(r"(?<!\d)(1[89]\d\d|2[0-6]\d\d)(0[1-9]|1[0-2])(?!\d)")
_FOLDER_MONTH_RE = re.compile(r"^(?:เดือน|month|m)?[\s_\-.]*(0?[1-9]|1[0-2])(?:\D|$)")
_BACKUP_HINT_RE = re.compile(r"backup|bak|old|copy|สำรอง|เก่า|สำเนา", re.IGNORECASE)


def _to_ce(year: int) -> int:
    return year - 543 if year >= 2400 else year


def months_in_text(text: str) -> set[int]:
    """เดือนที่เขียนเป็น "ชื่อ" (อังกฤษ/ไทย ย่อ/เต็ม) ในข้อความ"""
    t = text.lower()
    found = {_EN_WORDS[m] for m in _EN_MONTH_RE.findall(t)}
    found |= {_TH_WORDS[m] for m in _TH_MONTH_RE.findall(t.replace(".", ""))}
    return found


@dataclass
class NameInfo:
    years: list[int] = field(default_factory=list)   # ปีที่เจอในชื่อไฟล์ (ค.ศ.) ไม่ซ้ำ
    months: list[int] = field(default_factory=list)  # เดือนที่เจอในชื่อไฟล์
    station: Optional[str] = None                    # เลขยาว ≥ 4 หลักที่ไม่ใช่ปี
    folder_year: Optional[int] = None
    folder_month: Optional[int] = None


def parse_name(path: Path, root: Optional[Path] = None, known_station: str = "") -> NameInfo:
    """อ่านปี/เดือน/สถานีจากชื่อไฟล์ + ชื่อโฟลเดอร์ที่ครอบ (ไม่เปิดไฟล์)"""
    info = NameInfo()
    stem = path.stem.lower()
    if known_station and known_station.lower() in stem:
        stem = stem.replace(known_station.lower(), "#" * len(known_station))

    months = months_in_text(stem)
    years: list[int] = []
    for m in _YEARMONTH_RE.finditer(stem):
        years.append(_to_ce(int(m.group(1))))
        months.add(int(m.group(2)))
    spans = []
    for m in _YEAR_RE.finditer(stem):
        years.append(_to_ce(int(m.group(1))))
        spans.append(m.span())
        if not months:
            after = re.match(r"[\s_\-.](0?[1-9]|1[0-2])(?!\d)", stem[m.end():])
            before = re.search(r"(?<!\d)(0?[1-9]|1[0-2])[\s_\-.]$", stem[: m.start()])
            for hit in (after, before):
                if hit:
                    months.add(int(hit.group(1)))
    info.years = sorted(set(years))
    info.months = sorted(months)

    digit_runs = [
        m.group(0) for m in re.finditer(r"\d{4,}", stem)
        if not any(m.start() >= a and m.end() <= b for a, b in spans)
        and not _YEARMONTH_RE.fullmatch(m.group(0))
    ]
    if digit_runs:
        info.station = max(digit_runs, key=len)

    # โฟลเดอร์ที่ครอบ: ใกล้สุดก่อน หยุดที่โฟลเดอร์ต้นทาง — เดือนที่เขียนเป็น "ชื่อ" ในโฟลเดอร์
    # ชั้นไหนก็ตามมาก่อนโฟลเดอร์ที่เป็นตัวเลขล้วนเสมอ: โครงสร้างจริงของผู้ใช้คือ
    # "01_Rain381017Jan\01\Rain_381017_1981.CRD" ซึ่งโฟลเดอร์ย่อย "01", "02", ... เป็นลำดับ
    # กลุ่มปี (1981-1990, 1991-2000, ...) ไม่ใช่เดือน ถ้าเชื่อเลขใกล้สุดก่อน ไฟล์ใน "...Jan\02"
    # จะถูกอ่านเป็นเดือน ก.พ. ผิดเงียบๆ — เลขล้วนจึงใช้เฉพาะเมื่อไม่มีชื่อเดือนในชั้นไหนเลย
    numeric_month: Optional[int] = None
    for parent in path.parents:
        if root is not None and (parent == root or root not in parent.parents):
            break
        name = parent.name.lower()
        if info.folder_month is None:
            named = months_in_text(name)
            if len(named) == 1:
                info.folder_month = next(iter(named))
            elif not named and numeric_month is None:
                m = _FOLDER_MONTH_RE.match(name)
                if m:
                    numeric_month = int(m.group(1))
        if info.folder_year is None:
            ys = {_to_ce(int(y)) for y in _YEAR_RE.findall(name)}
            if len(ys) == 1:  # "1981" ใช่ / "1981-1990" (ช่วงปี) ไม่ใช่
                info.folder_year = next(iter(ys))
        if root is None:
            break
    if info.folder_month is None:
        info.folder_month = numeric_month
    return info


@dataclass
class Sniffed:
    kind: Optional[str] = None      # ชนิดตามบรรทัดแรกของไฟล์ (None = ไม่ใช่ไฟล์ CropWat 8.0)
    station: str = ""
    year: Optional[int] = None


def sniff(path: Path) -> Sniffed:
    """อ่านหัวไฟล์ (ไม่กี่ร้อยไบต์แรก) หาชนิด/สถานี/ปีที่ CropWat เขียนไว้เอง"""
    out = Sniffed()
    try:
        with open(path, "rb") as f:
            head = f.read(600).decode("latin-1", errors="replace")
    except OSError:
        return out
    lines = head.splitlines()
    if not lines or not lines[0].lower().startswith("cropwat"):
        return out
    first = lines[0].lower()
    for kind, word in _SIGNATURE_WORD.items():
        if f"{word} data" in first:
            out.kind = kind
            break
    station_line, year_line = {"climate": (3, 6), "rain": (1, 4)}.get(out.kind, (None, None))
    if station_line is not None and len(lines) > station_line:
        out.station = lines[station_line].strip()
    if year_line is not None and len(lines) > year_line:
        m = re.fullmatch(r"\s*(\d{4})\s*", lines[year_line])
        if m and 1800 <= int(m.group(1)) <= 2699:
            out.year = _to_ce(int(m.group(1)))
    return out


@dataclass
class DataFile:
    path: Path
    kind: str
    station: str
    year: int
    month: int


@dataclass
class Discovery:
    root: Path
    # ชนิด -> สถานี -> (ปี, เดือน) -> ไฟล์
    stations: dict = field(default_factory=lambda: {"climate": {}, "rain": {}})
    crops: list = field(default_factory=list)
    soils: list = field(default_factory=list)
    unreadable: list = field(default_factory=list)  # [{"path", "kind", "reason"}]
    warnings: list = field(default_factory=list)    # ข้อความไทย
    duplicate_copies: int = 0                       # ไฟล์ซ้ำที่เนื้อหาเหมือนกันทุกไบต์ (ข้ามเงียบๆ ได้)

    def station_folder(self, kind: str, station: str) -> Path:
        """โฟลเดอร์ร่วมที่ครอบทุกไฟล์ของสถานีนี้"""
        files = [f.path for f in self.stations[kind][station].values()]
        try:
            return Path(os.path.commonpath([str(p.parent) for p in files]))
        except ValueError:
            return self.root

    def pick_station(self, kind: str, override: str = "") -> str:
        """เลือกสถานีที่จะใช้: ค่าที่ผู้ใช้ตั้ง (รหัสสถานี หรือโฟลเดอร์) ก่อน ไม่ได้ตั้ง = ต้องมี
        สถานีเดียว — หลายสถานีแล้วไม่ได้เลือก = error บอกตัวเลือก (ไม่เดาแทนผู้ใช้)"""
        label = KIND_LABEL[kind]
        found = self.stations[kind]
        override = (override or "").strip()
        if override:
            for station in found:
                if station.lower() == override.lower():
                    return station
            folder = Path(override)
            if folder.is_dir():
                inside = [
                    s for s, files in found.items()
                    if any(folder == f.path.parent or folder in f.path.parents for f in files.values())
                ]
                if len(inside) == 1:
                    return inside[0]
                if not inside:
                    raise FileNotFoundError(f"ไม่พบไฟล์ {label} ในโฟลเดอร์ที่ตั้งไว้: {folder}")
                raise ValueError(
                    f"โฟลเดอร์ที่ตั้งไว้มีข้อมูล {label} มากกว่า 1 สถานี ({', '.join(sorted(inside))}) "
                    "— ใส่รหัสสถานีที่ต้องการแทนชื่อโฟลเดอร์ในหน้าตั้งค่า"
                )
            raise FileNotFoundError(
                f"ไม่พบสถานี {label} ที่ตั้งไว้ ({override}) — สถานีที่พบ: {', '.join(sorted(found)) or 'ไม่มี'}"
            )
        if not found:
            ext = ".PED" if kind == "climate" else ".CRD"
            hint = ""
            skipped = sum(1 for u in self.unreadable if u["kind"] == kind)
            if skipped:
                hint = f" (มีไฟล์ {ext} {skipped} ไฟล์ที่อ่านปี/เดือนไม่ออก — ดูรายการในหน้าตั้งค่า กด \"ตรวจสอบ\")"
            raise FileNotFoundError(f"ไม่พบไฟล์ข้อมูล {label} ({ext}) ใน {self.root}{hint}")
        if len(found) > 1:
            raise ValueError(
                f"พบข้อมูล {label} มากกว่า 1 สถานี ({', '.join(sorted(found))}) — เลือกสถานีที่จะใช้ในหน้าตั้งค่า"
            )
        return next(iter(found))

    def index(self, kind: str, override: str = "") -> StationIndex:
        station = self.pick_station(kind, override)
        return StationIndex(
            files_by_year_month={key: f.path for key, f in self.stations[kind][station].items()}
        )


def _same_content(a: Path, b: Path) -> bool:
    try:
        return a.stat().st_size == b.stat().st_size and a.read_bytes() == b.read_bytes()
    except OSError:
        return False


def _prefer(new: Path, old: Path) -> bool:
    """ไฟล์ซ้ำ (สถานี/ปี/เดือนเดียวกัน เนื้อหาต่างกัน): เลือกตัวที่ไม่ได้อยู่ในโฟลเดอร์สำรอง
    ก่อน แล้วค่อยเลือกตัวที่แก้ไขล่าสุด"""
    new_bak, old_bak = bool(_BACKUP_HINT_RE.search(str(new.parent))), bool(_BACKUP_HINT_RE.search(str(old.parent)))
    if new_bak != old_bak:
        return old_bak
    try:
        return new.stat().st_mtime > old.stat().st_mtime
    except OSError:
        return False


def _rel(path: Path, root: Path) -> str:
    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


def discover(root: Path, exclude: tuple = ()) -> Discovery:
    """เดินทั้งโฟลเดอร์ต้นทาง (ข้ามโฟลเดอร์ใน exclude เช่น โฟลเดอร์ผลลัพธ์ที่อยู่ข้างใน)"""
    root = Path(root)
    found = Discovery(root=root)
    if not root.is_dir():
        return found
    skip = {os.path.normcase(os.path.abspath(str(e))) for e in exclude if e}

    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(
            d for d in dirnames
            if os.path.normcase(os.path.abspath(os.path.join(dirpath, d))) not in skip
        )
        for name in sorted(filenames):
            kind = KIND_BY_EXT.get(os.path.splitext(name)[1].lower())
            if kind is None:
                continue
            path = Path(dirpath) / name
            if kind == "crop":
                found.crops.append(path)
            elif kind == "soil":
                found.soils.append(path)
            else:
                _add_station_file(found, path, kind)
    return found


def _add_station_file(found: Discovery, path: Path, kind: str) -> None:
    root = found.root
    rel = _rel(path, root)

    def reject(reason: str) -> None:
        found.unreadable.append({"path": rel, "kind": kind, "reason": reason})

    head = sniff(path)
    if head.kind is not None and head.kind != kind:
        reject(
            f"เนื้อไฟล์เป็นข้อมูล {KIND_LABEL[head.kind]} แต่นามสกุลเป็น {path.suffix.upper()} "
            f"(ไฟล์ {KIND_LABEL[kind]} ต้องเป็น {'.PED' if kind == 'climate' else '.CRD'})"
        )
        return
    name = parse_name(path, root, known_station=head.station if head.station.isdigit() else "")

    # ---- ปี: ชื่อไฟล์ก่อน (ถ้ามีหลายปีในชื่อ ใช้ปีที่ตรงกับเนื้อไฟล์) → เนื้อไฟล์ → โฟลเดอร์
    if len(name.years) == 1:
        year = name.years[0]
    elif len(name.years) > 1:
        if head.year in name.years:
            year = head.year
        else:
            reject(f"ชื่อไฟล์มีปีมากกว่า 1 ค่า ({', '.join(map(str, name.years))}) ระบุไม่ได้ว่าเป็นข้อมูลปีไหน")
            return
    else:
        year = head.year or name.folder_year
    if year is None:
        reject("อ่านปีไม่ออก — ไม่มีปีในชื่อไฟล์ ชื่อโฟลเดอร์ หรือหัวไฟล์")
        return

    # ---- เดือน: ชื่อไฟล์ก่อน → โฟลเดอร์ (เนื้อไฟล์ไม่มีเดือน)
    if len(name.months) == 1:
        month = name.months[0]
    elif len(name.months) > 1:
        reject(f"ชื่อไฟล์มีเดือนมากกว่า 1 ค่า ({', '.join(map(str, name.months))}) ระบุไม่ได้ว่าเป็นเดือนไหน")
        return
    elif name.folder_month is not None:
        month = name.folder_month
    else:
        reject(
            "อ่านเดือนไม่ออก — ใส่เดือนในชื่อไฟล์หรือชื่อโฟลเดอร์ "
            "(เช่น Apr, April, เม.ย., เมษายน, 1981-04 หรือโฟลเดอร์ชื่อ 04)"
        )
        return

    # ---- สถานี: หัวไฟล์ → เลขในชื่อไฟล์ → โฟลเดอร์ชั้นแรกใต้โฟลเดอร์ต้นทาง
    station = head.station or name.station
    if not station:
        parts = Path(rel).parts
        station = parts[0] if len(parts) > 1 else "(ไม่ระบุสถานี)"

    if head.year is not None and name.years and head.year != year:
        found.warnings.append(
            f"{rel}: ชื่อไฟล์บอกปี {year} แต่หัวไฟล์เขียนปี {head.year} — ใช้ตามชื่อไฟล์"
        )
    if head.station and name.station and head.station != name.station and head.station.isdigit():
        found.warnings.append(
            f"{rel}: ชื่อไฟล์มีรหัส {name.station} แต่หัวไฟล์เป็นสถานี {head.station} — จัดเป็นสถานี {head.station}"
        )

    by_month = found.stations[kind].setdefault(station, {})
    key = (year, month)
    existing = by_month.get(key)
    new = DataFile(path=path, kind=kind, station=station, year=year, month=month)
    if existing is None:
        by_month[key] = new
        return
    if _same_content(path, existing.path):
        found.duplicate_copies += 1
        return
    keep_new = _prefer(path, existing.path)
    used, dropped = (path, existing.path) if keep_new else (existing.path, path)
    if keep_new:
        by_month[key] = new
    found.warnings.append(
        f"สถานี {station} เดือน {month:02d}/{year} มี 2 ไฟล์ที่เนื้อหาไม่เหมือนกัน — ใช้ "
        f"{_rel(used, root)} (ไม่ใช้ {_rel(dropped, root)})"
    )
