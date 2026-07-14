"""
updater.py
==========
ระบบเช็คอัปเดต + อัปเดตตัวเองจาก GitHub Releases (v0.4.0)

**เช็ค**: เรียก GitHub API หา release ล่าสุดของ repo แล้วเทียบ tag กับ APP_VERSION
ที่ฝังอยู่ในตัวโปรแกรม (version.py) — ไม่ต้องล็อกอินเพราะ repo เป็น public

**อัปเดตตัวเอง**: ไฟล์ .exe ที่กำลังรันอยู่เขียนทับตัวเองไม่ได้ (Windows lock ไว้)
ใช้เทคนิคมาตรฐานของโปรแกรม self-update ทั่วไป:
  1. ดาวน์โหลด .exe เวอร์ชันใหม่มาวางข้างๆ (ชื่อ .update.exe)
  2. เขียนสคริปต์ .bat เล็กๆ ที่วนพยายาม move ทับไฟล์เดิม (จะสำเร็จก็ต่อเมื่อ
     โปรแกรมเดิมปิดไปแล้วเท่านั้น — ใช้การ move ล้มเหลวเป็นตัวรอในตัว)
     แล้วเปิดโปรแกรมใหม่ให้เอง เสร็จแล้วลบตัวสคริปต์ทิ้ง
  3. spawn สคริปต์แบบ detached (ไม่ตายตามโปรแกรมแม่) แล้วปิดตัวเอง
"""

from __future__ import annotations

import json
import logging
import os
import re
import subprocess
import sys
import threading
import urllib.request
from pathlib import Path

from version import APP_VERSION

logger = logging.getLogger("updater")

REPO = "phatwichpp-glitch/autoclickcropwat"
ASSET_NAME = "CropWatAutoRunner.exe"
API_LATEST = f"https://api.github.com/repos/{REPO}/releases/latest"


class UpdateError(Exception):
    """อัปเดตไม่ได้ด้วยเหตุที่อธิบายให้ผู้ใช้เข้าใจได้ (ไม่ใช่ bug)"""


def _version_tuple(v: str) -> tuple[int, ...]:
    nums = re.findall(r"\d+", v or "")
    return tuple(int(n) for n in nums[:3]) if nums else (0,)


def check_for_update() -> dict:
    """คืน dict สถานะเวอร์ชัน — ห้าม raise เด็ดขาด (frontend เรียกทุกครั้งที่เปิด
    โปรแกรม ถ้าออฟไลน์/GitHub ล่มก็แค่รายงานว่าเช็คไม่ได้ ไม่ใช่ error ของระบบ)"""
    result = {
        "current": APP_VERSION,
        "latest": None,
        "update_available": False,
        "notes": "",
        "asset_url": None,
        "error": None,
    }
    try:
        req = urllib.request.Request(
            API_LATEST,
            headers={"User-Agent": "CropWatAutoRunner", "Accept": "application/vnd.github+json"},
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.load(resp)
        latest_tag = (data.get("tag_name") or "").lstrip("v")
        asset_url = next(
            (a.get("browser_download_url") for a in data.get("assets", []) if a.get("name") == ASSET_NAME),
            None,
        )
        result["latest"] = latest_tag
        result["notes"] = (data.get("body") or "")[:2000]
        result["asset_url"] = asset_url
        result["update_available"] = (
            asset_url is not None and _version_tuple(latest_tag) > _version_tuple(APP_VERSION)
        )
    except Exception as exc:  # noqa: BLE001 -- ออฟไลน์/rate limit/GitHub ล่ม = เช็คไม่ได้เฉยๆ
        logger.warning("เช็คอัปเดตไม่สำเร็จ: %s", exc)
        result["error"] = f"เช็คอัปเดตไม่สำเร็จ (ดูอินเทอร์เน็ต): {exc}"
    return result


# สคริปต์สลับไฟล์: วน move ทับจนกว่าจะสำเร็จ (สำเร็จได้ก็ต่อเมื่อโปรแกรมเดิมปิด
# แล้วเท่านั้น — Windows lock ไฟล์ .exe ที่กำลังรันไว้) แล้วเปิดตัวใหม่ ลบตัวเองทิ้ง
# ping 127.0.0.1 = วิธี sleep ~1 วินาทีแบบพกพาของ .bat (timeout ใช้ใน non-console
# ไม่ได้เสมอไป) — จำกัด 60 รอบ (~1 นาที) กันวนไม่รู้จบถ้ามีอะไรผิดปกติ
#
# สำคัญมาก (v0.4.2 + v0.5.3): ต้องล้าง env var ของ PyInstaller bootloader ให้ครบ
# ทุกตัวก่อนเปิดตัวใหม่ ไม่งั้นตัวใหม่สืบทอดค่าแล้วเข้าใจผิดว่าตัวเอง "แตกไฟล์ไป
# แล้ว" ที่โฟลเดอร์ temp ของตัวเก่า (ซึ่งถูกลบทิ้งแล้ว) → error "Failed to load
# Python DLL ..._MEIxxxx" — v0.4.2 ล้างแค่ _MEIPASS2 (ชื่อยุค PyInstaller 5) แต่
# เรา build ด้วย PyInstaller 6 ซึ่งใช้ _PYI_ARCHIVE_FILE/_PYI_PARENT_PROCESS_LEVEL
# เป็นหลัก (ยืนยันจากอาการจริง: relaunch ยังพังหลังล้างตัวเดียว) — ล้างหมดทุกยุค
_UPDATER_BAT = """@echo off
set "_MEIPASS2="
set "_PYI_APPLICATION_HOME_DIR="
set "_PYI_ARCHIVE_FILE="
set "_PYI_PARENT_PROCESS_LEVEL="
set "_PYI_SPLASH_IPC="
set RETRIES=0
:loop
ping -n 2 127.0.0.1 >nul
move /y "{new_exe}" "{target_exe}" >nul 2>&1
if errorlevel 1 (
  set /a RETRIES+=1
  if %RETRIES% lss 60 goto loop
  exit /b 1
)
start "" "{target_exe}" --updated
del "%~f0"
"""
# หมายเหตุ "--updated": บอก launcher ว่านี่คือการเปิดหลังอัปเดต — ไม่ต้องเปิด
# หน้าต่างโปรแกรมใหม่ เพราะหน้าต่างเดิมของผู้ใช้ยังเปิดอยู่และจะ reload ตัวเอง
# เมื่อ backend กลับมา (ยืนยันจากผู้ใช้: ไม่ใส่ flag นี้จะได้ 2 หน้าต่างซ้อนกัน)


def apply_update() -> None:
    """ดาวน์โหลดเวอร์ชันใหม่ + วางสคริปต์สลับไฟล์ + ปิดโปรแกรม — เรียกจาก endpoint
    ที่เช็คแล้วว่าไม่มี run ค้างอยู่ ตัวโปรแกรมจะปิดเองใน ~1 วินาทีหลังฟังก์ชันนี้คืนค่า"""
    if not getattr(sys, "frozen", False):
        raise UpdateError(
            "รันอยู่ในโหมดนักพัฒนา (python ตรงๆ) อัปเดตอัตโนมัติไม่ได้ — ใช้ git pull แทน"
        )

    info = check_for_update()
    if info["error"]:
        raise UpdateError(info["error"])
    if not info["update_available"]:
        raise UpdateError(f"เป็นเวอร์ชันล่าสุดอยู่แล้ว (v{APP_VERSION})")

    target_exe = Path(sys.executable)
    new_exe = target_exe.with_name("CropWatAutoRunner.update.exe")

    logger.info("กำลังดาวน์โหลด v%s จาก %s", info["latest"], info["asset_url"])
    req = urllib.request.Request(info["asset_url"], headers={"User-Agent": "CropWatAutoRunner"})
    try:
        with urllib.request.urlopen(req, timeout=600) as resp, open(new_exe, "wb") as f:
            while True:
                chunk = resp.read(256 * 1024)
                if not chunk:
                    break
                f.write(chunk)
    except Exception as exc:  # noqa: BLE001
        new_exe.unlink(missing_ok=True)
        raise UpdateError(f"ดาวน์โหลดเวอร์ชันใหม่ไม่สำเร็จ: {exc}") from exc

    # กันไฟล์ครึ่งๆ กลางๆ/หน้า error แทนไฟล์จริง — .exe จริงใหญ่ระดับ 20+MB เสมอ
    if new_exe.stat().st_size < 5_000_000:
        new_exe.unlink(missing_ok=True)
        raise UpdateError("ไฟล์ที่ดาวน์โหลดมาเล็กผิดปกติ — ยกเลิกการอัปเดตไว้ก่อน")

    bat_path = target_exe.with_name("cropwat_update.bat")
    bat_path.write_text(
        _UPDATER_BAT.format(new_exe=new_exe, target_exe=target_exe), encoding="ascii"
    )

    DETACHED_PROCESS = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    # ล้าง env ของ PyInstaller ก่อน spawn (ดูเหตุผลใน _UPDATER_BAT) — ทำทั้ง 2 ชั้น
    # (ที่นี่และในตัว .bat) กันพลาด เพราะ child สืบทอด env จาก Popen นี้โดยตรง
    child_env = os.environ.copy()
    for var in (
        "_MEIPASS2",
        "_PYI_APPLICATION_HOME_DIR",
        "_PYI_ARCHIVE_FILE",
        "_PYI_PARENT_PROCESS_LEVEL",
        "_PYI_SPLASH_IPC",
    ):
        child_env.pop(var, None)
    subprocess.Popen(
        ["cmd", "/c", str(bat_path)],
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
        close_fds=True,
        cwd=str(target_exe.parent),
        env=child_env,
    )
    logger.info("วางสคริปต์อัปเดตแล้ว — ปิดโปรแกรมใน 1 วินาที ตัวใหม่จะเปิดเอง")
    # หน่วง 1 วิ ให้ HTTP response ของ endpoint ส่งกลับถึงหน้าจอก่อนค่อยปิดตัวเอง
    threading.Timer(1.0, os._exit, args=(0,)).start()
