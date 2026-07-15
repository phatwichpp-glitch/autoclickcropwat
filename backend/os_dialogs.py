"""
os_dialogs.py
==============
Native Windows folder/file-picker + "open in Explorer" / "launch exe"

v0.5.31 — CRASH FIX ยืนยันจาก Event Viewer จริงของผู้ใช้: โปรแกรมทั้งตัวปิดตัว
เองแบบไม่มี log/MessageBox เลย ("Faulting module: tcl86t.dll, Exception code:
0x80000003" ซ้ำหลายรอบ) ต้นเหตุคือไฟล์นี้เคย (v0.5.26) ใช้ tkinter.filedialog
สร้าง tk.Tk() root ใหม่ทุกครั้งที่กดปุ่ม Browse บน thread ของ FastAPI
(asyncio.to_thread) — ในขณะที่ overlay.py มี tk.Tk() root ของ tray icon/overlay
รันอยู่แล้วตลอดเวลาบน thread แยกของตัวเอง (มีชีวิตอยู่ตั้งแต่เปิดโปรแกรมจนปิด
ไม่ใช่แค่ตอนรัน) — Tcl/Tk **ไม่รองรับหลาย interpreter คนละ thread ในโปรเซส
เดียวกัน** พอมี 2 root พร้อมกันคนละ thread, Tcl ยิง internal consistency check
(breakpoint trap) ทำให้ process ทั้งตัวตายทันทีแบบไม่มีทาง catch เป็น Python
exception ได้เลย (เป็น native crash ระดับ DLL)

แก้โดยเลิกใช้ tkinter ในไฟล์นี้เด็ดขาด เปลี่ยนไปใช้ Win32 common dialog ตรงๆ
แทน (win32gui.GetOpenFileNameW สำหรับไฟล์, Shell.Application COM object สำหรับ
โฟลเดอร์) ทั้งคู่ไม่แตะ Tcl/Tk เลย ปลอดภัย 100% ต่อให้เรียกจาก thread ไหนก็ตาม
พร้อมกันกับ overlay ที่รันอยู่ — ทั้งสอง API เป็น dependency ที่มีอยู่แล้ว
(win32gui/win32com ใช้อยู่แล้วในโปรเจกต์นี้)
"""

from __future__ import annotations

import os
from pathlib import Path

import win32com.client
import win32con
import win32gui


def pick_folder(initial_dir: str = "", title: str = "เลือกโฟลเดอร์") -> str | None:
    """เปิด native folder picker ของ Windows (Shell.Application.BrowseForFolder)
    คืน path ที่เลือก หรือ None ถ้ากดยกเลิก/ปิดหน้าต่างเฉยๆ"""
    start = initial_dir if initial_dir and Path(initial_dir).is_dir() else ""
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        folder = shell.BrowseForFolder(0, title, 0, start)
        if folder is None:
            return None
        return folder.Self.Path or None
    except Exception:  # noqa: BLE001 -- ผู้ใช้กดยกเลิก หรือ COM ล้มเหลว = ถือว่าไม่ได้เลือก
        return None


def pick_file(initial_dir: str = "", title: str = "เลือกไฟล์") -> str | None:
    """เปิด native file picker ของ Windows (win32gui.GetOpenFileNameW — Win32
    common dialog ดั้งเดิม กรองเฉพาะ .exe) คืน path ที่เลือก หรือ None ถ้ากด
    ยกเลิก (GetOpenFileNameW โยน exception ตอนยกเลิก ไม่ใช่คืนค่าว่าง)"""
    start = initial_dir if initial_dir and Path(initial_dir).is_dir() else ""
    try:
        filename, _, _ = win32gui.GetOpenFileNameW(
            InitialDir=start,
            Filter="Program (*.exe)\0*.exe\0All files (*.*)\0*.*\0\0",
            Title=title,
            Flags=win32con.OFN_FILEMUSTEXIST | win32con.OFN_PATHMUSTEXIST | win32con.OFN_EXPLORER,
        )
    except Exception:  # noqa: BLE001 -- ผู้ใช้กดยกเลิก = โยน exception ไม่ใช่คืนค่าว่าง
        return None
    return filename or None


def open_in_explorer(path: str) -> None:
    """เปิดโฟลเดอร์ใน Windows Explorer — โยน FileNotFoundError ถ้า path ไม่ใช่
    โฟลเดอร์จริง (caller แปลงเป็น HTTP error ให้ผู้ใช้เห็นสาเหตุชัดเจน)"""
    p = Path(path)
    if not p.is_dir():
        raise FileNotFoundError(f"ไม่พบโฟลเดอร์: {path}")
    os.startfile(str(p))  # type: ignore[attr-defined]  # Windows-only API


def launch_exe(path: str) -> None:
    """เปิดโปรแกรม (เช่น CropWat 8.0) จาก path ที่ตั้งค่าไว้ — โยน
    FileNotFoundError ถ้า path ไม่ใช่ไฟล์จริง"""
    p = Path(path)
    if not p.is_file():
        raise FileNotFoundError(f"ไม่พบไฟล์โปรแกรม: {path}")
    os.startfile(str(p))  # type: ignore[attr-defined]  # Windows-only API
