"""
os_dialogs.py
==============
Native Windows folder-picker + "open in Explorer" — ใช้ tkinter.filedialog
(มากับ Python อยู่แล้ว เหมือนที่ overlay.py/cropwat_engine.py ใช้ tkinter อยู่แล้ว)
ให้ผู้ใช้เลือก path โฟลเดอร์ด้วยการคลิกแทนการพิมพ์/แปะเอง — ลดจุดพังจากการพิมพ์
path ผิดสำหรับผู้ใช้ที่ไม่ถนัดคอมพิวเตอร์ (v1.0 UX pass)

ต้องเรียกผ่าน asyncio.to_thread() จาก endpoint เสมอ (เป็น blocking call รอผู้ใช้
คลิกเลือกโฟลเดอร์/ปิด dialog) ห้ามเรียกตรงจาก async def เพราะจะบล็อก event loop
ทั้งตัวจนกว่าผู้ใช้จะปิด dialog — สร้าง Tk root ของตัวเองแยกทุกครั้ง ไม่ยุ่งกับ
root ของ overlay.py (รันอยู่ใน thread ของตัวเองอยู่แล้ว)
"""

from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog


def pick_folder(initial_dir: str = "", title: str = "เลือกโฟลเดอร์") -> str | None:
    """เปิด native folder picker ของ Windows คืน path ที่เลือก หรือ None ถ้ากด
    ยกเลิก/ปิดหน้าต่างเฉยๆ"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    start = initial_dir if initial_dir and Path(initial_dir).is_dir() else None
    try:
        chosen = filedialog.askdirectory(parent=root, initialdir=start, title=title)
    finally:
        root.destroy()
    return chosen or None


def pick_file(initial_dir: str = "", title: str = "เลือกไฟล์") -> str | None:
    """เปิด native file picker ของ Windows (กรองเฉพาะ .exe — ใช้เลือกไฟล์โปรแกรม
    CropWat 8.0) คืน path ที่เลือก หรือ None ถ้ากดยกเลิก"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    start = initial_dir if initial_dir and Path(initial_dir).is_dir() else None
    try:
        chosen = filedialog.askopenfilename(
            parent=root,
            initialdir=start,
            title=title,
            filetypes=[("Program", "*.exe"), ("All files", "*.*")],
        )
    finally:
        root.destroy()
    return chosen or None


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
