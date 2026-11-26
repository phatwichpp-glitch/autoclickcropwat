"""
system_capacity.py
==================
v1.3.0 — เลือกจำนวนหน้าต่าง CropWat ที่รันพร้อมกันให้อัตโนมัติตามกำลังเครื่อง

ดูจาก 3 อย่าง (เอาค่าที่น้อยที่สุด = ปลอดภัยสุด) แล้วบีบให้อยู่ 1..MAX_PARALLEL_WINDOWS:
- จำนวนเธรด CPU: ทุก 6 เธรด = 1 หน้าต่าง (v1.3.1 จากเดิม 4 — งานจริงปฏิทินค่าเริ่มต้นบน
  เครื่อง 32 เธรด 4 หน้าต่างยังเริ่มมีหน้าต่าง climate ซ้อน) และโหมดอัตโนมัติไม่เกิน AUTO_MAX
  (ตั้งเองได้ถึง MAX_PARALLEL_WINDOWS) — — แต่ละหน้าต่างใช้ CropWat 1 process + thread
  automation ของเราอีก 1 + ระบบ Windows ต้องวาดหน้าต่างให้ (ทดสอบสดบนเครื่อง 32 เธรด:
  4 หน้าต่างผ่าน 36/36, 8 หน้าต่างตก 3/48 เพราะ CropWat ตั้งชื่อหน้าต่างไม่ทัน)
- CPU ว่างตอนนี้: ถ้าผู้ใช้เปิดงานหนักอื่นอยู่ (CPU ถูกใช้ไปเยอะ) ลดลงตาม
- แรมว่าง: กันไว้ 1 GB ให้ระบบ + ~400 MB ต่อหน้าต่าง (CropWat จริงใช้ ~50 MB เผื่อเยอะ)

ใช้ Win32 API ผ่าน ctypes ล้วน (ไม่ต้องติดตั้ง psutil เพิ่มใน exe)
"""

from __future__ import annotations

import ctypes
import os
import time
from ctypes import wintypes

from config import MAX_PARALLEL_WINDOWS

# v1.3.1 — เพดานของโหมดอัตโนมัติ (ระวังไว้ก่อน ผู้ใช้เลือกเองได้ถึง MAX_PARALLEL_WINDOWS)
AUTO_MAX = 3


class _MEMORYSTATUSEX(ctypes.Structure):
    _fields_ = [
        ("dwLength", wintypes.DWORD),
        ("dwMemoryLoad", wintypes.DWORD),
        ("ullTotalPhys", ctypes.c_ulonglong),
        ("ullAvailPhys", ctypes.c_ulonglong),
        ("ullTotalPageFile", ctypes.c_ulonglong),
        ("ullAvailPageFile", ctypes.c_ulonglong),
        ("ullTotalVirtual", ctypes.c_ulonglong),
        ("ullAvailVirtual", ctypes.c_ulonglong),
        ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
    ]


def _memory_mb() -> tuple[float, float]:
    """(แรมทั้งหมด, แรมว่าง) หน่วย MB"""
    st = _MEMORYSTATUSEX()
    st.dwLength = ctypes.sizeof(_MEMORYSTATUSEX)
    if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(st)):
        return 0.0, 0.0
    return st.ullTotalPhys / 1048576, st.ullAvailPhys / 1048576


def _system_times() -> tuple[int, int, int]:
    idle, kernel, user = wintypes.FILETIME(), wintypes.FILETIME(), wintypes.FILETIME()
    ctypes.windll.kernel32.GetSystemTimes(ctypes.byref(idle), ctypes.byref(kernel), ctypes.byref(user))

    def _v(ft):
        return (ft.dwHighDateTime << 32) | ft.dwLowDateTime

    return _v(idle), _v(kernel), _v(user)


def cpu_busy_fraction(sample_seconds: float = 0.5) -> float:
    """สัดส่วน CPU ที่ถูกใช้อยู่ตอนนี้ (0..1) วัดจริงช่วงสั้นๆ — kernel time รวม idle อยู่แล้ว"""
    i1, k1, u1 = _system_times()
    time.sleep(sample_seconds)
    i2, k2, u2 = _system_times()
    total = (k2 - k1) + (u2 - u1)
    if total <= 0:
        return 0.0
    return max(0.0, min(1.0, 1.0 - (i2 - i1) / total))


def recommend(info: dict | None = None) -> dict:
    """คืน dict: recommended (จำนวนหน้าต่าง), threads, ram_total_gb, ram_free_gb,
    cpu_busy_pct, reason (ข้อความไทยสั้นๆ บอกว่าทำไมได้เลขนี้)"""
    if info is None:
        total_mb, avail_mb = _memory_mb()
        info = {
            "threads": os.cpu_count() or 1,
            "ram_total_mb": total_mb,
            "ram_free_mb": avail_mb,
            "cpu_busy": cpu_busy_fraction(),
        }
    threads = max(1, int(info["threads"]))
    busy = float(info["cpu_busy"])
    free_mb = float(info["ram_free_mb"])

    by_cpu = threads // 6
    by_load = int(threads * (1.0 - busy) // 3)
    by_ram = int((free_mb - 1024) // 400) if free_mb > 0 else MAX_PARALLEL_WINDOWS
    n = max(1, min(AUTO_MAX, by_cpu, by_load, by_ram))

    limit = min(
        (by_cpu, "จำนวนเธรด CPU"), (by_load, "เครื่องกำลังทำงานอื่นอยู่"), (by_ram, "แรมว่างเหลือน้อย"),
        key=lambda t: t[0],
    )
    if n >= AUTO_MAX:
        why = "เครื่องแรงพอ ใช้จำนวนสูงสุดของโหมดอัตโนมัติ"
    else:
        why = f"จำกัดด้วย{limit[1]}"
    return {
        "recommended": n,
        "max": MAX_PARALLEL_WINDOWS,
        "threads": threads,
        "ram_total_gb": round(float(info.get("ram_total_mb", 0)) / 1024, 1),
        "ram_free_gb": round(free_mb / 1024, 1),
        "cpu_busy_pct": round(busy * 100),
        "reason": f"CPU {threads} เธรด · แรมว่าง {free_mb / 1024:.1f} GB · "
                  f"เครื่องใช้งานอยู่ {busy * 100:.0f}% — {why}",
    }


def recommend_windows() -> int:
    try:
        return int(recommend()["recommended"])
    except Exception:  # noqa: BLE001 -- วัดไม่ได้ = ปลอดภัยไว้ก่อน
        return 1
