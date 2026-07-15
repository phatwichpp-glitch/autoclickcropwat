"""
desktop_session.py
===================
โหมด "เดสก์ท็อปซ่อน" (v0.6.0) — รัน CropWat บน Win32 Desktop object แยกต่างหาก
ที่มองไม่เห็นบนจอผู้ใช้ เมาส์/คีย์บอร์ดบนจอหลักไม่แตะของในเดสก์ท็อปนี้เลย และ
automation ในนั้นจะคลิก/แย่งโฟกัสยังไงก็ไม่กระทบจอผู้ใช้

พิสูจน์แล้วด้วยการทดสอบสด (prototype): CropWat เปิด/คำนวณ/print/จับภาพ (PrintWindow)
บนเดสก์ท็อปซ่อนได้ครบ และรันข้ามเดือน/ข้ามปีได้โดยไม่ crash — ต่างจากโหมดเบื้อง
หลังปกติที่ต้อง "เหวี่ยงหน้าต่างออกนอกจอ" ซึ่งเป็นตัวจุดชนวน error ภายในของ CropWat
("Cannot make a visible window modal") ตอนสลับไฟล์

ต่างจาก VM: เบากว่ามาก (ไม่ต้องติดตั้ง Windows/CropWat ซ้ำ) แต่ยังแยก input/display
ระดับ OS เหมือนกัน — ข้อจำกัด: ต้องให้โปรแกรมเปิด CropWat ให้เอง (ผู้ใช้เปิดเอง
ไม่ได้เพราะอยู่คนละเดสก์ท็อป) จึงต้องตั้ง path ไฟล์ .exe ของ CropWat ไว้ก่อน

pywin32 มี CreateDesktop/GetThreadDesktop แต่ไม่มี SetThreadDesktop/CloseDesktop
เลยใช้ ctypes เรียก user32 ตรงสำหรับสองตัวนั้น
"""

from __future__ import annotations

import ctypes
import logging
import time
from pathlib import Path
from typing import Optional

import win32api
import win32con
import win32gui
import win32process
from pywinauto.findwindows import find_windows

logger = logging.getLogger("desktop_session")

_user32 = ctypes.windll.user32
_GENERIC_ALL = 0x10000000
_DESKTOP_NAME = "CropWatAutoRunnerDesktop"


class DesktopSessionError(Exception):
    """เปิดเดสก์ท็อปซ่อน/launch CropWat ในนั้นไม่สำเร็จ"""


class HiddenDesktopSession:
    """จัดการวงจรชีวิตของเดสก์ท็อปซ่อน + CropWat instance ที่รันในนั้น

    กติกาการใช้ (สำคัญ): bind_and_launch() ต้องถูกเรียกจาก "thread เดียวกัน" กับที่
    จะรัน automation ทั้งหมด เพราะ SetThreadDesktop เปลี่ยน desktop ของ thread ที่
    เรียกเท่านั้น — pywinauto/win32 หลังจากนั้นจะ enumerate หน้าต่างบนเดสก์ท็อป
    ซ่อนโดยอัตโนมัติ ปกติเรียกจาก runner background thread (_run_years)"""

    def __init__(self, cropwat_exe_path: str) -> None:
        self.exe_path = cropwat_exe_path
        self._hdesk: Optional[int] = None
        self._orig_desktop: Optional[int] = None
        self._proc_handle = None
        self.pid: Optional[int] = None

    def bind_and_launch(self) -> int:
        """สร้างเดสก์ท็อปซ่อน → ผูก thread ปัจจุบันเข้ากับมัน → เปิด CropWat ในนั้น
        → รอจนหน้าต่างหลักโผล่ + ปิด Welcome dialog คืน pid ของ CropWat ที่เปิด"""
        exe = Path(self.exe_path)
        if not exe.is_file():
            raise DesktopSessionError(
                f"ไม่พบไฟล์โปรแกรม CropWat: {self.exe_path} — โหมดเดสก์ท็อปซ่อนต้อง"
                "ตั้ง path ไฟล์ .exe ของ CropWat ในหน้าตั้งค่าก่อน (เพราะโปรแกรมต้อง"
                "เปิด CropWat ให้เองในเดสก์ท็อปซ่อน)"
            )

        # เก็บ desktop เดิมของ thread ไว้คืนตอนจบ (GetThreadDesktop คืน pseudo-
        # handle ไม่ต้อง close) — ให้ CloseDesktop สำเร็จ (ปิด desktop ที่ thread
        # ยัง bind อยู่ไม่ได้)
        self._orig_desktop = _user32.GetThreadDesktop(
            ctypes.windll.kernel32.GetCurrentThreadId()
        )

        self._hdesk = _user32.CreateDesktopW(
            _DESKTOP_NAME, None, None, 0, _GENERIC_ALL, None
        )
        if not self._hdesk:
            raise DesktopSessionError(
                f"สร้างเดสก์ท็อปซ่อนไม่สำเร็จ (CreateDesktopW err={ctypes.get_last_error()})"
            )

        # ผูก thread ปัจจุบันเข้ากับเดสก์ท็อปซ่อน — ต้องทำ "ก่อน" ใช้ pywinauto ใดๆ
        # (SetThreadDesktop จะล้มเหลวถ้า thread มีหน้าต่างอยู่แล้ว — runner thread
        # ไม่สร้างหน้าต่างจึงปลอดภัย)
        if not _user32.SetThreadDesktop(self._hdesk):
            err = ctypes.get_last_error()
            _user32.CloseDesktop(self._hdesk)
            self._hdesk = None
            raise DesktopSessionError(
                f"ผูก thread เข้ากับเดสก์ท็อปซ่อนไม่สำเร็จ (SetThreadDesktop err={err})"
            )

        # เปิด CropWat บนเดสก์ท็อปซ่อน
        startup = win32process.STARTUPINFO()
        startup.lpDesktop = _DESKTOP_NAME
        try:
            self._proc_handle, _thr, self.pid, _tid = win32process.CreateProcess(
                str(exe), None, None, None, False, 0, None, str(exe.parent), startup
            )
        except Exception as exc:  # noqa: BLE001
            self.stop()
            raise DesktopSessionError(f"เปิด CropWat ในเดสก์ท็อปซ่อนไม่สำเร็จ: {exc}") from exc

        logger.info("เปิด CropWat (pid=%s) บนเดสก์ท็อปซ่อนแล้ว", self.pid)
        self._wait_ready_and_dismiss_welcome()
        return self.pid

    def _wait_ready_and_dismiss_welcome(self, timeout: float = 20.0) -> None:
        deadline = time.monotonic() + timeout
        main_hwnd = None
        while time.monotonic() < deadline:
            hs = find_windows(
                title_re="CROPWAT.*", class_name="TMainForm",
                top_level_only=True, process=self.pid, visible_only=False,
            )
            if hs:
                main_hwnd = hs[0]
                break
            time.sleep(0.5)
        if main_hwnd is None:
            raise DesktopSessionError(
                "เปิด CropWat แล้วแต่หน้าต่างหลักไม่โผล่ในเดสก์ท็อปซ่อนภายในเวลาที่กำหนด"
            )
        # ปิด Welcome dialog ที่อาจเด้งตอนเปิดโปรแกรมใหม่
        time.sleep(1.0)
        for h in find_windows(
            class_name="TWelcomeForm", top_level_only=True,
            process=self.pid, visible_only=False,
        ):
            win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
            logger.info("ปิด Welcome dialog ของ CropWat อัตโนมัติ")

    def stop(self) -> None:
        """ปิด CropWat ที่เราเปิดในเดสก์ท็อปซ่อน + ปล่อย handle เดสก์ท็อป — เป็น
        process ที่โปรแกรมนี้เปิดเอง (ไม่ใช่ของผู้ใช้) จึง terminate ได้ปลอดภัย"""
        if self._proc_handle is not None:
            try:
                win32api.TerminateProcess(self._proc_handle, 0)
                logger.info("ปิด CropWat (pid=%s) บนเดสก์ท็อปซ่อนแล้ว", self.pid)
            except Exception:  # noqa: BLE001 -- process ปิดไปเองแล้ว = ปกติ
                pass
            self._proc_handle = None
        # คืน thread กลับไป desktop เดิมก่อน แล้วค่อยปิด desktop ซ่อน (ปิด desktop
        # ที่ thread ยัง bind อยู่ไม่ได้)
        if self._orig_desktop is not None:
            try:
                _user32.SetThreadDesktop(self._orig_desktop)
            except Exception:  # noqa: BLE001
                pass
            self._orig_desktop = None
        if self._hdesk is not None:
            try:
                _user32.CloseDesktop(self._hdesk)
            except Exception:  # noqa: BLE001
                pass
            self._hdesk = None
