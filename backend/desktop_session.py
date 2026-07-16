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

v0.8.0 — ยืนยันจากผู้ใช้จริงว่า "เข้าดูเดสก์ท็อปซ่อน" (v0.7.3-v0.7.5, สลับจอด้วย
SwitchDesktop) ยังชนกับ "Cannot make a visible window modal" ซ้ำๆ แม้พยายามแก้
หลายรอบแล้ว (เพราะการสลับ input desktop ชนจังหวะ CropWat โชว์ modal ได้ตลอดเวลา
ไม่มีทางป้องกันได้ 100% — ผู้ใช้ตัดสินใจถอดฟีเจอร์นี้ทิ้งทั้งหมด) แทนที่ด้วย
"ถ่ายภาพหน้าจอตอนนี้" (peek screenshot) ผ่าน PrintWindow ล้วนๆ — ไม่ต้องสลับจอ
เลย จึงไม่มีความเสี่ยงชน modal แบบเดิมอีกต่อไป (ใช้กลไกเดียวกับที่พิสูจน์แล้วว่า
เสถียรมากสำหรับ screenshot ที่ต้องส่งงาน)
"""

from __future__ import annotations

import ctypes
import logging
import threading
import time
from pathlib import Path
from typing import Optional

import win32api
import win32process
from pywinauto.findwindows import find_windows

logger = logging.getLogger("desktop_session")

_user32 = ctypes.windll.user32
_GENERIC_ALL = 0x10000000
_DESKTOP_SWITCHDESKTOP = 0x0100
_DESKTOP_NAME = "CropWatAutoRunnerDesktop"

# path มาตรฐานที่ตัวติดตั้ง CropWat 8.0 ใช้ — ลองหาให้เองถ้าผู้ใช้ยังไม่ได้ตั้งค่า
_DEFAULT_EXE_CANDIDATES = [
    r"C:\Program Files (x86)\CROPWAT\cropwat.exe",
    r"C:\Program Files\CROPWAT\cropwat.exe",
]

# ชื่อโฟลเดอร์เก็บภาพ "แอบดู" (สำหรับเช็คเฉยๆ) — แยกออกจาก screenshots/ ที่เป็น
# ของจริงที่ต้องส่งงาน (ดู config.screenshot_dir) เด็ดขาด กันปนกันตอนสร้าง
# Screenshots.docx (ไฟล์นี้เก็บแค่ภาพล่าสุด เขียนทับทุกครั้ง ไม่ใช่ของที่ต้องเก็บ
# ประวัติ)
PEEK_SCREENSHOT_DIR_NAME = "_peek_screenshots"


class DesktopSessionError(Exception):
    """เปิดเดสก์ท็อปซ่อน/launch CropWat ในนั้นไม่สำเร็จ"""


def resolve_cropwat_exe(configured_path: str) -> Optional[Path]:
    """คืน path ไฟล์ CropWat ที่ใช้ได้จริง: ใช้ค่าที่ตั้งไว้ก่อน ถ้าว่าง/ไม่มีจริง
    ลองหาตาม path มาตรฐานของตัวติดตั้ง — คืน None ถ้าไม่เจอเลย"""
    if configured_path:
        p = Path(configured_path)
        if p.is_file():
            return p
    for cand in _DEFAULT_EXE_CANDIDATES:
        p = Path(cand)
        if p.is_file():
            return p
    return None


def hidden_desktop_exists() -> bool:
    """เช็คว่าเดสก์ท็อปซ่อนมีอยู่จริงตอนนี้ไหม (มีอยู่ = มี run ที่ใช้มันทำงานอยู่)"""
    h = _user32.OpenDesktopW(_DESKTOP_NAME, 0, False, _DESKTOP_SWITCHDESKTOP)
    if not h:
        return False
    _user32.CloseDesktop(h)
    return True


# --------------------------------------------------------------------------
# ถ่ายภาพหน้าจอตอนนี้ (peek screenshot, v0.8.0) — ขอ/รับผลผ่าน Event สื่อสารกับ
# automation thread เพราะ PrintWindow ต้องเรียกจาก thread ที่ bind กับเดสก์ท็อป
# ซ่อนอยู่ (เหมือนเหตุผลเดียวกับที่ pywinauto ต้องรันบน thread ที่ bind ไว้) —
# engine เช็ค+บริการคำขอที่ "จุดปลอดภัย" เดียวกับที่เคยใช้ pause automation (ก่อน
# ทุกคำสั่งเมนู ใน _invoke_menu) แต่ตอนนี้แค่ถ่ายภาพเฉยๆ ไม่ต้อง pause/park อะไร
# เลย (PrintWindow มี timeout ในตัวอยู่แล้ว ~8 วิ ไม่บล็อกอะไรจริงจัง)
# --------------------------------------------------------------------------

_screenshot_request = threading.Event()
_screenshot_ready = threading.Event()
_screenshot_path: Optional[Path] = None
_screenshot_error: Optional[str] = None


def request_peek_screenshot(output_dir: str, timeout: float = 15.0) -> Path:
    """เรียกจาก API endpoint — ขอให้ automation thread ถ่ายภาพหน้าจอปัจจุบันให้
    แล้วรอผล (ปกติเสร็จใน 1-2 วิ เพราะเช็คถี่ที่ทุกคำสั่งเมนู) คืน path ไฟล์ภาพที่
    ถ่ายเสร็จแล้ว — โยน DesktopSessionError ถ้าไม่มี run อยู่ (ไม่มีใครมาบริการคำขอ)
    หรือถ่ายไม่สำเร็จ"""
    global _screenshot_path, _screenshot_error
    _screenshot_path = None
    _screenshot_error = None
    _screenshot_ready.clear()
    _screenshot_request.set()
    if not _screenshot_ready.wait(timeout=timeout):
        _screenshot_request.clear()
        raise DesktopSessionError(
            "รอภาพหน้าจอไม่ทัน — อาจไม่มีการรันในโหมดเดสก์ท็อปซ่อนอยู่ตอนนี้ ลองใหม่อีกครั้ง"
        )
    if _screenshot_error:
        raise DesktopSessionError(_screenshot_error)
    if _screenshot_path is None:
        raise DesktopSessionError("ถ่ายภาพไม่สำเร็จ (ไม่ทราบสาเหตุ)")
    return _screenshot_path


def service_peek_request(engine, output_dir: str) -> None:
    """เรียกจาก safe point ใน engine (ก่อนทุกคำสั่งเมนู ดู CropWatEngine.pause_check
    / _invoke_menu) — ถ้ามีคำขอถ่ายภาพค้างอยู่ ให้ถ่ายให้ทันทีด้วย PrintWindow
    (เร็ว มี timeout ในตัว ~8 วิ ไม่ต้อง pause/park อะไรเลยต่างจากการสลับจอเดิม)
    เก็บแยกโฟลเดอร์จาก screenshot ที่ต้องส่งงานเสมอ (PEEK_SCREENSHOT_DIR_NAME)"""
    global _screenshot_path, _screenshot_error
    if not _screenshot_request.is_set():
        return
    _screenshot_request.clear()
    try:
        peek_dir = Path(output_dir) / PEEK_SCREENSHOT_DIR_NAME
        peek_dir.mkdir(parents=True, exist_ok=True)
        path = peek_dir / "latest.png"
        engine._capture_main_window(path)  # noqa: SLF001 -- ใช้ตัวเดียวกับ screenshot ที่ต้องส่งงาน
        _screenshot_path = path
    except Exception as exc:  # noqa: BLE001
        logger.exception("ถ่ายภาพหน้าจอ (peek) ไม่สำเร็จ")
        _screenshot_error = str(exc)
    finally:
        _screenshot_ready.set()


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
        exe = resolve_cropwat_exe(self.exe_path)
        if exe is None:
            raise DesktopSessionError(
                "หาไฟล์โปรแกรม CropWat ไม่เจอ — โหมดเดสก์ท็อปซ่อนต้องตั้ง path ไฟล์ "
                ".exe ของ CropWat ในหน้าตั้งค่า (การ์ด 'โปรแกรม CropWat 8.0') ก่อน "
                "เพราะโปรแกรมต้องเปิด CropWat ให้เองในเดสก์ท็อปซ่อน (ลองหาที่ "
                "C:\\Program Files (x86)\\CROPWAT\\ ให้แล้วก็ไม่เจอ)"
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
        import win32con
        import win32gui

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
