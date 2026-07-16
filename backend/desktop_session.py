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
import threading
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
_DESKTOP_SWITCHDESKTOP = 0x0100
_DESKTOP_NAME = "CropWatAutoRunnerDesktop"

# path มาตรฐานที่ตัวติดตั้ง CropWat 8.0 ใช้ — ลองหาให้เองถ้าผู้ใช้ยังไม่ได้ตั้งค่า
_DEFAULT_EXE_CANDIDATES = [
    r"C:\Program Files (x86)\CROPWAT\cropwat.exe",
    r"C:\Program Files\CROPWAT\cropwat.exe",
]


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


# --------------------------------------------------------------------------
# เข้า/ออก เดสก์ท็อปซ่อน (v0.7.3) — toggle เข้าดูตอนไหนก็ได้ ออกตอนไหนก็ได้
#
# ปัญหาที่ต้องแก้: พอ SwitchDesktop ไปเดสก์ท็อปซ่อน เมาส์/คีย์บอร์ดผู้ใช้จะอยู่
# ในนั้น แต่เดสก์ท็อปซ่อนไม่มี taskbar/shell เลย (มีแต่ CropWat) — ปุ่ม "กลับ" บน
# หน้าเว็บ (อยู่เดสก์ท็อปหลัก) กดไม่ได้ จึงต้องวาง "ปุ่มกลับ" ไว้บนเดสก์ท็อปซ่อน
# เอง (หน้าต่าง win32 ล้วน — ห้ามใช้ tkinter เพราะจะชนกับ tray icon แล้ว crash)
# ผู้ใช้คลิกปุ่มนั้นเมื่อไหร่ก็กลับจอหลักได้ — ไม่มี auto-timer 10 วิอีกต่อไป
#
# การสลับจอ 2 จังหวะ (เข้า/ออก) ยัง park automation ที่จุดปลอดภัยก่อนเสมอ กัน
# error "Cannot make a visible window modal" — แต่ "ระหว่างดู" ปล่อยให้ทำงานสด
# ผู้ใช้จะเห็น CropWat ทำงานจริงๆ (ไม่ได้หยุดนิ่ง) เพราะเดสก์ท็อปนิ่งแล้ว ไม่มี
# การสลับจอมาชน modal อีก / ตาข่ายสุดท้าย: Ctrl+Alt+Del แล้ว Cancel = กลับจอหลัก
# --------------------------------------------------------------------------

# ประสานงาน "park automation ชั่วคราวตอนสลับจอ": _pause_requested ตั้งตอนจะสลับ
# เพื่อขอให้ engine ไป park ที่จุดปลอดภัย (ระหว่างวันปลูก ไม่มี dialog เปิด) _parked
# ตั้งโดย engine เมื่อ park แล้ว — สลับจอหลังจากนั้นเสมอ
_pause_requested = threading.Event()
_parked = threading.Event()

# hwnd ของ "ปุ่มกลับ" บนเดสก์ท็อปซ่อน (สร้างโดย HiddenDesktopSession) — ให้
# enter/leave สั่ง show/hide ได้ (ShowWindow เรียกข้าม thread ได้)
_return_hwnd: Optional[int] = None
_viewing = threading.Event()  # ตั้งเมื่อกำลังดูเดสก์ท็อปซ่อนอยู่


def wait_if_peek_paused() -> None:
    """engine เรียกที่ "จุดปลอดภัย" ระหว่างวันปลูก — ถ้ามีการขอสลับจออยู่ ให้ park
    (บอกว่าพร้อมสลับได้แล้ว) แล้วค้างจนกว่าการสลับจอจะเสร็จ (ชั่วครู่)"""
    if not _pause_requested.is_set():
        return
    _parked.set()
    while _pause_requested.is_set():
        time.sleep(0.1)
    _parked.clear()


def hidden_desktop_exists() -> bool:
    """เช็คว่าเดสก์ท็อปซ่อนมีอยู่จริงตอนนี้ไหม (มีอยู่ = มี run ที่ใช้มันทำงานอยู่)"""
    h = _user32.OpenDesktopW(_DESKTOP_NAME, 0, False, _DESKTOP_SWITCHDESKTOP)
    if not h:
        return False
    _user32.CloseDesktop(h)
    return True


def is_viewing_hidden_desktop() -> bool:
    return _viewing.is_set()


def _park_for_switch() -> None:
    """ขอให้ engine park ที่จุดปลอดภัยก่อนสลับจอ แล้วรอจน park จริง (เพดาน 40 วิ
    เผื่อวันปลูกที่กำลังทำใช้เวลานาน/ไม่มี run active) — ครบเพดานก็สลับ best-effort"""
    _pause_requested.set()
    if not _parked.wait(timeout=40.0):
        logger.warning("รอ automation park ไม่ทันใน 40 วิ — สลับจอแบบ best-effort")


def enter_hidden_desktop() -> bool:
    """สลับจอเข้าไปดูเดสก์ท็อปซ่อน (ค้างอยู่จนกว่าจะกดปุ่มกลับ) — คืน False ถ้า
    เดสก์ท็อปซ่อนไม่มีอยู่ (ยังไม่ได้กำลังรันโหมดนี้)"""
    hdesk = _user32.OpenDesktopW(_DESKTOP_NAME, 0, False, _DESKTOP_SWITCHDESKTOP)
    if not hdesk:
        return False
    _park_for_switch()
    switched = bool(_user32.SwitchDesktop(hdesk))
    _user32.CloseDesktop(hdesk)
    _pause_requested.clear()  # ปล่อยให้ทำงานสดระหว่างดู (เดสก์ท็อปนิ่งแล้ว ปลอดภัย)
    if not switched:
        raise DesktopSessionError(
            f"สลับไปเดสก์ท็อปซ่อนไม่สำเร็จ (SwitchDesktop err={ctypes.get_last_error()})"
        )
    _viewing.set()
    if _return_hwnd:
        _user32.ShowWindow(_return_hwnd, 5)  # SW_SHOW
        _user32.SetWindowPos(_return_hwnd, -1, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0010)  # TOPMOST,NOMOVE,NOSIZE,NOACTIVATE
    logger.info("เข้าดูเดสก์ท็อปซ่อน (ค้างจนกว่าจะกดปุ่มกลับ)")
    return True


def leave_hidden_desktop() -> None:
    """สลับจอกลับเดสก์ท็อปหลัก (Default) + ซ่อนปุ่มกลับ + ปลด park ให้ทำงานต่อ —
    เรียกได้จากทั้งการคลิกปุ่มกลับ (บนเดสก์ท็อปซ่อน) และจากภายนอก"""
    if _return_hwnd:
        _user32.ShowWindow(_return_hwnd, 0)  # SW_HIDE
    _park_for_switch()
    h = _user32.OpenDesktopW("Default", 0, False, _DESKTOP_SWITCHDESKTOP)
    if h:
        _user32.SwitchDesktop(h)
        _user32.CloseDesktop(h)
    _viewing.clear()
    _pause_requested.clear()
    logger.info("กลับเดสก์ท็อปหลักแล้ว")


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
        self._start_return_window()
        return self.pid

    def _start_return_window(self) -> None:
        """สร้าง "ปุ่มกลับ" (หน้าต่าง win32 ล้วน) บนเดสก์ท็อปซ่อน รันใน thread ของ
        ตัวเอง (bind เข้าเดสก์ท็อปซ่อน) — เริ่มต้นซ่อนไว้ โผล่เมื่อ enter_hidden_
        desktop() สั่ง show คลิกที่หน้าต่าง (หรือปิด) = กลับจอหลัก"""
        hdesk = self._hdesk

        def _loop() -> None:
            global _return_hwnd
            try:
                _user32.SetThreadDesktop(hdesk)

                def _wndproc(hwnd, msg, wparam, lparam):
                    if msg in (win32con.WM_LBUTTONDOWN, win32con.WM_CLOSE):
                        try:
                            leave_hidden_desktop()
                        except Exception:  # noqa: BLE001
                            logger.exception("กลับจอหลักจากปุ่มกลับไม่สำเร็จ")
                        return 0
                    if msg == win32con.WM_PAINT:
                        self._paint_return_window(hwnd)
                        return 0
                    if msg == win32con.WM_DESTROY:
                        win32gui.PostQuitMessage(0)
                        return 0
                    return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

                wc = win32gui.WNDCLASS()
                wc.lpszClassName = "CropWatReturnWindow"
                wc.lpfnWndProc = _wndproc
                wc.hbrBackground = win32con.COLOR_INFOBK + 1
                wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_HAND)
                cls = win32gui.RegisterClass(wc)
                # กลางบนของจอ กว้างพออ่านง่าย
                sw = _user32.GetSystemMetrics(0)
                width, height = 620, 130
                hwnd = win32gui.CreateWindowEx(
                    win32con.WS_EX_TOPMOST,
                    cls,
                    "กลับหน้าจอหลัก",
                    win32con.WS_POPUP | win32con.WS_CAPTION | win32con.WS_SYSMENU,
                    (sw - width) // 2, 40, width, height,
                    0, 0, 0, None,
                )
                _return_hwnd = hwnd
                win32gui.ShowWindow(hwnd, 0)  # SW_HIDE — โผล่ตอน enter เท่านั้น
                win32gui.PumpMessages()
            except Exception:  # noqa: BLE001 -- ปุ่มกลับพังไม่ควรล้มการรัน (ยังมี Ctrl+Alt+Del)
                logger.exception("สร้างปุ่มกลับบนเดสก์ท็อปซ่อนไม่สำเร็จ")

        threading.Thread(target=_loop, daemon=True, name="hidden-desktop-return-btn").start()

    @staticmethod
    def _paint_return_window(hwnd: int) -> None:
        hdc, ps = win32gui.BeginPaint(hwnd)
        try:
            rect = win32gui.GetClientRect(hwnd)
            win32gui.SetBkMode(hdc, win32con.TRANSPARENT)
            win32gui.DrawText(
                hdc,
                "◀  คลิกที่นี่ เพื่อกลับหน้าจอหลัก  ◀\n(หรือปิดหน้าต่างนี้ / กด Ctrl+Alt+Del แล้ว Cancel)",
                -1, rect,
                win32con.DT_CENTER | win32con.DT_VCENTER | win32con.DT_WORDBREAK,
            )
        finally:
            win32gui.EndPaint(hwnd, ps)

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
        global _return_hwnd
        # ถ้ากำลังดูเดสก์ท็อปซ่อนอยู่ตอน run จบ ต้องสลับจอกลับก่อนทำลาย desktop
        # ไม่งั้นจอผู้ใช้ค้างอยู่บนเดสก์ท็อปที่กำลังจะหายไป
        if _viewing.is_set():
            try:
                leave_hidden_desktop()
            except Exception:  # noqa: BLE001
                pass
        if _return_hwnd:
            try:
                win32gui.PostMessage(_return_hwnd, win32con.WM_DESTROY, 0, 0)
            except Exception:  # noqa: BLE001
                pass
            _return_hwnd = None
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
