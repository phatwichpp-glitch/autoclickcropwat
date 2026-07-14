"""
automation/cropwat_engine.py
=============================
โมดูลควบคุม CropWat 8.0 ด้วย pywinauto ตาม flow ใน spec หัวข้อ "Flow การรันต่อ 1 ปี"

ออกแบบให้แยกเป็นฟังก์ชันทีละ step (เปิดไฟล์ climate -> เปิดไฟล์ rain -> ตั้งวันปลูก ->
calculate -> เปิด irrigation schedule -> print/export) เพื่อ debug ง่าย — เรียกทีละ
ฟังก์ชันจาก REPL หรือสคริปต์ทดสอบแยกได้โดยไม่ต้องรันทั้ง flow

ทุก step อ้างอิงชื่อ control จาก cropwat_controls.py เท่านั้น ไม่มี string ชื่อ
control จริงฝังอยู่ในไฟล์นี้ เพื่อให้แก้ config ที่เดียวจบเวลาชื่อ control เปลี่ยน

สำคัญ (ยืนยันจากการ inspect จริง): CropWat เป็น MDI app — แต่ละโมดูล (Climate/ETo,
Rain, Crop, Irrigation Schedule) เป็น MDI child window ที่ต้อง set_focus() ก่อน
เรียกเมนู "File -> Open"/"File -> Print" เสมอ เพราะเมนูพวกนี้ทำงานกับ MDI child ที่
active อยู่ตอนนั้น ไม่ใช่ทำงานกับหน้าต่างหลักตรงๆ — โมดูลจับด้วย class_name เพราะ
class คงที่ ต่างจาก title ที่เปลี่ยนตามไฟล์/ค่าที่โหลดอยู่

สำคัญมาก (ยืนยันจากผู้ใช้แล้ว — เปลี่ยนความเข้าใจเดิมทั้งหมด): 1 ปี ไม่ได้รันแค่ครั้ง
เดียว! แต่ละคอลัมน์ใน Excel Result sheet คือการทดลอง "ถ้าปลูกวันนี้จะเกิดอะไรขึ้น" —
เปิดไฟล์ climate/rain ของปีนั้น "ครั้งเดียว" แล้ว "วนซ้ำ" ตั้ง Planting date เป็นวันที่
ทดลองแต่ละวัน (list ที่ผู้ใช้เลือกเอง ไม่มีกฎตายตัว) สั่ง Calculate (CWR + Scheduling)
และ Print ใหม่ทุกครั้งที่เปลี่ยนวันปลูก — 1 ปีจึงมีได้หลายสิบรอบคำนวณ (นี่คือเหตุผลหลัก
ที่ต้องมี automation) ดู run_year() ที่เป็น orchestrator ระดับปีซึ่งวนเรียก
run_candidate_planting_date() ต่อวันปลูกที่ทดลอง

การแก้ path ของไฟล์ climate/rain ต่อปี (โฟลเดอร์ shift-year, นามสกุล .PED/.CRD,
กติกาเลือกเดือนตาม spec) เป็นหน้าที่ของ "File engine" (ขั้นที่ 3) ไม่ใช่ที่นี่ —
เมธอดในไฟล์นี้รับแค่ path ไฟล์ที่ resolve แล้วมาโดยตรง

หมายเหตุ: ฟังก์ชันในไฟล์นี้ยังรันไม่ได้จริงจนกว่าจะกรอก cropwat_controls.py
ให้ครบ (connect() จะโยน ControlsNotConfiguredError ถ้ายังไม่ครบ)
"""

from __future__ import annotations

import logging
import re
import threading
import time
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

import win32con
import win32gui
from pywinauto import Application
from pywinauto.findwindows import ElementAmbiguousError, ElementNotFoundError, find_windows
from pywinauto.timings import TimeoutError as PywinautoTimeoutError

import cropwat_controls as controls
from automation.exceptions import (
    ControlsNotConfiguredError,
    CropWatNotRunningError,
    CropWatReportedError,
    DuplicateWindowError,
    StepTimeoutError,
)

logger = logging.getLogger("cropwat_engine")


@dataclass(frozen=True)
class PlantingDateTask:
    """1 วันปลูกที่จะทดลองในปีหนึ่งๆ — capture_screenshot บอกว่าต้อง capture
    หน้าจอเพิ่มไหม (เป็น subset ของวันปลูกที่ทดลองทั้งหมด ยืนยันกับผู้ใช้แล้วว่า
    ไม่ใช่ทุกวันปลูกที่ทดลองจะต้อง capture — ส่วนใหญ่แค่บางวันเป็น checkpoint)"""

    planting_date: date
    capture_screenshot: bool = False


@dataclass
class CandidateRunResult:
    """ผลของการทดลองปลูก 1 วันที่ (1 คอลัมน์ใน Result sheet)"""

    year: int
    planting_date: date
    ok: bool
    error_message: Optional[str] = None
    exported_file: Optional[Path] = None
    schedule_screenshot: Optional[Path] = None
    graph_screenshot: Optional[Path] = None


@dataclass
class YearRunResult:
    """ผลรวมของทั้งปี — ok=True ก็ต่อเมื่อทุกวันปลูกที่ทดลองสำเร็จหมด
    (สถานะระดับปีที่ dashboard ใช้แสดงยังคงเป็น 4 สถานะเดิมตาม spec แต่เบื้องหลัง
    เป็นการรวมผลจากหลาย candidate)"""

    year: int
    ok: bool
    candidates: list[CandidateRunResult] = field(default_factory=list)
    error_message: Optional[str] = None
    # True = ปีนี้ถูกสั่งหยุดกลางคัน (ยังทำวันปลูกไม่ครบ) — runner ใช้ตัดสินว่า
    # ควร mark ปีนี้กลับเป็น "รอคิว" ให้รันใหม่ได้ ไม่ใช่ "เสร็จ/error"
    stopped: bool = False


class CropWatEngine:
    """ห่อ pywinauto Application ของ CropWat ไว้ 1 ตัว ใช้ตลอดการรันหลายปีต่อกัน
    (เปิด CropWat ครั้งเดียว ไม่เปิดใหม่ทุกปี เพื่อความเร็วและลดจุดพัง)

    background_mode (v0.5.0, ทดลอง): ควบคุม CropWat แบบ "message ล้วน 100%" โดย
    ไม่แตะ focus/เมาส์/คีย์บอร์ดของระบบเลยแม้แต่ครั้งเดียว — CropWat อยู่หลัง
    หน้าต่างอื่นได้ ผู้ใช้ทำงานอื่นบนเครื่องไปพร้อมกันได้ (ห้าม minimize CropWat
    เท่านั้น เพราะ PrintWindow ถ่ายภาพหน้าต่างที่ยุบอยู่ไม่ได้) — 4 จุดที่ต่างจาก
    โหมดปกติ:
      1. สลับโมดูล: ส่ง WM_MDIACTIVATE ตรงไปที่ MDIClient (สถานะ active ของ MDI
         child เป็นเรื่องภายในโปรแกรม ไม่เกี่ยวกับ foreground ของระบบ)
      2. เรียกเมนู: หา id ของ menu item แล้ว post WM_COMMAND ตรงไปที่หน้าต่างหลัก
         (menu_select ของ pywinauto จะ set_focus ก่อนเสมอ ใช้ไม่ได้)
      3. ยืนยันค่าวันปลูก: ส่ง CM_EXIT (message ภายในของ Delphi VCL, 0xB011) ให้
         ช่องรัน logic "ออกจากช่อง" (validate+commit) แทนการกด Tab จริง
      4. Screenshot: PrintWindow (PW_RENDERFULLCONTENT) ถ่ายจากตัวหน้าต่างตรงๆ
         แม้ถูกบังอยู่ — ไม่ใช่ถ่ายจากพิกเซลบนจอ (ข้อดีพลอยได้: overlay ไม่ติดมา
         ในภาพ ไม่ต้องซ่อน)"""

    def __init__(self, background_mode: bool = False) -> None:
        self.app: Optional[Application] = None
        self.main_window = None
        self.background_mode = background_mode
        self._watcher_stop: Optional[threading.Event] = None
        # รูปแบบวันที่ของช่องวันปลูกบนเครื่องนี้ (ตรวจจับครั้งแรกที่ตั้งวันปลูก
        # แล้วจำไว้) — "dm"=วัน/เดือน "md"=เดือน/วัน ต่างกันตาม Region ของ Windows
        self._date_order: Optional[str] = None
        self._date_sep: str = "/"

    # ------------------------------------------------------------------
    # Step 0: ต่อเข้ากับ CropWat ที่เปิดอยู่แล้ว (ผู้ใช้เปิดโปรแกรมเองก่อนหน้านี้)
    # ------------------------------------------------------------------
    def connect(self) -> None:
        missing = controls.require_configured()
        if missing:
            raise ControlsNotConfiguredError(
                "cropwat_controls.py ยังกรอกไม่ครบ: " + ", ".join(missing)
            )
        try:
            self.app = Application(backend=controls.PYWINAUTO_BACKEND).connect(
                title_re=controls.MAIN_WINDOW_TITLE_RE,
                class_name=controls.MAIN_WINDOW_CLASS_NAME,
            )
            self.main_window = self.app.window(
                title_re=controls.MAIN_WINDOW_TITLE_RE,
                class_name=controls.MAIN_WINDOW_CLASS_NAME,
            )
            self.main_window.wait("exists enabled visible ready", timeout=10)
        except (ElementNotFoundError, PywinautoTimeoutError) as exc:
            raise CropWatNotRunningError(
                f"หา CropWat ไม่เจอ (title_re={controls.MAIN_WINDOW_TITLE_RE!r}) "
                "ตรวจสอบว่าเปิดโปรแกรมค้างไว้อยู่หรือยัง"
            ) from exc
        except ElementAmbiguousError as exc:
            raise CropWatNotRunningError(
                f"เจอหน้าต่างที่ตรงกับ title_re={controls.MAIN_WINDOW_TITLE_RE!r} + "
                f"class_name={controls.MAIN_WINDOW_CLASS_NAME!r} มากกว่า 1 หน้าต่าง "
                "(อาจมี CropWat เปิดค้างมากกว่า 1 session พร้อมกัน — ปิดให้เหลือ session เดียว)"
            ) from exc

    def _require_connected(self) -> None:
        if self.app is None or self.main_window is None:
            raise CropWatNotRunningError("ยังไม่ได้ connect() เข้ากับ CropWat")

    def _focus_mdi_child(self, class_name: str):
        """หา MDI child window ด้วย class_name (คงที่ไม่ว่าจะโหลดไฟล์ไหนอยู่) แล้ว
        ดึงขึ้นมา active — ต้องทำก่อนเรียกเมนู File->Open/File->Print เสมอ เพราะ
        เมนูพวกนี้ทำงานกับ MDI child ที่ active อยู่ตอนนั้น

        สำคัญ: ต้องระบุ top_level_only=False เสมอ — ยืนยันจากการทดสอบจริงแล้วว่า
        Application.window() แบบ default (top_level_only=True) หา MDI child ไม่
        เจอเลย เพราะ MDI child (เช่น TCropForm) เป็น child window ของ MDIClient
        ไม่ใช่ top-level window ของ process ทั้งที่หน้าตาดูเหมือนหน้าต่างแยก

        สำคัญ (v0.1.10): เช็คจำนวนหน้าต่างที่ตรงกับ class_name ก่อนเสมอ แทนที่จะ
        ปล่อยให้ pywinauto โยน ElementAmbiguousError ดิบๆ ออกไปตรงๆ (ข้อความเดิม
        "There are 2 elements that match..." ไม่บอกอะไรเลยว่าอันไหนคืออันไหน) —
        ยืนยันจากผู้ใช้แล้วว่าเจอ TDayEToPMForm ซ้ำ 2 หน้าต่าง คาดว่าเกิดจาก
        File->Open ของโมดูล Climate/Rain ไม่ได้แทนที่ไฟล์ในหน้าต่างเดิม แต่สร้าง
        MDI child ใหม่ทุกครั้งที่เปิดไฟล์ (ต่างจาก Crop/Soil ที่ยืนยันแล้วว่าเป็น
        หน้าต่างเดียวคงที่) — ถ้าจริง จะพอกซ้ำขึ้นเรื่อยๆ ทุกปีที่รัน ดังนั้นถ้าเจอ
        มากกว่า 1 หน้าต่าง ให้แจ้ง title ของทุกหน้าต่างที่ซ้ำ (มักมีชื่อไฟล์ที่โหลด
        อยู่ในนั้น) เพื่อให้ผู้ใช้ไปดูเมนู Window ใน CropWat แล้วปิดส่วนเกินเอง
        แทนที่จะเดาว่าอันไหนถูก"""
        handles = find_windows(
            class_name=class_name, top_level_only=False, process=self.app.process
        )
        if len(handles) > 1:
            titles = [win32gui.GetWindowText(h) or "(ไม่มีชื่อ)" for h in handles]
            raise DuplicateWindowError(
                f"เจอหน้าต่างโมดูลนี้ ({class_name}) เปิดค้างพร้อมกัน {len(handles)} "
                f"หน้าต่าง: {', '.join(titles)} — ระบุไม่ได้ว่าอันไหนถูกต้อง กรุณาเปิด "
                "เมนู Window ใน CropWat แล้วปิดหน้าต่างที่ซ้ำให้เหลือแค่บานเดียวก่อนรันใหม่"
            )

        window = self.app.window(class_name=class_name, top_level_only=False)
        window.wait("exists enabled visible ready", timeout=10)
        if self.background_mode:
            # สั่ง MDI activate ผ่าน message ตรงไปที่ MDIClient — ไม่แตะ foreground
            # ของระบบเลย (สถานะ "MDI child ไหน active" เป็นเรื่องภายในโปรแกรม)
            WM_MDIACTIVATE = 0x0222
            mdiclient = self.main_window.child_window(class_name="MDIClient")
            win32gui.SendMessage(mdiclient.handle, WM_MDIACTIVATE, window.handle, 0)
        else:
            window.set_focus()
        # เผื่อเวลาให้ CropWat ประมวลผลการเปลี่ยนหน้าต่าง active ก่อน — เจอจริงว่าถ้า
        # ยิงคำสั่งเมนูต่อทันทีโดยไม่รอเลย บางครั้ง CropWat ยังทำงานกับหน้าต่างเดิม
        # ที่ active อยู่ก่อนหน้า ไม่ใช่ตัวที่เพิ่งสั่ง activate ไป
        time.sleep(0.2)
        return window

    def _invoke_menu(self, menu_path: str) -> None:
        """เรียกเมนูของหน้าต่างหลัก — โหมดปกติใช้ menu_select ของ pywinauto (ซึ่ง
        set_focus ดึงหน้าต่างขึ้นมาก่อนเสมอ) โหมดเบื้องหลังหา id ของ menu item จาก
        โครงสร้างเมนู (อ่านได้โดยไม่ต้อง focus) แล้ว post WM_COMMAND ตรงไปที่
        หน้าต่างหลักแบบเดียวกับที่ Windows ส่งให้ตอนผู้ใช้คลิกเมนูจริง"""
        if not self.background_mode:
            self.main_window.menu_select(menu_path)
            return
        item = self.main_window.menu().get_menu_path(menu_path)[-1]
        if not item.is_enabled():
            raise CropWatReportedError(f"เมนู {menu_path!r} ถูก disable อยู่ (สถานะโปรแกรมยังไม่พร้อม)")
        # เมธอดชื่อ item_id() (ยืนยันจาก source ของ pywinauto — ไม่ใช่ .id())
        win32gui.PostMessage(self.main_window.handle, win32con.WM_COMMAND, item.item_id(), 0)

    def start_background_watcher(self) -> None:
        """โหมดเบื้องหลัง (v0.5.4): thread เฝ้ายามที่สแกน "ทุก" หน้าต่าง top-level
        ของ process CropWat ที่ไม่ใช่หน้าต่างหลัก แล้วย้ายออกนอกจอทันทีที่โผล่
        (ตรวจทุก 40ms) — จับได้หมดรวมถึงหน้าต่างที่ CropWat เด้งเอง-หายเองโดยเรา
        ไม่ได้กรอกอะไร เช่น "Printing progress" ระหว่าง print (ยืนยันจาก screenshot
        ผู้ใช้ว่าคือตัวที่ยังกระพริบหลัง v0.5.3 — การย้ายเฉพาะ dialog ที่เรารู้จัก
        ตอนจะกรอกไม่ครอบคลุมตัวนี้) — เรียกตอนเริ่มรัน และต้อง stop ตอนจบรันเสมอ
        ไม่งั้นถ้าผู้ใช้กลับมาใช้ CropWat เองต่อ dialog ของเขาจะโดนเหวี่ยงหนีจอไปด้วย"""
        if not self.background_mode or self._watcher_stop is not None:
            return
        self._require_connected()
        stop_event = threading.Event()
        self._watcher_stop = stop_event
        pid = self.app.process
        main_hwnd = self.main_window.handle
        flags = win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE

        def _fling_offscreen(hwnd: int) -> None:
            """v0.5.6: "ย้ายออกนอกจอ" อย่างเดียวเอาไม่อยู่ — ยืนยันจากผู้ใช้ว่า
            Printing progress ยังโผล่ทั้งที่ hook ทำงาน: ฟอร์ม progress ของ Delphi
            จัดตำแหน่งตัวเองกลับกลางจอระหว่างทำงาน เหวี่ยงไปก็เด้งกลับ — เพิ่มไม้
            ตาย: ตั้ง WS_EX_LAYERED + alpha 0 ให้หน้าต่าง "โปร่งใสสนิท" ซึ่งติด
            ถาวรกับ style ของหน้าต่าง ต่อให้มันย้ายกลับมากลางจอ/โชว์ซ้ำกี่รอบ
            ก็มองไม่เห็นอยู่ดี (ยังย้ายออกนอกจอควบด้วยกันพลาด)"""
            try:
                if hwnd == main_hwnd or not win32gui.IsWindow(hwnd):
                    return
                GWL_EXSTYLE = -20
                WS_EX_LAYERED = 0x00080000
                LWA_ALPHA = 0x2
                ex_style = win32gui.GetWindowLong(hwnd, GWL_EXSTYLE)
                if not ex_style & WS_EX_LAYERED:
                    win32gui.SetWindowLong(hwnd, GWL_EXSTYLE, ex_style | WS_EX_LAYERED)
                win32gui.SetLayeredWindowAttributes(hwnd, 0, 0, LWA_ALPHA)
                if win32gui.GetWindowRect(hwnd)[0] > -20000:
                    win32gui.SetWindowPos(hwnd, 0, -32000, -32000, 0, 0, flags)
            except Exception:  # noqa: BLE001 -- หน้าต่างตายไประหว่างจัดการ = ปกติ
                pass

        def _watch() -> None:
            """v0.5.5: เปลี่ยนจาก polling ทุก 40ms (ช้าไป — ตายังจับแวบ 40-100ms
            ได้ ยืนยันจากผู้ใช้ว่า "Printing progress" ยังโผล่) เป็น WinEvent hook
            แบบ event-driven: Windows เรียก callback เราทันทีที่หน้าต่างของ
            process CropWat "ถูกสร้าง" (EVENT_OBJECT_CREATE — ก่อนวาดตัวเองเสร็จ)
            → เหวี่ยงออกนอกจอตั้งแต่ยังไม่ทันปรากฏบนจอจริง
            หมายเหตุ: hook ต้องลงทะเบียน + วน message loop ใน thread เดียวกัน
            และต้องเก็บ reference ของ callback ไว้กัน GC เก็บกลางอากาศ"""
            import ctypes
            from ctypes import wintypes

            user32 = ctypes.windll.user32
            EVENT_OBJECT_CREATE = 0x8000
            EVENT_OBJECT_SHOW = 0x8002
            WINEVENT_OUTOFCONTEXT = 0x0
            OBJID_WINDOW = 0
            GA_ROOT = 2

            WinEventProc = ctypes.WINFUNCTYPE(
                None, wintypes.HANDLE, wintypes.DWORD, wintypes.HWND,
                wintypes.LONG, wintypes.LONG, wintypes.DWORD, wintypes.DWORD,
            )

            def _on_event(_hook, _event, hwnd, id_object, _id_child, _tid, _time):
                if not hwnd or id_object != OBJID_WINDOW:
                    return
                # เฉพาะหน้าต่าง top-level (ลูกๆ ข้างในไม่เกี่ยว)
                if user32.GetAncestor(hwnd, GA_ROOT) != hwnd:
                    return
                _fling_offscreen(hwnd)

            callback = WinEventProc(_on_event)
            # ช่วง CREATE(0x8000)..SHOW(0x8002) + กรองเฉพาะ process ของ CropWat
            hook = user32.SetWinEventHook(
                EVENT_OBJECT_CREATE, EVENT_OBJECT_SHOW, 0, callback,
                pid, 0, WINEVENT_OUTOFCONTEXT,
            )
            if not hook:
                logger.warning("ตั้ง WinEvent hook ไม่สำเร็จ — ใช้ polling สำรองอย่างเดียว")

            msg = wintypes.MSG()
            poll_countdown = 0
            while not stop_event.is_set():
                # hook callback ถูกส่งผ่าน message queue ของ thread นี้ — ต้องปั๊มเสมอ
                while user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
                    user32.TranslateMessage(ctypes.byref(msg))
                    user32.DispatchMessageW(ctypes.byref(msg))
                # polling สำรองทุก ~200ms เผื่อ event หลุด (เช่นหน้าต่างที่โผล่ก่อน hook ทัน)
                poll_countdown -= 1
                if poll_countdown <= 0:
                    poll_countdown = 20
                    try:
                        for hwnd in find_windows(process=pid, top_level_only=True):
                            _fling_offscreen(hwnd)
                    except Exception:  # noqa: BLE001
                        pass
                time.sleep(0.01)

            if hook:
                user32.UnhookWinEvent(hook)
            del callback  # ปลด reference หลัง unhook แล้วเท่านั้น

        threading.Thread(target=_watch, daemon=True, name="transient-window-watcher").start()
        logger.info("เริ่ม watcher (WinEvent hook) เฝ้าเหวี่ยงหน้าต่างชั่วคราวออกนอกจอ")

    def stop_background_watcher(self) -> None:
        if self._watcher_stop is not None:
            self._watcher_stop.set()
            self._watcher_stop = None
            logger.info("หยุด watcher หน้าต่างชั่วคราวแล้ว")

    def _hide_dialog_offscreen(self, dialog) -> None:
        """โหมดเบื้องหลัง (v0.5.3): ย้าย dialog ที่ CropWat เด้งขึ้นมา (Open/Save/
        Print/prompt) ออกไปนอกจอทันทีที่เจอ — dialog พวกนี้เป็นหน้าต่าง top-level
        แยกจากหน้าต่างหลัก จึงโผล่บนจอทั้งที่ CropWat ถูกบังอยู่ (ต้นเหตุอาการ
        "หน้าต่างเล็กกระพริบ" ที่ผู้ใช้เห็นระหว่างทำงานอื่น) — การกรอก/กดปุ่มของ
        เราเป็น message ล้วนอยู่แล้ว dialog อยู่นอกจอก็ทำงานได้ปกติทุกอย่าง"""
        if not self.background_mode:
            return
        try:
            flags = win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE
            win32gui.SetWindowPos(dialog.handle, 0, -32000, -32000, 0, 0, flags)
        except Exception:  # noqa: BLE001 -- ย้ายไม่ได้ก็แค่เห็นกระพริบ ไม่ใช่เหตุให้ล้ม
            logger.debug("ย้าย dialog ออกนอกจอไม่สำเร็จ", exc_info=True)

    def _ensure_module_window(self, class_name: str, new_menu_path: Optional[str], module_label: str):
        """คืนหน้าต่างโมดูล (focused) — ถ้ายังไม่มีเลย (CropWat เพิ่งเปิดมาเปล่าๆ)
        สั่ง File->New สร้างฟอร์ม "เปล่า" ก่อนเพื่อให้ MDI child ของโมดูลนั้นโผล่
        ขึ้นมา แล้วผู้เรียกค่อย File->Open ไฟล์จริงทับ — จำเป็นเพราะแถบไอคอนซ้าย
        (Climate/Rain/Crop/Soil) เป็นปุ่มวาดเองไม่มี HWND สั่งคลิกแบบ message
        ไม่ได้ File->New คือทางเดียวที่สร้างหน้าต่างโมดูลผ่านเมนูได้จริง"""
        handles = find_windows(
            class_name=class_name, top_level_only=False, process=self.app.process
        )
        if not handles:
            if not new_menu_path:
                raise CropWatNotRunningError(
                    f"ยังไม่มีหน้าต่าง {module_label} ใน CropWat และไม่ได้ตั้งเมนู "
                    f"File->New สำหรับสร้าง — เปิดไฟล์ {module_label} เองก่อนหนึ่งครั้ง"
                )
            logger.info("%s: ยังไม่มีหน้าต่างโมดูล — สร้างฟอร์มเปล่าผ่าน %s", module_label, new_menu_path)
            self._invoke_menu(new_menu_path)
        return self._focus_mdi_child(class_name)

    def _close_stale_module_windows(
        self,
        class_name: str,
        keep_file_name: Optional[str],
        module_label: str,
        close_all_if_no_match: bool = True,
    ) -> bool:
        """ปิดหน้าต่างของโมดูล (class เดียวกัน) ทุกบานที่ "ไม่ได้" โหลดไฟล์เป้าหมาย
        อยู่ — ยืนยันจากการรันจริงข้ามปี (v0.2.1): File->Open ของ CropWat เปิด
        "หน้าต่างใหม่" แทนที่จะโหลดทับบานเดิม ทำให้หน้าต่างพอกซ้ำขึ้นทุกปีจน
        engine แยกไม่ออกว่าบานไหนคือบานปัจจุบัน (ต้นตอของ DuplicateWindowError
        ตอนข้ามปี) — ตรวจจากการเทียบชื่อไฟล์กับ title ของแต่ละบาน (title มี path
        ไฟล์ที่โหลดอยู่เสมอ) ปิดด้วย WM_CLOSE (message ล้วน ไม่แตะเมาส์) และตอบ
        No ให้ prompt ถามบันทึกที่อาจเด้งตอนปิด

        คืน True ถ้าเจอบานที่โหลดไฟล์เป้าหมายอยู่แล้ว (เก็บไว้ 1 บาน ปิดที่เหลือ)
        keep_file_name=None = ปิดทุกบาน / close_all_if_no_match=False = ถ้าไม่เจอ
        บานที่ตรง อย่าปิดอะไรเลย (ใช้ตอนเก็บกวาด "หลัง" เปิดไฟล์สำเร็จ — กันเคส
        title มีรูปแบบไม่คาดคิดแล้วเผลอปิดบานที่เพิ่งเปิดเสร็จทิ้ง)"""
        handles = find_windows(
            class_name=class_name, top_level_only=False, process=self.app.process
        )
        if not handles:
            return False
        titles = {h: (win32gui.GetWindowText(h) or "") for h in handles}

        keep_handle = None
        if keep_file_name:
            for handle, title in titles.items():
                if keep_file_name.lower() in title.lower():
                    keep_handle = handle
                    break
        if keep_handle is None and not close_all_if_no_match:
            return False

        for handle, title in titles.items():
            if handle == keep_handle:
                continue
            logger.info("%s: ปิดหน้าต่างเก่า/ซ้ำ '%s'", module_label, title or "(ไม่มีชื่อ)")
            win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)
            deadline = time.monotonic() + 5
            while win32gui.IsWindow(handle) and time.monotonic() < deadline:
                self._answer_no_to_save_prompt()
                time.sleep(0.15)
            if win32gui.IsWindow(handle):
                logger.warning("%s: ปิดหน้าต่าง '%s' ไม่สำเร็จภายใน 5 วินาที", module_label, title)
        return keep_handle is not None

    def open_module_file(
        self,
        class_name: str,
        new_menu_path: Optional[str],
        file_path: Path,
        dialog_title_re,
        filename_field,
        open_button,
        module_label: str,
    ) -> None:
        """เปิดไฟล์เข้าโมดูลหนึ่งๆ แบบครบวงจร: สร้างหน้าต่างโมดูลถ้ายังไม่มี →
        ข้ามถ้าไฟล์เป้าหมายเปิดอยู่แล้ว (เช็คจาก title ของหน้าต่างซึ่งมี path ไฟล์
        ที่โหลดอยู่ต่อท้ายเสมอ) → File->Open + จัดการ prompt "Save changes?" เอง

        นี่คือสิ่งที่ทำให้ผู้ใช้ไม่ต้องเปิด crop/soil เองก่อนรันอีกต่อไป (v0.2.0) —
        ข้อสรุปเดิมที่ว่า "เปิดไฟล์ทับแล้ว CropWat error ต้องให้ผู้ใช้เปิดเอง" แท้จริง
        คือ dialog ถาม "Save changes to current ... data ?" (Yes/No/Cancel) ที่โค้ด
        เก่าอ่านผิดว่าเป็น error — ตอนนี้ระบบตอบ No ให้เองแล้ว จึงเปิดทับได้ปกติ

        v0.2.1: เปลี่ยนเป็น "ปิดบานเก่าก่อน เปิดใหม่เสมอ" — File->Open ของ CropWat
        เปิดหน้าต่างใหม่แทนที่จะโหลดทับบานเดิม (ยืนยันจากการรันจริงข้ามปี) ถ้าไม่
        ปิดบานปีเก่าทิ้งก่อน หน้าต่างจะพอกซ้ำทุกปีจนระบบแยกไม่ออกว่าบานไหนจริง"""
        self._require_connected()
        file_path = Path(file_path)

        # ปิดทุกบานของโมดูลนี้ที่โหลด "ไฟล์อื่น" อยู่ (ไฟล์ปีเก่า/ฟอร์มเปล่าค้าง) —
        # ถ้าเจอบานที่โหลดไฟล์เป้าหมายอยู่แล้ว เก็บไว้และจบเลย ไม่ต้องเปิดซ้ำ
        if self._close_stale_module_windows(class_name, file_path.name, module_label):
            logger.info("%s: ไฟล์ %s เปิดอยู่แล้ว ไม่ต้องเปิดซ้ำ", module_label, file_path.name)
            self._focus_mdi_child(class_name)
            return

        self._ensure_module_window(class_name, new_menu_path, module_label)
        self._open_file_via_dialog(file_path, dialog_title_re, filename_field, open_button)
        self._raise_if_error_dialog(f"เปิดไฟล์ {module_label} {file_path}")

        # เก็บกวาดหลังเปิด: ถ้า File->Open สร้างบานใหม่ ฟอร์มเปล่าที่ใช้เบิกทาง
        # ยังค้างอยู่เป็นบานที่สอง — ปิดทิ้งโดยเก็บเฉพาะบานที่โหลดไฟล์เป้าหมาย
        # (close_all_if_no_match=False กันเผลอปิดบานที่เพิ่งเปิดถ้า title ไม่ตรงคาด)
        self._close_stale_module_windows(
            class_name, file_path.name, module_label, close_all_if_no_match=False
        )

    def _open_file_via_dialog(self, file_path: Path, dialog_title_re, filename_field, open_button) -> None:
        """ทำ flow เปิดไฟล์ผ่าน Windows-style open dialog ที่เด้งขึ้นมาเป็น
        หน้าต่างแยกต่างหาก (ไม่ใช่ลูกของ CROPWAT หลัก) — ใช้ร่วมกันทั้ง climate/rain

        สำคัญ: ยืนยัน "อ่านค่ากลับ" หลังพิมพ์ path ว่าช่อง filename มีข้อความตรงกับ
        ที่ตั้งใจพิมพ์จริงๆ ก่อนจะกดปุ่ม Open เสมอ — เจอจริงว่าบางครั้งกด Open ไป
        ทั้งที่ช่องยังว่างอยู่ (dialog ยังโหลดไม่เสร็จตอนพิมพ์ หรือ control ที่ set
        ข้อความไม่ใช่ตัวที่ถูกต้อง) ทำให้ CropWat error เพราะไม่มีไฟล์ให้เปิดจริง
        ถ้าอ่านกลับมาไม่ตรง จะลองพิมพ์ใหม่อีกครั้งก่อนจะ fail แบบมีข้อความชัดเจน
        แทนที่จะกด Open ไปทั้งที่รู้อยู่แล้วว่าข้อมูลผิด

        สำคัญ (v0.1.9): ใช้ set_edit_text()/click() ที่ส่ง message ตรงไปที่ handle
        ของ control เท่านั้น ไม่ใช้ click_input()/type_keys() ที่จำลองเมาส์/คีย์บอร์ด
        จริงตามพิกัดหน้าจอ — set_edit_text ใช้ WM_SETTEXT ซึ่งไม่ต้องคลิกโฟกัสก่อน
        เลยด้วยซ้ำ (เดิมมี field.click_input() ก่อนพิมพ์ทั้งที่ไม่จำเป็น และเป็นจุด
        เสี่ยง "คลิกหลุดเป้า" ถ้าหน้าต่างซ้อนผิดจังหวะ — เอาออก)"""
        if not file_path.exists():
            raise FileNotFoundError(f"ไม่พบไฟล์: {file_path}")

        self._invoke_menu("File->Open")

        # ยืนยันจาก screenshot จริงของผู้ใช้ (v0.1.12): ก่อน dialog เปิดไฟล์จะโผล่
        # CropWat อาจเด้งถาม "Save changes to current climate/rain data ?"
        # (Yes/No/Cancel) ถ้านับว่าข้อมูลในหน้าต่างเดิมถูกแก้ค้างอยู่ — เป็น modal
        # ที่บังไม่ให้ dialog เปิดไฟล์โผล่จนกว่าจะตอบ ต้องคอยเช็คแล้วตอบ No ให้
        # ระหว่างรอ ไม่งั้น flow ค้างจน timeout ทั้งที่ทุกอย่างปกติดี
        dialog = self.app.window(title_re=dialog_title_re)
        deadline = time.monotonic() + 10
        while not dialog.exists(timeout=0):
            self._answer_no_to_save_prompt()
            if time.monotonic() > deadline:
                raise StepTimeoutError(
                    f"dialog เปิดไฟล์ (title_re={dialog_title_re!r}) ไม่โผล่ภายใน 10 วินาที"
                )
            time.sleep(0.2)
        self._hide_dialog_offscreen(dialog)
        dialog.wait("exists enabled visible ready", timeout=10)
        time.sleep(0.3)  # เผื่อเวลาให้ dialog พร้อมรับ input จริงๆ ก่อนพิมพ์

        target = str(file_path)
        field = dialog[filename_field]
        for attempt in range(2):
            field.set_edit_text(target)
            time.sleep(0.2)
            actual = field.window_text()
            if actual.strip().strip('"') == target.strip().strip('"'):
                break
            if attempt == 0:
                continue  # ลองพิมพ์ซ้ำอีกครั้งก่อน fail
            raise StepTimeoutError(
                f"พิมพ์ path ไฟล์ในช่อง filename ไม่สำเร็จ (ตั้งใจ: {target!r}, "
                f"อ่านได้จริง: {actual!r}) — ไม่กด Open เพราะข้อมูลไม่ตรง"
            )

        dialog[open_button].click()

    # ------------------------------------------------------------------
    # Step 1a: เปิดไฟล์ Climate (.PED) ของปีนั้น
    # ------------------------------------------------------------------
    def open_climate_file(self, file_path: Path) -> None:
        cfg = controls.CLIMATE_SCREEN
        self.open_module_file(
            cfg.window_class_name,
            cfg.new_menu_path,
            Path(file_path),
            cfg.file_dialog_title_re,
            cfg.file_dialog_filename_field,
            cfg.file_dialog_open_button,
            "climate",
        )

    # ------------------------------------------------------------------
    # Step 1b: เปิดไฟล์ Rain (.CRD) ของปีนั้น
    # ------------------------------------------------------------------
    def open_rain_file(self, file_path: Path) -> None:
        cfg = controls.RAIN_SCREEN
        self.open_module_file(
            cfg.window_class_name,
            cfg.new_menu_path,
            Path(file_path),
            cfg.file_dialog_title_re,
            cfg.file_dialog_filename_field,
            cfg.file_dialog_open_button,
            "rain",
        )

    # ------------------------------------------------------------------
    # เปิดไฟล์ Crop/Soil ให้เองอัตโนมัติ (v0.2.0 — ผู้ใช้ไม่ต้องเปิดเองอีกต่อไป)
    # เรียกแค่ "ครั้งเดียวต่อ batch" ก่อนเริ่มวนหลายปี เพราะไฟล์คงที่ตลอด — ถ้า
    # เปิดอยู่แล้วและ title ตรงกับไฟล์เป้าหมาย open_module_file จะข้ามให้เอง
    # ------------------------------------------------------------------
    def ensure_crop_soil_open(self, crop_file: Path, soil_file: Path) -> None:
        cfg = controls.CROP_SCREEN
        self.open_module_file(
            cfg.window_class_name,
            cfg.new_menu_path,
            Path(crop_file),
            cfg.file_dialog_title_re,
            cfg.file_dialog_filename_field,
            cfg.file_dialog_open_button,
            "crop",
        )
        scfg = controls.SOIL_SCREEN
        self.open_module_file(
            scfg.window_class_name,
            scfg.new_menu_path,
            Path(soil_file),
            scfg.file_dialog_title_re,
            scfg.file_dialog_filename_field,
            scfg.file_dialog_open_button,
            "soil",
        )

    # ------------------------------------------------------------------
    # Step 2: ตั้งวันปลูกใน module Crop (crop file เองคงที่ทุกปี ไม่ต้องเปิดใหม่)
    # ------------------------------------------------------------------
    def _commit_field(self, field, cfg) -> None:
        """ยืนยันค่าในช่อง — โหมดเบื้องหลังส่ง CM_EXIT (0xB011, message ภายในของ
        Delphi VCL ที่สั่งรัน logic "ออกจากช่อง": validate + commit + OnExit)
        โหมดปกติกด Tab จริง (จำลองคีย์บอร์ด ไม่ขยับเมาส์)"""
        if self.background_mode:
            CM_EXIT = 0xB011
            win32gui.SendMessage(field.handle, CM_EXIT, 0, 0)
        elif cfg.confirm_key:
            field.type_keys(cfg.confirm_key)

    @staticmethod
    def _parse_field_date(text: str) -> Optional[tuple[int, str, int]]:
        """แยกข้อความในช่องวันที่เป็น (เลขหน้า, ตัวคั่น, เลขหลัง) — ไม่ตีความว่า
        อันไหนวัน/เดือน (ขึ้นกับ Region ของเครื่อง)"""
        m = re.match(r"^\s*(\d{1,2})(\D)(\d{1,2})\s*$", text or "")
        if not m:
            return None
        return int(m.group(1)), m.group(2), int(m.group(3))

    def set_planting_date(self, planting_date: date) -> None:
        """v0.5.6 — บทเรียนจากเครื่องเพื่อนของผู้ใช้ (บั๊กอันตรายสุดที่เคยเจอ):
        ช่องวันปลูกอิง "รูปแบบวันที่ของ Windows เครื่องนั้น" — บางเครื่องเป็น
        เดือน/วัน หรือใช้ตัวคั่นอื่น ถ้าส่งรูปแบบผิด CropWat ปัดทิ้งแล้วรีเซ็ตเป็น
        "วันนี้" แบบเงียบสนิท ทุกวันปลูกเลยกลายเป็นวันเดียวกันหมดโดยไม่มี error —
        แก้ 2 ชั้น: (1) ตรวจจับลำดับวัน/เดือน + ตัวคั่นจากค่าที่อยู่ในช่องเอง
        (2) หลัง commit ต้องอ่านกลับมาตรวจว่าค่า "ติดจริง" ไม่ถูกรีเซ็ต — ผิดให้
        ลองสลับลำดับหนึ่งครั้ง ยังผิดอีกให้ fail ดังๆ พร้อมบอกวิธีแก้ที่ Region"""
        self._require_connected()
        cfg = controls.CROP_SCREEN
        crop_window = self._focus_mdi_child(cfg.window_class_name)
        field = crop_window.child_window(class_name=cfg.planting_date_field_class_name)

        # ตรวจ "ลำดับ" และ "ตัวคั่น" จากค่าปัจจุบันในช่อง (ครั้งแรกของ batch ช่องจะ
        # เป็นวันนี้ที่ CropWat ตั้งเอง — ใช้เทียบกับวันนี้จริงเพื่อรู้ลำดับ) แล้วจำไว้
        if self._date_order is None:
            current = self._parse_field_date(field.window_text())
            order, sep = "dm", "/"
            if current:
                a, sep, b = current
                today = date.today()
                if (a, b) == (today.day, today.month) and a != b:
                    order = "dm"
                elif (a, b) == (today.month, today.day) and a != b:
                    order = "md"
                elif a > 12:
                    order = "dm"
                elif b > 12:
                    order = "md"
            self._date_order, self._date_sep = order, sep
            logger.info("รูปแบบวันที่ของช่องวันปลูกเครื่องนี้: %s (คั่นด้วย %r)", order, sep)

        def _text_for(order: str) -> str:
            d, m = planting_date.day, planting_date.month
            first, second = (d, m) if order == "dm" else (m, d)
            return f"{first:02d}{self._date_sep}{second:02d}"

        def _type_into_field(text: str) -> None:
            """พิมพ์ "ทีละตัวอักษร" ผ่าน WM_CHAR ตรงเข้า control — บทเรียนสำคัญ
            (ยืนยันจากไฟล์ output จริงทั้งของผู้ใช้และเพื่อน): WM_SETTEXT เปลี่ยน
            แค่ข้อความที่แสดง แต่ Delphi ไม่ตั้ง Modified flag → ตอนออกจากช่อง
            CropWat ไม่อัปเดตค่าวันปลูกใน model ภายใน → ทุกไฟล์คำนวณด้วยวันปลูก
            เก่าค้างเดิมทั้งที่ช่องโชว์ค่าใหม่ — WM_CHAR เดินผ่านเส้นทางเดียวกับ
            การพิมพ์จริงของผู้ใช้ (mask/Modified/OnChange ครบ) โดยไม่ต้องแตะ
            คีย์บอร์ดจริงของระบบ"""
            hwnd = field.handle
            # เลือกทั้งหมดก่อน ให้ตัวแรกที่พิมพ์แทนที่ค่าเดิมทั้งช่อง
            win32gui.SendMessage(hwnd, win32con.EM_SETSEL, 0, -1)
            for ch in text:
                win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord(ch), 0)
                time.sleep(0.02)

        def _read_model_planting_date() -> Optional[tuple[int, int]]:
            """อ่านค่าวันปลูกที่ CropWat "ใช้จริง" จากแถบสถานะล่างของหน้าต่างหลัก
            (แผง TPanel ที่ข้อความเป็นรูปวันที่ — ยืนยันจาก inspect dump ว่ามีจริง)
            — ต่างจากข้อความในช่องกรอกที่อาจโชว์ค่าที่ model ไม่รับรู้ คืน (a, b)
            ตามที่แสดง (ยังไม่ตีความวัน/เดือน)"""
            try:
                for panel in self.main_window.descendants(class_name="TPanel"):
                    parsed = self._parse_field_date(panel.window_text())
                    if parsed:
                        return parsed[0], parsed[2]
            except Exception:  # noqa: BLE001 -- อ่านไม่ได้ = ใช้ fallback ช่องกรอกแทน
                pass
            return None

        orders_to_try = [self._date_order, "md" if self._date_order == "dm" else "dm"]
        last_seen = ""
        for order in orders_to_try:
            text = _text_for(order)
            _type_into_field(text)
            self._commit_field(field, cfg)
            self._raise_if_error_dialog(f"ตั้งวันปลูก {text}")

            expected = (
                (planting_date.day, planting_date.month)
                if order == "dm"
                else (planting_date.month, planting_date.day)
            )
            # ตรวจจาก "แถบสถานะ" (ค่า model จริง) ก่อน — fallback เป็นช่องกรอก
            # วนอ่านซ้ำสั้นๆ เผื่อแถบสถานะอัปเดตช้ากว่า commit เสี้ยววินาที
            # (กันตัดสิน "ไม่ติด" ผิดๆ แล้วไปลองสลับลำดับโดยไม่จำเป็น)
            committed = False
            for _ in range(4):
                model_value = _read_model_planting_date()
                if model_value is not None:
                    committed = model_value == expected
                    last_seen = f"status bar={model_value}"
                else:
                    after = self._parse_field_date(field.window_text())
                    committed = bool(after) and (after[0], after[2]) == expected
                    last_seen = f"field={field.window_text()!r}"
                if committed:
                    break
                time.sleep(0.15)
            if committed:
                self._date_order = order  # จำลำดับที่ใช้ได้จริงไว้ใช้รอบถัดไป
                return
            logger.warning(
                "วันปลูกไม่ติดที่ model (ส่ง %r, เห็น %s) — ลองสลับลำดับวัน/เดือน", text, last_seen
            )

        raise CropWatReportedError(
            f"CropWat ไม่รับวันปลูก {planting_date:%d/%m/%Y} เข้า model จริง (ลองทั้ง "
            f"วัน/เดือน และ เดือน/วัน แล้ว, เห็น {last_seen}) — ตรวจรูปแบบวันที่ของ "
            "Windows: Settings → Time & language → Language & region → Regional format "
            "→ Short date = dd/MM/yyyy แล้วปิด-เปิด CropWat และรันใหม่"
        )

    # ------------------------------------------------------------------
    # Step 3: สั่งคำนวณ Crop Water Requirements — ยืนยันจาก inspect_menu.py แล้ว
    # ว่าเป็น 2 ขั้นตอนแยกกัน ต้องรันอันนี้ก่อนเสมอ (Irrigation Scheduling ต้องพึ่ง
    # ผลนี้)
    # ------------------------------------------------------------------
    def calculate(self) -> None:
        """หมายเหตุ (v0.1.11 — แก้บั๊กช้าตัวแม่): เดิมมี polling loop ที่วนเช็ค
        error จนครบ calculate_timeout_seconds (30 วิ) เต็มๆ โดย "ไม่มีเงื่อนไขออก
        เมื่อสำเร็จ" เลย — เสียเวลาฟรี 30 วิทุกครั้งที่คำนวณ ทั้งที่ CropWat คำนวณ
        ข้อมูลรายวัน 12 เดือนเสร็จแทบทันที (Delphi ประมวลผลใน UI thread แบบ
        synchronous — พอคำสั่งเมนูถูกประมวลผลเสร็จ ผลคำนวณก็เสร็จแล้ว ถ้าพังจะ
        เด้ง error dialog ขึ้นมาเลยทันที) เช็ค error หนึ่งรอบสั้นๆ ก็พอ"""
        self._require_connected()
        cfg = controls.CALCULATE
        self._invoke_menu(cfg.crop_water_requirements_menu_path)
        self._raise_if_error_dialog("คำนวณ Crop Water Requirements")

    # ------------------------------------------------------------------
    # Step 4: สั่งคำนวณ Irrigation Scheduling — นี่คือสิ่งที่ "เปิดหน้า Irrigation
    # Schedule" จริงๆ ในเมนู ไม่มีคำสั่ง "เปิดหน้า" แยกต่างหาก การสั่งคำนวณนี้
    # จะทำให้หน้าต่างผลลัพธ์ (class TCropScheduleform) โผล่ขึ้นมาเป็นผลพลอยได้เลย
    # ------------------------------------------------------------------
    def open_irrigation_schedule(self) -> None:
        """หมายเหตุ (v0.1.11): เอา polling loop 30 วิเต็มออกด้วยเหตุผลเดียวกับ
        calculate() — สำหรับขั้นนี้มี "สัญญาณเสร็จ" ที่ดีกว่าด้วยซ้ำ: หน้าต่างผลลัพธ์
        (TCropScheduleform) ต้องโผล่ขึ้นมา — _focus_mdi_child ด้านล่างรอหน้าต่างนี้
        อยู่แล้ว (timeout 10 วิ) เป็นการยืนยันว่าคำนวณสำเร็จจริงในตัวเอง"""
        self._require_connected()
        cfg = controls.CALCULATE
        self._invoke_menu(cfg.irrigation_scheduling_menu_path)
        self._raise_if_error_dialog("คำนวณ Irrigation Scheduling")

        sched_cfg = controls.IRRIGATION_SCHEDULE
        if sched_cfg.window_class_name:
            self._focus_mdi_child(sched_cfg.window_class_name)

    # ------------------------------------------------------------------
    # Step 5: กด Print (print-to-file) -> ได้ไฟล์ CSV/txt
    #
    # ยืนยันจาก inspect_cropwat.py แล้วว่าเป็น 2 จอ: จอตัวเลือก (TPrintForm — เลือก
    # ASCII file / ติ๊ก comma / เลือก checkbox "Irrigation schedule") กด OK แล้ว
    # ค่อยไปเจอ dialog เลือก path ปลายทางจริงอีกที (คาดว่าเป็น Windows common
    # file dialog เหมือน Open — รอยืนยัน field ที่แน่นอน)
    # ------------------------------------------------------------------
    def export_results(self, year: int, planting_date: date, export_dir: Path) -> Path:
        self._require_connected()
        cfg = controls.IRRIGATION_SCHEDULE
        export_dir = Path(export_dir)
        export_dir.mkdir(parents=True, exist_ok=True)
        # ตั้งชื่อไฟล์รวมวันปลูกที่ทดลองด้วย เพราะ 1 ปีมีได้หลายไฟล์ (1 ไฟล์ต่อ
        # 1 วันปลูกที่ทดลอง)
        target_file = export_dir / f"{year}_{planting_date:%m%d}.txt"
        # ลบไฟล์เก่าทิ้งก่อนเสมอถ้ามีค้างจากรอบก่อน (เช่น กดรันปีเดิมซ้ำ) — กัน 2
        # ปัญหาพร้อมกัน: (1) Save As จะเด้ง prompt ยืนยันเขียนทับที่โค้ดไม่ได้กดให้
        # ทำให้ flow ค้าง (2) _wait_for_file ด้านล่างจะเจอไฟล์เก่าแล้วรายงาน
        # "สำเร็จ" ทั้งที่รอบนี้ยังไม่ได้ print อะไรออกมาจริงเลย
        target_file.unlink(missing_ok=True)

        schedule_window = self._focus_mdi_child(cfg.window_class_name)
        # ยืนยันจากผู้ใช้แล้ว: ต้องสลับไปดู "Daily soil moisture balance" ก่อนพิมพ์เสมอ
        # click() (ไม่ใช่ click_input()) — ส่ง message ตรงไปที่ handle ของ radio
        # button เลย ไม่ต้องพึ่งพิกัดเมาส์จริงบนหน้าจอ (ดู bulk note ที่ _open_file_via_dialog)
        schedule_window[cfg.table_format_daily_soil_moisture_radio].click()

        self._invoke_menu(cfg.print_menu_path)

        options = self.app.window(title_re=cfg.print_options_dialog_title_re)
        options.wait("exists enabled visible ready", timeout=10)
        self._hide_dialog_offscreen(options)
        options[cfg.print_options_ascii_file_radio].click()
        commas_checkbox = options[cfg.print_options_use_commas_checkbox].wrapper_object()
        if not commas_checkbox.get_check_state():
            commas_checkbox.click()
        irrigation_checkbox = options[cfg.print_options_irrigation_schedule_checkbox].wrapper_object()
        if not irrigation_checkbox.get_check_state():
            irrigation_checkbox.click()
        options[cfg.print_options_ok_button].click()

        save_dialog = self.app.window(title_re=cfg.print_save_dialog_title_re)
        save_dialog.wait("exists enabled visible ready", timeout=10)
        self._hide_dialog_offscreen(save_dialog)
        save_dialog[cfg.print_save_dialog_filename_field].set_edit_text(str(target_file))
        save_dialog[cfg.print_save_dialog_save_button].click()

        self._raise_if_error_dialog(f"print ผลลัพธ์ปี {year} วันปลูก {planting_date:%d/%m}")

        if not self._wait_for_file(target_file, timeout=15):
            raise StepTimeoutError(
                f"รอไฟล์ print ปี {year} วันปลูก {planting_date:%d/%m} ไม่เจอ: {target_file}"
            )

        return target_file

    # ------------------------------------------------------------------
    # Step 6 (เสริม, เฉพาะวันปลูกที่ถูกเลือกให้ capture): ถ่ายภาพหน้าจอทั้งหน้าต่าง
    # CropWat 2 ภาพ — ยืนยันจากไฟล์ .docx ตัวอย่างจริงของผู้ใช้ (1,204 ภาพ):
    # 1) หน้าต่าง "Crop irrigation schedule" (ตาราง Daily soil moisture balance +
    #    Totals/Yield reductions ที่เพิ่งเลือกไว้ตอน export_results)
    # 2) หน้าต่าง "Irrigation scheduling graph" (กราฟ depletion)
    # ทั้งสองภาพ capture ทั้งหน้าต่างหลัก CropWat เสมอ (ไม่ใช่แค่ MDI child) เพื่อให้
    # เห็นชื่อไฟล์ climate/rain/crop/soil + วันปลูกที่ status bar ด้านล่างด้วย
    # (ต้องเห็นว่าใช้ไฟล์ตัวไหน วันที่เท่าไหร่ ตามที่ผู้ใช้ระบุไว้)
    # ------------------------------------------------------------------
    def capture_screenshots(
        self, year: int, planting_date: date, screenshot_dir: Path
    ) -> tuple[Path, Path]:
        self._require_connected()
        cfg = controls.IRRIGATION_SCHEDULE
        screenshot_dir = Path(screenshot_dir)
        screenshot_dir.mkdir(parents=True, exist_ok=True)
        stem = f"{year}_{planting_date:%m%d}"

        # โหมดปกติ: ซ่อน overlay ลอยชั่วคราวระหว่าง capture (ถ่ายจากพิกเซลจริงบนจอ
        # overlay จะติดมาในภาพ) — โหมดเบื้องหลังไม่ต้อง เพราะ PrintWindow ถ่ายจาก
        # ตัวหน้าต่าง CropWat ตรงๆ ไม่ใช่จากจอ
        pause = None
        if not self.background_mode:
            try:
                from overlay import capture_pause as pause
            except Exception:  # noqa: BLE001 -- overlay เป็น optional ไม่มีก็ capture ได้
                pass
        if pause is not None:
            pause.set()
            time.sleep(0.8)

        try:
            self._focus_mdi_child(cfg.window_class_name)
            schedule_path = screenshot_dir / f"{stem}_schedule.png"
            self._capture_main_window(schedule_path)

            # ปิดหน้าต่างกราฟเก่าทิ้งแล้วเปิดใหม่ผ่านเมนู "ทุกครั้ง" — ยืนยันจาก
            # การรันจริง (v0.2.1): หน้าต่างกราฟที่เปิดค้างไว้ไม่รีเฟรชตามผลคำนวณ
            # ใหม่ ถ้าเปิดครั้งเดียวแล้ว capture ซ้ำๆ จะได้ภาพกราฟของวันปลูกแรก
            # ทุกใบ — ต้องเปิดสดใหม่ให้ตรงกับผลคำนวณของวันปลูกปัจจุบันเสมอ
            self._close_stale_module_windows(
                cfg.graph_window_class_name, None, "irrigation schedule graph"
            )
            self._invoke_menu(cfg.graph_menu_path)
            self._focus_mdi_child(cfg.graph_window_class_name)
            graph_path = screenshot_dir / f"{stem}_graph.png"
            self._capture_main_window(graph_path)
        finally:
            if pause is not None:
                pause.clear()

        return schedule_path, graph_path

    def _capture_main_window(self, save_path: Path) -> None:
        """ถ่ายภาพหน้าต่างหลัก CropWat ทั้งบาน — โหมดปกติถ่ายจากพิกเซลบนจอ
        (ต้องเห็นหน้าต่าง) โหมดเบื้องหลังใช้ PrintWindow สั่งให้ตัวโปรแกรมวาด
        เนื้อหาของตัวเองลง bitmap โดยตรง ใช้ได้แม้หน้าต่างถูกบานอื่นบังมิดอยู่
        (แต่ใช้ไม่ได้ถ้า minimize — เนื้อหาไม่ถูกวาดเลย จึงห้ามยุบ CropWat)"""
        if not self.background_mode:
            self.main_window.capture_as_image().save(save_path)
            return

        import ctypes

        import win32ui
        from PIL import Image

        hwnd = self.main_window.handle
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width, height = right - left, bottom - top

        # แก้ภาพมีขอบดำบนจอที่ตั้ง display scaling 125%/150% (ยืนยันจากภาพจริง
        # ของผู้ใช้): CropWat เป็นโปรแกรมเก่าที่ไม่รู้จัก DPI — PrintWindow ให้มัน
        # วาดตัวเองที่ขนาด "logical" ของมันเอง ขณะที่ GetWindowRect (จาก process
        # เราที่ DPI-aware) คืนขนาด "physical" ที่ใหญ่กว่า → bitmap เหลือที่ว่าง
        # เป็นสีดำ — คำนวณขนาดที่มันจะวาดจริงจากอัตราส่วน DPI ของหน้าต่างเทียบ
        # กับจอ: unaware ได้ 96/จอ (เล็กลงพอดีเป๊ะ), aware ได้ 1:1 (เท่าเดิม)
        try:
            import ctypes as _ct

            win_dpi = _ct.windll.user32.GetDpiForWindow(hwnd) or 96
            hmon = _ct.windll.user32.MonitorFromWindow(hwnd, 2)  # MONITOR_DEFAULTTONEAREST
            mon_x, mon_y = _ct.c_uint(96), _ct.c_uint(96)
            _ct.windll.shcore.GetDpiForMonitor(hmon, 0, _ct.byref(mon_x), _ct.byref(mon_y))
            mon_dpi = mon_x.value or 96
            width = round(width * win_dpi / mon_dpi)
            height = round(height * win_dpi / mon_dpi)
        except Exception:  # noqa: BLE001 -- Windows เก่าไม่มี API นี้ → ใช้ขนาด physical เดิม
            pass

        hwnd_dc = win32gui.GetWindowDC(hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        try:
            bmp.CreateCompatibleBitmap(mfc_dc, width, height)
            save_dc.SelectObject(bmp)
            PW_RENDERFULLCONTENT = 2  # จำเป็นสำหรับหน้าต่างที่ถูกบัง (Win 8.1+)
            ok = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), PW_RENDERFULLCONTENT)
            if not ok:
                # บางหน้าต่าง/ไดรเวอร์ไม่รองรับ PrintWindow → ถอยไปถ่ายจากจอแทน
                # (ภาพอาจมีอย่างอื่นบัง แต่ดีกว่าไม่มีภาพเลย)
                logger.warning("PrintWindow ล้มเหลว — ถอยไปถ่ายจากหน้าจอแทน (ภาพอาจถูกบัง)")
                self.main_window.capture_as_image().save(save_path)
                return
            info = bmp.GetInfo()
            pixels = bmp.GetBitmapBits(True)
            img = Image.frombuffer(
                "RGB", (info["bmWidth"], info["bmHeight"]), pixels, "raw", "BGRX", 0, 1
            )
            img.save(save_path)
        finally:
            win32gui.DeleteObject(bmp.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwnd_dc)

    # ------------------------------------------------------------------
    # จัดการ prompt คำถาม Yes/No ของ CropWat (เช่น "Save changes to current
    # rain data ?" ตอนเปิดไฟล์ใหม่ทับหน้าต่างที่ข้อมูลถูกนับว่าแก้ค้างอยู่)
    # ------------------------------------------------------------------
    def _answer_no_to_save_prompt(self) -> bool:
        """เช็คครั้งเดียว (ไม่รอ) ว่ามี prompt คำถามแบบ Yes/No/Cancel เด้งอยู่ไหม
        ถ้ามีให้ตอบ No ทันที — ต้อง No เสมอ ห้าม Yes เด็ดขาด เพราะ Yes จะบันทึก
        state ชั่วคราวในหน้าต่างทับไฟล์ข้อมูลต้นทางของผู้ใช้ (ไฟล์ climate/rain
        ดิบที่ automation ไม่มีสิทธิ์ไปแก้) — "ความเปลี่ยนแปลง" ที่ CropWat ถามถึง
        เกิดจากการโหลด/คำนวณ eff. rain อัตโนมัติ ไม่ใช่สิ่งที่ผู้ใช้ตั้งใจแก้จริง"""
        try:
            prompt = self.app.window(title_re=r"Warning|Confirm")
            if not prompt.exists(timeout=0):
                return False
            no_button = prompt.child_window(title_re=r"&?No$")
            if not no_button.exists(timeout=0):
                return False
            self._hide_dialog_offscreen(prompt)
            no_button.click()
            logger.info("ตอบ No อัตโนมัติให้ prompt ถามบันทึกข้อมูล (ไม่บันทึกทับไฟล์ต้นทาง)")
            time.sleep(0.2)
            return True
        except (ElementNotFoundError, ElementAmbiguousError):
            return False

    # ------------------------------------------------------------------
    # ตรวจจับ error/warning dialog (ใช้ทั้งเช็ค inline และ poll หลัง calculate)
    # ------------------------------------------------------------------
    def _poll_error_dialog(self) -> Optional[str]:
        cfg = controls.ERROR_DIALOG
        try:
            dialog = self.app.window(title_re=cfg.title_re)
            if not dialog.exists(timeout=0):
                return None
            self._hide_dialog_offscreen(dialog)
            # แยก "คำถาม" ออกจาก "error" ก่อน: dialog ที่มีปุ่ม No (Yes/No/Cancel)
            # คือคำถามให้เลือก ไม่ใช่ error — เดิมโค้ดเหมารวมทุก title "Warning"
            # เป็น error แล้วพยายามกดปุ่ม OK ที่ไม่มีอยู่จริง ทำให้ล้มเหลวเงียบๆ
            # แล้วทิ้ง dialog ค้างบังทุกอย่างต่อจากนั้น (ยืนยันจาก screenshot ผู้ใช้)
            no_button = dialog.child_window(title_re=r"&?No$")
            if no_button.exists(timeout=0):
                no_button.click()
                logger.info("ตอบ No อัตโนมัติให้ prompt ถามบันทึกข้อมูล (ไม่ใช่ error)")
                return None
            if cfg.message_text_control:
                message = dialog[cfg.message_text_control].window_text()
            else:
                # ยืนยันจาก inspect แล้ว: TMessageForm ("Error") ไม่มี control
                # ข้อความแยกต่างหาก (มีแค่ปุ่ม OK เป็นลูก) ข้อความวาดตรงบนตัว
                # dialog เอง อ่านละเอียดด้วย child_window ไม่ได้ — ใช้ title ของ
                # dialog เองเป็น fallback แทน อย่างน้อยก็รู้ว่ามี error เกิดขึ้น
                message = f"CropWat แสดง dialog '{dialog.window_text()}' (อ่านข้อความละเอียดไม่ได้)"
            if cfg.dismiss_button:
                dialog[cfg.dismiss_button].click()
            return message
        except ElementNotFoundError:
            return None
        except ElementAmbiguousError:
            # เจอ dialog ที่ title ตรงกับ error/warning มากกว่า 1 อัน — ไม่กล้าเดา
            # กดปุ่มปิดมั่วๆ (เสี่ยงกดผิดตัว) แค่รายงานว่ามีปัญหาให้ candidate นี้
            # ถูก mark error ไว้ตรวจสอบเอง ดีกว่าปล่อยให้ exception หลุดออกไปทำให้
            # ทั้ง batch ล้ม
            return "เจอ dialog error/warning มากกว่า 1 อันพร้อมกัน (ไม่สามารถระบุได้ว่าอันไหน)"

    def _raise_if_error_dialog(self, context: str) -> None:
        """v0.5.7 — เร่งความเร็ว: เดิม sleep(poll_timeout) แบบตายตัวแล้วเช็ครอบเดียว
        จุดนี้ถูกเรียก 4 ครั้ง/วันปลูก คือต้นทุนคงที่ก้อนใหญ่สุดที่เหลือ (รอเปล่าๆ
        ~3 วิ/วันปลูกทั้งที่ปกติไม่มี error เลย) — เปลี่ยนเป็น poll ถี่ๆ แล้ว "ออก
        ทันทีที่เจอ error" (Delphi เด้ง dialog แทบทันทีที่ประมวลผลคำสั่งเมนูเสร็จ
        เพราะทำ synchronous ใน UI thread) ถ้าไม่มี error ก็รอแค่ window สั้นๆ แล้ว
        ไปต่อ — ทุกจุดที่เรียกยังมี "สัญญาณสำเร็จ" ตามหลังเป็นตาข่ายกันพลาดชั้นสอง
        (ค่าวันปลูกติดใน model จริง, หน้าต่างผลลัพธ์โผล่, ไฟล์ .txt ถูกเขียนออกมา)
        error ที่โผล่ช้ากว่า window นี้จะถูกจับที่จุดเช็คของ step ถัดไปอยู่ดี"""
        cfg = controls.ERROR_DIALOG
        deadline = time.monotonic() + cfg.poll_timeout_seconds
        while True:
            message = self._poll_error_dialog()
            if message:
                raise CropWatReportedError(f"CropWat แจ้ง error ระหว่าง {context}: {message}")
            if time.monotonic() >= deadline:
                return
            time.sleep(0.04)

    @staticmethod
    def _wait_for_file(path: Path, timeout: float) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if path.exists() and path.stat().st_size > 0:
                return True
            time.sleep(0.1)  # v0.5.7 จาก 0.5 → 0.1: ไฟล์ print มักโผล่ใน ~0.1-0.3 วิ
        return False

    # ------------------------------------------------------------------
    # Orchestrator ระดับ "1 วันปลูกที่ทดลอง" (1 คอลัมน์ใน Result sheet)
    #
    # ไม่โยน exception ออกไปข้างนอก — ครอบทุก error เป็น CandidateRunResult(ok=False)
    # เพื่อให้วันปลูกอื่นๆ ในปีเดียวกัน (และปีอื่นๆ) รันต่อได้ ไม่ให้ 1 candidate
    # error ทำให้ทั้ง batch หยุด
    # ------------------------------------------------------------------
    def run_candidate_planting_date(
        self,
        year: int,
        task: PlantingDateTask,
        export_dir: Path,
        screenshot_dir: Optional[Path] = None,
    ) -> CandidateRunResult:
        planting_date = task.planting_date
        try:
            self.set_planting_date(planting_date)
            self.calculate()
            self.open_irrigation_schedule()
            exported_file = self.export_results(year, planting_date, export_dir)

            schedule_shot = graph_shot = None
            if task.capture_screenshot:
                if not screenshot_dir:
                    raise ControlsNotConfiguredError(
                        "capture_screenshot=True แต่ไม่ได้ระบุ screenshot_dir"
                    )
                schedule_shot, graph_shot = self.capture_screenshots(
                    year, planting_date, screenshot_dir
                )

            return CandidateRunResult(
                year=year,
                planting_date=planting_date,
                ok=True,
                exported_file=exported_file,
                schedule_screenshot=schedule_shot,
                graph_screenshot=graph_shot,
            )
        except Exception as exc:  # noqa: BLE001 -- กันไม่ให้ 1 วันปลูก error ล้มทั้ง batch
            logger.warning(
                "ปี %s วันปลูก %s รันไม่สำเร็จ: %s", year, planting_date, exc
            )
            return CandidateRunResult(
                year=year, planting_date=planting_date, ok=False, error_message=str(exc)
            )

    # ------------------------------------------------------------------
    # Orchestrator ระดับ "1 ปี": เปิดไฟล์ climate/rain ครั้งเดียว แล้ววนสั่งคำนวณ
    # ใหม่ทุกวันปลูกที่ทดลอง (planting_dates) — ยืนยันจากผู้ใช้แล้วว่านี่คือ flow
    # จริง ไม่ใช่รันครั้งเดียวจบต่อปี
    #
    # ถ้าเปิดไฟล์ climate/rain ไม่สำเร็จตั้งแต่แรก ถือว่าทั้งปีนี้ error ไปเลย
    # (ไม่มีทางรันวันปลูกไหนต่อได้ถ้าไฟล์พื้นฐานเปิดไม่ได้)
    # ------------------------------------------------------------------
    def run_year(
        self,
        year: int,
        tasks: list[PlantingDateTask],
        climate_file: Path,
        rain_file: Path,
        export_dir: Path,
        screenshot_dir: Optional[Path] = None,
        on_candidate_done=None,
        should_stop=None,
    ) -> YearRunResult:
        try:
            self.open_climate_file(climate_file)
            self.open_rain_file(rain_file)
        except Exception as exc:  # noqa: BLE001 -- กันไม่ให้ 1 ปี error ล้มทั้ง batch
            logger.warning("ปี %s เปิดไฟล์ climate/rain ไม่สำเร็จ: %s", year, exc)
            return YearRunResult(
                year=year, ok=False, error_message=f"เปิดไฟล์ climate/rain ไม่สำเร็จ: {exc}"
            )

        candidates = []
        stopped = False
        for task in tasks:
            # เช็คคำสั่งหยุด "ก่อนเริ่มทุกวันปลูก" ไม่ใช่แค่ตอนขึ้นปีใหม่ (v0.2.3 —
            # เดิมกดหยุดแล้วต้องรอจนครบทั้งปีถึงหยุดจริง) — หยุดกลางวันปลูกที่
            # กำลังทำอยู่ไม่ได้ เพราะจะทิ้ง CropWat ค้างครึ่งทาง (dialog เปิดค้าง
            # ฯลฯ) แต่หยุดระหว่างวันปลูก = สถานะสะอาด รันต่อ/รันใหม่ได้เสมอ
            if should_stop is not None and should_stop():
                logger.info("ปี %s: ได้รับคำสั่งหยุด — หยุดหลังทำไป %s วันปลูก", year, len(candidates))
                stopped = True
                break
            result = self.run_candidate_planting_date(year, task, export_dir, screenshot_dir)
            candidates.append(result)
            # แจ้ง progress ระดับวันปลูกให้ผู้เรียก (runner ใช้ขับ progress bar
            # ทั้งหน้าเว็บและ overlay ลอย) — callback พังต้องไม่ล้มการรัน
            if on_candidate_done is not None:
                try:
                    on_candidate_done(result)
                except Exception:  # noqa: BLE001
                    logger.exception("on_candidate_done callback ล้มเหลว (ไม่กระทบการรัน)")
        failed = [c for c in candidates if not c.ok]
        ok = not failed and not stopped
        error_message = (
            None
            if not failed
            else f"{len(failed)}/{len(candidates)} วันปลูกที่ทดลองล้มเหลว "
            f"(เช่น {failed[0].planting_date:%d/%m}: {failed[0].error_message})"
        )
        return YearRunResult(
            year=year, ok=ok, candidates=candidates, error_message=error_message, stopped=stopped
        )
