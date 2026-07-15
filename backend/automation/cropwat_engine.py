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

import ctypes
import logging
import re
import threading
import time
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

import win32api
import win32con
import win32gui
import win32process
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

    def __init__(self, background_mode: bool = False, speed_multiplier: float = 1.0) -> None:
        self.app: Optional[Application] = None
        self.main_window = None
        self.background_mode = background_mode
        # v0.5.14: ตัวคูณเวลาหน่วงทุกจุดที่รอ CropWat ประมวลผลคำสั่ง (ดู _sleep) —
        # มาจาก config.SPEED_MULTIPLIERS ตาม settings.speed_preset ของผู้ใช้ กัน
        # คอมรุ่นเก่าแฮงค์เพราะรอไม่พอ (ค่าเริ่มต้น 1.0 = พฤติกรรมเดิมทุกประการ)
        self.speed_multiplier = speed_multiplier

    def _sleep(self, base_seconds: float) -> None:
        """time.sleep คูณด้วย speed_multiplier — ใช้แทน time.sleep() ตรงๆ ทุกจุด
        ในไฟล์นี้ที่รอ CropWat ประมวลผลคำสั่ง (ไม่ใช่ทุก sleep — ตัวที่เป็น "ช่วง
        เวลารอสัญญาณ" เช่น deadline ของ error dialog ก็ scale เหมือนกันเพื่อความ
        สม่ำเสมอ)"""
        time.sleep(base_seconds * self.speed_multiplier)

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

    @staticmethod
    def _duplicate_window_message(class_name: str, handles: list[int]) -> str:
        titles = [win32gui.GetWindowText(h) or "(ไม่มีชื่อ)" for h in handles]
        return (
            f"เจอหน้าต่างโมดูลนี้ ({class_name}) เปิดค้างพร้อมกัน {len(handles)} "
            f"หน้าต่าง: {', '.join(titles)} — ระบุไม่ได้ว่าอันไหนถูกต้อง กรุณาเปิด "
            "เมนู Window ใน CropWat แล้วปิดหน้าต่างที่ซ้ำให้เหลือแค่บานเดียวก่อนรันใหม่"
        )

    @staticmethod
    def _is_empty_module_window(hwnd: int) -> bool:
        """ฟอร์มเปล่าที่ File->New สร้างไว้เบิกทาง (ยังไม่โหลดไฟล์) — title จะลงท้าย
        'untitled' หรือว่างเปล่า ต่างจากบานที่โหลดไฟล์แล้วซึ่งมี path เต็มต่อท้าย"""
        title = (win32gui.GetWindowText(hwnd) or "").strip().lower()
        return title == "" or title.endswith("untitled")

    def _resolve_duplicate_module_windows(self, class_name: str, handles: list[int]) -> list[int]:
        """v0.7.1 — เขียนใหม่จากการรันจริงบทผู้ใช้ (journey audit): เจอกรณีที่ทั้ง
        2 บานไม่มี path ในชื่อ (เช่น 'Day ETo PenMon' + '... - untitled') ทำให้
        ตรรกะเดิม (เก็บเฉพาะบานที่มี backslash) หาบานเก็บไม่เจอ → raise → ทุกวัน
        ปลูกหลังจากนั้นล้มทั้งปี — เปลี่ยนหลักการเป็น "ลดให้เหลือบานเดียวเสมอ ไม่
        ยอมแพ้ ไม่ raise": ปิดฟอร์มเปล่า 'untitled' ทิ้งก่อน (ส่วนเกินแน่ๆ) เก็บ 1
        บานที่เหลือ — เลือกบานที่โหลดไฟล์แล้วใหม่สุด ถ้าไม่มีเลยก็เก็บบานใหม่สุด
        (handle มากสุด = สร้างทีหลังสุด) การรันแบบทิ้งไว้ต้อง self-heal เองเสมอ"""
        loaded = [h for h in handles if not self._is_empty_module_window(h)]
        keep = max(loaded) if loaded else max(handles)
        for hwnd in handles:
            if hwnd == keep:
                continue
            reason = "ฟอร์มเปล่า/untitled" if self._is_empty_module_window(hwnd) else "บานเก่า/ซ้ำ"
            logger.warning(
                "%s: หน้าต่างซ้ำ — ปิดบาน %s '%s' อัตโนมัติ (self-heal)",
                class_name, reason, win32gui.GetWindowText(hwnd) or "(ไม่มีชื่อ)",
            )
            self._force_close_window(hwnd)

        remaining = find_windows(
            class_name=class_name, top_level_only=False, process=self.app.process
        )
        if len(remaining) > 1:
            # ปิดไม่ลงรอบแรก — best-effort ปิดทุกบานยกเว้นใหม่สุดอีกครั้ง
            newest = max(remaining)
            for hwnd in remaining:
                if hwnd != newest:
                    self._force_close_window(hwnd, timeout=3.0)
            remaining = find_windows(
                class_name=class_name, top_level_only=False, process=self.app.process
            )
        return remaining or handles

    def _force_close_window(self, hwnd: int, timeout: float = 5.0) -> None:
        """ปิดหน้าต่าง MDI child แบบพยายามหนัก: WM_CLOSE ซ้ำทุก 1 วิ + ตอบ No ให้
        save-prompt + เคลียร์ error dialog ที่อาจบัง จนกว่าจะปิดได้หรือหมดเวลา"""
        deadline = time.monotonic() + timeout
        next_signal = 0.0
        while win32gui.IsWindow(hwnd) and time.monotonic() < deadline:
            now = time.monotonic()
            if now >= next_signal:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                next_signal = now + 1.0
            self._answer_no_to_save_prompt()
            self._poll_error_dialog()
            self._sleep(0.15)

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
            handles = self._resolve_duplicate_module_windows(class_name, handles)

        try:
            window = self.app.window(class_name=class_name, top_level_only=False)
            window.wait("exists enabled visible ready", timeout=10)
        except ElementAmbiguousError:
            # v0.5.32 — ยืนยันจากผู้ใช้จริง: เช็คจำนวนหน้าต่างข้างบนแล้วผ่าน (เจอ
            # แค่ 1 ตอนนั้น) แต่ระหว่างที่ยังไม่ทันเรียกบรรทัดนี้ มีหน้าต่างที่ 2 โผล่
            # ขึ้นมาแทรก (race condition) — v0.5.34 ลอง self-heal ก่อน (ปิดฟอร์มเปล่า
            # "untitled" ที่ค้าง) แล้วค่อยลองใหม่ ถ้ายังกำกวมจริงถึงจะยอมแพ้
            handles = find_windows(
                class_name=class_name, top_level_only=False, process=self.app.process
            )
            self._resolve_duplicate_module_windows(class_name, handles)
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
        self._sleep(0.2)
        return window

    def _invoke_menu(self, menu_path: str) -> None:
        """เรียกเมนูของหน้าต่างหลัก — โหมดปกติใช้ menu_select ของ pywinauto (ซึ่ง
        set_focus ดึงหน้าต่างขึ้นมาก่อนเสมอ) โหมดเบื้องหลังหา id ของ menu item จาก
        โครงสร้างเมนู (อ่านได้โดยไม่ต้อง focus) แล้ว post WM_COMMAND ตรงไปที่
        หน้าต่างหลักแบบเดียวกับที่ Windows ส่งให้ตอนผู้ใช้คลิกเมนูจริง

        v0.5.10 — บทเรียนสำคัญจากการ probe CropWat จริงระหว่าง audit: item.is_enabled()
        ของเมนู Delphi MDI ตัวนี้ "คืน False แบบไม่ตรงความจริง" — PostMessage คำสั่ง
        เดียวกันทำงานได้ปกติ (harvest คำนวณใหม่, หน้าต่างผลลัพธ์โผล่, ไฟล์ถูกเขียน)
        ทั้งที่ is_enabled() รายงานว่า disabled — เพราะสถานะ enabled ของเมนูจะถูก
        อัปเดตจริงตอน WM_INITMENU (ตอนผู้ใช้กดเปิดเมนู) เท่านั้น ไม่ใช่ตอนอ่านผ่าน
        API เฉยๆ เคย gate ด้วย is_enabled() (v0.5.1-v0.5.9) เลยทำให้ File->Open/
        calculate fail ทุกปีด้วย false negative (ยืนยันจาก log จริงของผู้ใช้: "เมนู
        File->Open ถูก disable" ปี 1996-2025 ทั้งที่ปีอื่นรันได้) — เลิกเช็ค
        is_enabled() ปล่อยให้สัญญาณ downstream (หน้าต่างผลลัพธ์โผล่/ไฟล์ถูกเขียน)
        เป็นตัวจับกรณีคำสั่งไม่มีผลจริงแทน + settle หลังยิงคำสั่งกัน race กับคำสั่งถัดไป"""
        if not self.background_mode:
            self.main_window.menu_select(menu_path)
            return
        item = self.main_window.menu().get_menu_path(menu_path)[-1]
        # เมธอดชื่อ item_id() (ยืนยันจาก source ของ pywinauto — ไม่ใช่ .id())
        win32gui.PostMessage(self.main_window.handle, win32con.WM_COMMAND, item.item_id(), 0)
        self._sleep(0.15)  # settle: ให้เวลา CropWat ประมวลผลคำสั่งนี้ก่อนยิงคำสั่งถัดไป

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
            # v0.5.32 — ยืนยันจากผู้ใช้จริง: ส่ง WM_CLOSE ครั้งเดียวแล้วรอเฉยๆ ไม่พอ
            # บางครั้งหน้าต่างไม่ตอบสนองครั้งแรก (จังหวะไม่ตรง/มี dialog บังอยู่) —
            # ส่งซ้ำทุก 1 วิ + เช็ค error dialog ที่อาจบล็อกการปิดอยู่ด้วย (ไม่ใช่แค่
            # save-prompt) ยืดเวลารอจาก 5 เป็น 8 วิ ให้สอดคล้องกับ
            # REVEAL_TIMEOUT_SECONDS ของ background watcher
            deadline = time.monotonic() + 8
            next_close_signal = 0.0
            while win32gui.IsWindow(handle) and time.monotonic() < deadline:
                now = time.monotonic()
                if now >= next_close_signal:
                    win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)
                    next_close_signal = now + 1.0
                self._answer_no_to_save_prompt()
                self._poll_error_dialog()
                self._sleep(0.15)
            if win32gui.IsWindow(handle):
                logger.warning("%s: ปิดหน้าต่าง '%s' ไม่สำเร็จภายใน 8 วินาที", module_label, title)
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

        # v0.7.1 — บั๊กที่จับได้จากการรันจริงบทผู้ใช้: มีหน้าต่างชื่อตรง title_re
        # ("Open") มากกว่า 1 บาน (เช่น dialog ของไฟล์ก่อนหน้ายังไม่ทันปิดสนิทตอน
        # เรียกไฟล์ถัดไป) ทำให้ self.app.window(title_re=...) โยน ambiguous ดิบๆ
        # ("There are 2 elements...") ล้มทั้ง batch — แก้ทั้งวงจร: (1) เคลียร์
        # dialog เก่าที่ค้างก่อนเรียกเมนู (2) ชี้เป้า dialog ใหม่ด้วย hwnd ตรงๆ
        # (เลือกเฉพาะบานที่ visible, เอาบานที่ไม่เคยเห็นก่อนเรียกเมนู) (3) หลังกด
        # Open ต้อง "รอจน dialog ปิดจริง" ก่อนไปต่อ — กันการทิ้ง dialog ค้างเป็น
        # ระเบิดเวลาให้ไฟล์ถัดไป
        stale = self._find_file_dialogs(dialog_title_re)
        for hwnd in stale:
            logger.warning("เจอ dialog เปิดไฟล์ค้างจากขั้นก่อนหน้า — ปิดทิ้งก่อน (hwnd=%#x)", hwnd)
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            self._sleep(0.3)
        before = set(self._find_file_dialogs(dialog_title_re))

        self._invoke_menu("File->Open")

        # ยืนยันจาก screenshot จริงของผู้ใช้ (v0.1.12): ก่อน dialog เปิดไฟล์จะโผล่
        # CropWat อาจเด้งถาม "Save changes to current climate/rain data ?"
        # (Yes/No/Cancel) — เป็น modal ที่บังไม่ให้ dialog เปิดไฟล์โผล่จนกว่าจะตอบ
        # ต้องคอยเช็คแล้วตอบ No ให้ระหว่างรอ
        dialog_hwnd: Optional[int] = None
        deadline = time.monotonic() + 10
        while dialog_hwnd is None:
            fresh = [h for h in self._find_file_dialogs(dialog_title_re) if h not in before]
            if fresh:
                dialog_hwnd = fresh[-1]  # บานใหม่ล่าสุด
                break
            self._answer_no_to_save_prompt()
            if time.monotonic() > deadline:
                raise StepTimeoutError(
                    f"dialog เปิดไฟล์ (title_re={dialog_title_re!r}) ไม่โผล่ภายใน 10 วินาที"
                )
            self._sleep(0.2)

        dialog = self.app.window(handle=dialog_hwnd)
        self._hide_dialog_offscreen(dialog)
        dialog.wait("exists enabled visible ready", timeout=10)
        self._sleep(0.3)  # เผื่อเวลาให้ dialog พร้อมรับ input จริงๆ ก่อนพิมพ์

        target = str(file_path)
        field = dialog[filename_field]
        for attempt in range(2):
            field.set_edit_text(target)
            self._sleep(0.2)
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

        # รอจน dialog ปิดจริง (ไฟล์ถูกเปิดแล้ว) — ถ้า 3 วิแล้วยังอยู่ ลองกด Open
        # ซ้ำหนึ่งครั้ง (message แรกอาจหลุด) ก่อนจะ fail ชัดเจน
        close_deadline = time.monotonic() + 3.0
        retried = False
        while win32gui.IsWindow(dialog_hwnd) and win32gui.IsWindowVisible(dialog_hwnd):
            if time.monotonic() > close_deadline:
                if not retried:
                    retried = True
                    ok_btn = self._find_child_button(dialog_hwnd, r"&?Open$")
                    if ok_btn is not None:
                        win32gui.SendMessage(ok_btn, self._BM_CLICK, 0, 0)
                    close_deadline = time.monotonic() + 3.0
                    continue
                raise StepTimeoutError(
                    f"กด Open แล้ว dialog เปิดไฟล์ไม่ยอมปิด (ไฟล์ {file_path.name}) — "
                    "CropWat อาจไม่ยอมรับไฟล์นี้"
                )
            self._answer_no_to_save_prompt()
            self._sleep(0.15)

    def _find_file_dialogs(self, title_re: str) -> list[int]:
        """หา file dialog (#32770) ของ process ที่ title ตรง pattern และ "มองเห็น
        ได้จริง" เท่านั้น — คืนเรียงตาม hwnd (บานใหม่มักได้ hwnd ใหม่กว่า)"""
        pattern = re.compile(title_re)
        out: list[int] = []
        try:
            for hwnd in find_windows(
                class_name="#32770", top_level_only=True,
                process=self.app.process, visible_only=False,
            ):
                if win32gui.IsWindowVisible(hwnd) and pattern.match(win32gui.GetWindowText(hwnd) or ""):
                    out.append(hwnd)
        except ElementNotFoundError:
            pass
        return sorted(out)

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
    def _real_set_focus(self, hwnd: int) -> None:
        """v0.5.12 — ย้าย focus จริงข้าม process ด้วย SetFocus() ผ่าน AttachThreadInput
        (ไม่ใช่ post message จำลอง WM_SETFOCUS) บทเรียนสำคัญที่สุดจากการ probe ซ้ำ
        หลายสิบรอบระหว่าง audit: TCropForm's TMaskEdit มีค่า "ข้อความที่เห็น" กับ
        "ค่าที่ CropWat ใช้คำนวณจริง" เป็นคนละตัวกัน — WM_SETTEXT/CM_EXIT/WM_KEYDOWN
        Tab (ทุกแบบที่เป็นแค่ post message จำลอง) เปลี่ยนได้แค่ตัวแรก ไม่เคย commit
        เข้าตัวที่สองเลย ไม่ว่าจะรอนานแค่ไหนหรือลองกี่รอบ — SetFocus() API ตัวจริง
        (ให้ Windows ส่ง WM_KILLFOCUS/WM_SETFOCUS ตามกลไกจริงเหมือนผู้ใช้คลิกเอง)
        เท่านั้นที่ commit เข้า model จริง — ยืนยันด้วย pipeline เต็ม 9 วันปลูกต่างกัน
        (ตั้งวันปลูก→CWR→Schedule→export→อ่านไฟล์จริง) ถูกต้องครบทุกอัน ไม่ต้อง
        ดึงหน้าต่างมา foreground เลย (AttachThreadInput ทำงานได้แม้ z-order เดิม)"""
        target_tid, _ = win32process.GetWindowThreadProcessId(hwnd)
        cur_tid = win32api.GetCurrentThreadId()
        attached = False
        try:
            attached = bool(ctypes.windll.user32.AttachThreadInput(cur_tid, target_tid, True))
            ctypes.windll.user32.SetFocus(hwnd)
        finally:
            if attached:
                ctypes.windll.user32.AttachThreadInput(cur_tid, target_tid, False)

    def _commit_field(self, field, cfg) -> None:
        """ยืนยันค่าในช่อง — โหมดเบื้องหลังย้าย focus จริงเข้าช่อง (ให้พิมพ์ค่าไป
        ก่อนหน้านี้แล้ว) แล้วย้ายออกไปหน้าต่างหลัก (ดู _real_set_focus) โหมดปกติ
        กด Tab จริง (จำลองคีย์บอร์ด ไม่ขยับเมาส์)"""
        if self.background_mode:
            self._real_set_focus(self.main_window.handle)
        elif cfg.confirm_key:
            field.type_keys(cfg.confirm_key)

    @staticmethod
    def _parse_field_date(text: str) -> Optional[tuple[int, str, int]]:
        """แยกข้อความในช่องวันที่เป็น (เลขหน้า, ตัวคั่น, เลขหลัง) — ไม่ตีความว่า
        อันไหนวัน/เดือน (ขึ้นกับ Region ของเครื่อง)

        ยืนยันจาก probe จริง: CropWat อาจเติมปีต่อท้ายให้เอง (เช่น '01/05' → '01/05/94')
        หลัง commit — จึงยอมให้มีส่วนปีต่อท้าย (ไม่บังคับ) แล้วมองข้ามไป อ่านแค่
        วัน/เดือน 2 ตัวแรกพอ ไม่งั้น readback จะ parse ไม่ผ่านทั้งที่ค่าถูกต้อง"""
        m = re.match(r"^\s*(\d{1,2})(\D)(\d{1,2})(?:\D\d{2,4})?\s*$", text or "")
        if not m:
            return None
        return int(m.group(1)), m.group(2), int(m.group(3))

    def set_planting_date(self, planting_date: date) -> None:
        """v0.5.12 — เขียนใหม่รอบที่ 3 จากการ probe CropWat จริงหลายสิบรอบระหว่าง
        audit (ล้มสมมติฐานเดิมทุกข้อที่เคยเชื่อ รวมถึงของ v0.5.10 เอง):

        ช่องวันปลูกเป็น TMaskEdit "dd/mm" ตายตัว มี "ข้อความที่เห็น" กับ "ค่าที่
        CropWat ใช้คำนวณจริง" เป็นคนละตัวกัน — set_edit_text (WM_SETTEXT) เปลี่ยน
        ได้แค่ตัวแรกเท่านั้น ไม่เคย commit เข้าตัวที่สองเลยไม่ว่าจะใช้ CM_EXIT/Tab/
        WM_KILLFOCUS/รอนานแค่ไหน (v0.5.10/v0.5.11 ที่ "ทดสอบผ่าน" ตอนนั้นเป็นเพราะ
        ฟอร์มมีค่า commit ค้างจากการทดลองมือของผู้พัฒนาก่อนหน้า ไม่ใช่ set_edit_text
        ทำงานจริง — พบตอนทดสอบกับ CropWat ที่เพิ่ง restart สะอาดๆ ครั้งแรก) วิธีที่
        commit เข้า model จริงคือ SetFocus() API ตัวจริง (ดู _real_set_focus) ย้าย
        focus เข้าช่องก่อนพิมพ์ แล้วย้ายออกหลังพิมพ์ — ยืนยันด้วย pipeline เต็ม
        9 วันปลูกต่างกันติดต่อกัน (ตั้งวันปลูก→CWR→Schedule→export→อ่านไฟล์จริง)
        ถูกต้องครบทุกอัน"""
        self._require_connected()
        cfg = controls.CROP_SCREEN
        crop_window = self._focus_mdi_child(cfg.window_class_name)
        field = crop_window.child_window(class_name=cfg.planting_date_field_class_name)

        text = f"{planting_date.day:02d}/{planting_date.month:02d}"
        if self.background_mode:
            self._real_set_focus(field.handle)
            self._sleep(0.15)
        field.set_edit_text(text)
        self._commit_field(field, cfg)
        self._raise_if_error_dialog(f"ตั้งวันปลูก {text}")

        # อ่านช่องกลับมายืนยันว่าค่าลงจริง (วน 4 รอบสั้นๆ เผื่ออัปเดตช้าเสี้ยววินาที)
        # — ถ้าเครื่องไหน mask เป็น mm/dd (ยังไม่เคยเจอ) ช่องจะโชว์ไม่ตรง แล้วด่าน
        # ตรวจไฟล์ .txt จะจับได้อีกชั้น
        expected = (planting_date.day, planting_date.month)
        for _ in range(4):
            after = self._parse_field_date(field.window_text())
            if after and (after[0], after[2]) == expected:
                return
            self._sleep(0.1)

        raise CropWatReportedError(
            f"CropWat ไม่รับวันปลูก {planting_date:%d/%m} เข้าช่องกรอก "
            f"(เห็น field={field.window_text()!r}) — ลองปิด-เปิด CropWat แล้วรันใหม่"
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
        อยู่แล้ว (timeout 10 วิ) เป็นการยืนยันว่าคำนวณสำเร็จจริงในตัวเอง

        v0.5.9 — เจอบั๊กโครงสร้างที่มีมาแต่ต้น (เพิ่งชัดตอนรันเบื้องหลังหลายวันปลูก
        ติดกันเร็วๆ, ยืนยันจาก screenshot จริงของผู้ใช้ที่แถบสถานะโชว์วันปลูกค้าง
        เป็นวันนี้เป็นสีแดง ทั้งที่ตั้งวันปลูกใหม่ไปแล้ว): หน้าต่างนี้ถ้าเปิดค้างจาก
        วันปลูกก่อนหน้าอยู่แล้ว _focus_mdi_child ด้านล่างจะ "เจอ" มันทันทีโดยไม่ต้อง
        รอคำนวณใหม่เสร็จจริงเลย เพราะมันไม่เคยหายไปไหน — สัญญาณ "หน้าต่างโผล่แล้ว"
        เลยกลายเป็นเท็จ ทำให้ export ไปได้ผลลัพธ์ค้างของวันปลูกก่อนหน้าซ้ำ (เหมือน
        ที่เคยแก้ไว้แล้วสำหรับหน้าต่างกราฟใน capture_screenshots — จุดนี้ตกหล่นไป)
        ปิดหน้าต่างเก่าทิ้งก่อนคำนวณทุกครั้ง บังคับให้ CropWat ต้องสร้างหน้าต่างใหม่
        จริงๆ ตอนคำนวณเสร็จ ทำให้ _focus_mdi_child ด้านล่างเป็นสัญญาณ "คำนวณเสร็จ
        จริง" ที่เชื่อถือได้อีกครั้ง"""
        self._require_connected()
        cfg = controls.CALCULATE
        sched_cfg = controls.IRRIGATION_SCHEDULE
        if sched_cfg.window_class_name:
            self._close_stale_module_windows(
                sched_cfg.window_class_name, None, "irrigation schedule table"
            )

        self._invoke_menu(cfg.irrigation_scheduling_menu_path)
        self._raise_if_error_dialog("คำนวณ Irrigation Scheduling")

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
            self._sleep(0.8)

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

    # เพดานเวลาที่รอ PrintWindow — ดู docstring ของ _print_window_with_timeout
    _PRINT_WINDOW_TIMEOUT_SECONDS = 8.0

    def _print_window_with_timeout(self, hwnd: int, hdc: int, flags: int) -> bool:
        """v0.5.19 — PrintWindow (Win32 API) ไม่มี timeout ในตัวเอง เป็น call
        แบบ synchronous ที่รอให้หน้าต่างเป้าหมายวาดตัวเองเสร็จ ถ้า UI thread ของ
        มันไม่ได้ปั๊ม message (ระบบมีงานอื่นแย่ง CPU/GPU หนักๆ) เรียกตรงๆ อาจค้าง
        ตลอดกาล — รันใน thread แยกแล้วรอแบบมีเพดานเวลาเอง ถ้าไม่เสร็จทันถือว่า
        ล้มเหลว (คืน False ให้ caller ถอยไปใช้วิธีอื่น) — thread ที่ค้างจะถูกทิ้งไว้
        เบื้องหลัง (daemon) ตายไปพร้อม process ตอนปิดโปรแกรม ไม่ได้ฆ่ากลางคัน
        เพราะ Win32 API ไม่มีทางบังคับยกเลิก PrintWindow ที่กำลังทำงานอยู่ได้"""
        result: dict[str, bool] = {}

        def _call() -> None:
            result["ok"] = bool(ctypes.windll.user32.PrintWindow(hwnd, hdc, flags))

        t = threading.Thread(target=_call, daemon=True, name="printwindow-guard")
        t.start()
        t.join(self._PRINT_WINDOW_TIMEOUT_SECONDS)
        return result.get("ok", False)

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
            # v0.5.19 — บั๊กที่เจอจากผู้ใช้จริง: "แฮงค์บางครั้งถ้ามีโปรแกรมอื่นมา
            # บังจอ" — PrintWindow เป็นคำสั่งเดียวในระบบที่ไม่มี timeout ในตัวเอง
            # เลย มันรอให้ CropWat วาดตัวเองเสร็จแบบไม่มีเพดานเวลา ถ้า UI thread
            # ของ CropWat ประมวลผลช้า (เช่นตอนโปรแกรมอื่นแย่ง GPU/CPU หนักๆ ตอน
            # บังจอ) เรียกแล้วอาจค้างตลอดกาล ทำให้ automation ทั้งตัวแฮงค์ตามไป
            # ด้วย — ใส่ deadline เอง (รันใน thread แยก รอแบบมีเพดานเวลา) ถ้าไม่
            # เสร็จใน 8 วิ ถือว่าล้มเหลว ถอยไปถ่ายจากหน้าจอแทน (ไม่ต้องพึ่ง CropWat
            # ตอบสนองเลย จึงไม่ติดปัญหาเดียวกัน)
            ok = self._print_window_with_timeout(hwnd, save_dc.GetSafeHdc(), PW_RENDERFULLCONTENT)
            if not ok:
                # บางหน้าต่าง/ไดรเวอร์ไม่รองรับ PrintWindow หรือค้างเกินเวลา →
                # ถอยไปถ่ายจากจอแทน (ภาพอาจมีอย่างอื่นบัง แต่ดีกว่าไม่มีภาพเลย)
                logger.warning(
                    "PrintWindow ล้มเหลวหรือค้างเกิน %ss — ถอยไปถ่ายจากหน้าจอแทน (ภาพอาจถูกบัง)",
                    self._PRINT_WINDOW_TIMEOUT_SECONDS,
                )
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
            self._sleep(0.2)
            return True
        except (ElementNotFoundError, ElementAmbiguousError):
            return False

    # ------------------------------------------------------------------
    # ตรวจจับ error/warning dialog (ใช้ทั้งเช็ค inline และ poll หลัง calculate)
    # ------------------------------------------------------------------
    # ปุ่มบน dialog กดผ่าน message ตรงๆ (BM_CLICK) ไม่ต้องพึ่ง focus/เมาส์จริง
    _BM_CLICK = 0x00F5

    def _find_error_dialogs(self) -> list[int]:
        """หา "dialog จริง" ของ CropWat ที่เข้าข่าย error/warning ทั้งหมดตอนนี้ —
        กรองด้วย class (TMessageForm/#32770) ควบคู่กับ title เสมอ

        v0.5.33 — REGRESSION FIX ยืนยันจากการ probe เครื่องจริง: เดิมใช้
        self.app.window(title_re=...) ซึ่งกรองแค่ title — ไปจับโดนหน้าต่างซ่อน
        ภายในของ Delphi (class TApplication, title "FAO CROPWAT 8.0 for
        Windows", visible ตลอดเวลาที่โปรแกรมเปิด) ทันทีที่ dialog จริงโผล่ จึงมี
        2 หน้าต่างตรง pattern พร้อมกัน → เข้าเคส ambiguous → รายงาน error ได้แต่
        ไม่เคยกดปิด dialog จริง → dialog ค้างเป็น modal บล็อกทุกอย่าง → หน้าต่าง
        โมดูลพอกซ้ำ → ทุกวันปลูกถัดไปล้มเหลวหมด (ตรงกับ screenshot ผู้ใช้)"""
        cfg = controls.ERROR_DIALOG
        title_pattern = re.compile(cfg.title_re)
        dialogs: list[int] = []
        for class_name in cfg.dialog_class_names:
            try:
                # visible_only=False — dialog ที่โดน watcher เหวี่ยงออกนอกจอ/ตั้ง
                # alpha=0 ยังต้องหาเจอ+กดปิดได้ (find_windows default visible_only=True
                # เสี่ยงพลาดถ้าอนาคตมีการ SW_HIDE dialog เหล่านี้)
                for hwnd in find_windows(
                    class_name=class_name, top_level_only=True,
                    process=self.app.process, visible_only=False,
                ):
                    title = win32gui.GetWindowText(hwnd) or ""
                    if title_pattern.match(title):
                        dialogs.append(hwnd)
            except ElementNotFoundError:
                continue
        return dialogs

    @staticmethod
    def _find_child_button(dialog_hwnd: int, label_re: str) -> Optional[int]:
        """หาปุ่มลูกของ dialog จาก caption (รองรับ & accelerator เช่น '&No')"""
        pattern = re.compile(label_re)
        result: list[int] = []

        def _cb(child: int, _param) -> bool:
            cls = win32gui.GetClassName(child)
            if cls in ("Button", "TButton") and pattern.match(win32gui.GetWindowText(child) or ""):
                result.append(child)
            return True

        try:
            win32gui.EnumChildWindows(dialog_hwnd, _cb, None)
        except Exception:  # noqa: BLE001 -- dialog ตายไประหว่าง enumerate = ไม่เจอปุ่ม
            pass
        return result[0] if result else None

    @staticmethod
    def _read_dialog_statics(dialog_hwnd: int) -> str:
        """อ่านข้อความจริงบน dialog จาก Static child controls — native MessageBox
        (#32770) เก็บข้อความ error ("Access violation at ...") ไว้ใน Static เสมอ
        อ่านได้ตรงๆ (ต่างจาก TMessageForm ของ Delphi ที่วาดข้อความเอง อ่านไม่ได้)"""
        texts: list[str] = []

        def _cb(child: int, _param) -> bool:
            if win32gui.GetClassName(child) == "Static":
                text = (win32gui.GetWindowText(child) or "").strip()
                if text:
                    texts.append(text)
            return True

        try:
            win32gui.EnumChildWindows(dialog_hwnd, _cb, None)
        except Exception:  # noqa: BLE001
            pass
        return " ".join(texts)

    def _poll_error_dialog(self) -> Optional[str]:
        """เช็คครั้งเดียว (ไม่รอ) ว่ามี error dialog เด้งอยู่ไหม ถ้ามีให้จัดการ
        "ทุกบาน" ที่เจอ (กดปิด + เก็บข้อความ) แล้วคืนข้อความรวม — v0.5.33 เขียน
        ใหม่ทั้งตัวด้วย win32 ตรงๆ แทน pywinauto WindowSpecification เพื่อตัด
        ปัญหา ambiguous ออกไปเลย (จัดการหลาย dialog พร้อมกันได้ ไม่ใช่ยอมแพ้)"""
        dialogs = self._find_error_dialogs()
        if not dialogs:
            return None

        messages: list[str] = []
        for hwnd in dialogs:
            title = win32gui.GetWindowText(hwnd) or "(ไม่มีชื่อ)"
            try:
                self._fling_hwnd_offscreen(hwnd)
                # แยก "คำถาม" ออกจาก "error": dialog ที่มีปุ่ม No คือคำถาม
                # (Yes/No/Cancel) ตอบ No เสมอ — ห้าม Yes เพราะจะบันทึกทับไฟล์
                # ข้อมูลต้นทางของผู้ใช้
                no_button = self._find_child_button(hwnd, r"&?No$")
                if no_button is not None:
                    win32gui.SendMessage(no_button, self._BM_CLICK, 0, 0)
                    logger.info("ตอบ No อัตโนมัติให้ prompt ถามบันทึกข้อมูล (ไม่ใช่ error)")
                    continue

                detail = self._read_dialog_statics(hwnd)
                if detail:
                    messages.append(f"CropWat แจ้ง: {detail} (dialog '{title}')")
                else:
                    messages.append(f"CropWat แสดง dialog '{title}' (อ่านข้อความละเอียดไม่ได้)")

                ok_button = self._find_child_button(hwnd, r"&?OK$")
                if ok_button is not None:
                    win32gui.SendMessage(ok_button, self._BM_CLICK, 0, 0)
                else:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except Exception:  # noqa: BLE001 -- dialog ตายไปพอดีระหว่างจัดการ = ข้าม
                logger.debug("จัดการ error dialog '%s' ไม่สำเร็จ", title, exc_info=True)

        return "; ".join(messages) if messages else None

    def _fling_hwnd_offscreen(self, hwnd: int) -> None:
        """เหวี่ยง dialog ออกนอกจอ (โหมดเบื้องหลังเท่านั้น) แบบรับ hwnd ตรงๆ"""
        if not self.background_mode:
            return
        try:
            flags = win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE
            win32gui.SetWindowPos(hwnd, 0, -32000, -32000, 0, 0, flags)
        except Exception:  # noqa: BLE001
            pass

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
            self._sleep(0.04)

    @staticmethod
    def _wait_for_file(path: Path, timeout: float) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if path.exists() and path.stat().st_size > 0:
                return True
            self._sleep(0.1)  # v0.5.7 จาก 0.5 → 0.1: ไฟล์ print มักโผล่ใน ~0.1-0.3 วิ
        return False

    # ------------------------------------------------------------------
    # Orchestrator ระดับ "1 วันปลูกที่ทดลอง" (1 คอลัมน์ใน Result sheet)
    #
    # ไม่โยน exception ออกไปข้างนอก — ครอบทุก error เป็น CandidateRunResult(ok=False)
    # เพื่อให้วันปลูกอื่นๆ ในปีเดียวกัน (และปีอื่นๆ) รันต่อได้ ไม่ให้ 1 candidate
    # error ทำให้ทั้ง batch หยุด
    # ------------------------------------------------------------------
    # จำนวนครั้งสูงสุดที่ลองทำ "ทั้ง candidate ใหม่" (ตั้งวันปลูก+คำนวณ+export) ถ้า
    # ไฟล์ที่ออกมาวันปลูกไม่ตรง — ดู docstring ของ run_candidate_planting_date
    _MAX_CANDIDATE_ATTEMPTS = 3

    def run_candidate_planting_date(
        self,
        year: int,
        task: PlantingDateTask,
        export_dir: Path,
        screenshot_dir: Optional[Path] = None,
    ) -> CandidateRunResult:
        """v0.5.12 — บทเรียนสำคัญจากการ probe ซ้ำหลายสิบรอบระหว่าง audit: การ commit
        วันปลูกเข้า "โมเดลคำนวณจริง" ของ CropWat (ต่างจากแค่ข้อความในช่อง ซึ่งเปลี่ยน
        ทันทีเสมอ) มีจังหวะที่ไม่แน่นอน (คล้าย race condition ภายใน Delphi เอง) —
        ทดสอบแล้วว่า: รอนานขึ้นเฉยๆ ไม่ช่วย, ยิง WM_KILLFOCUS/SetFocus จริงก็ไม่ช่วย
        100%, แต่ "ลองทำใหม่ทั้งชุด" (ตั้งวันปลูก+คำนวณ+export ใหม่) ในที่สุดจะติด —
        แทนที่จะยอมแพ้ทันทีเมื่อไฟล์ที่ออกมาวันปลูกไม่ตรง (ตาข่ายจาก v0.5.9) ให้ลอง
        ทั้ง candidate ใหม่สูงสุด 3 รอบก่อนถือว่า fail จริง"""
        planting_date = task.planting_date
        from file_engine.txt_parser import TxtParseError, parse_txt

        last_exc: Optional[Exception] = None
        for attempt in range(1, self._MAX_CANDIDATE_ATTEMPTS + 1):
            try:
                self.set_planting_date(planting_date)
                self.calculate()
                self.open_irrigation_schedule()
                exported_file = self.export_results(year, planting_date, export_dir)

                # ตาข่ายสุดท้ายกันบั๊กเงียบ: ตรวจจากไฟล์ .txt ที่ CropWat พิมพ์ออกมาเอง
                # ว่าวันปลูกตรงกับ candidate ที่สั่งจริง (ยอมรับทั้ง dm/md กัน Region ต่าง)
                try:
                    parsed = parse_txt(exported_file)
                except (TxtParseError, OSError) as exc:
                    raise CropWatReportedError(
                        f"ไฟล์ผลลัพธ์ {exported_file.name} อ่านค่า Planting date ไม่ได้: {exc}"
                    ) from exc

                pd_match = re.match(r"^\s*(\d{1,2})/(\d{1,2})\s*$", parsed.planting_date)
                if not pd_match:
                    raise CropWatReportedError(
                        f"ไฟล์ {exported_file.name} มี Planting date รูปแบบแปลก: {parsed.planting_date!r}"
                    )
                a, b = int(pd_match.group(1)), int(pd_match.group(2))
                want_day, want_month = planting_date.day, planting_date.month
                if (a, b) != (want_day, want_month) and (a, b) != (want_month, want_day):
                    raise CropWatReportedError(
                        "ตั้งวันปลูกไม่ติดในผลคำนวณจริง: "
                        f"ไฟล์ {exported_file.name} แสดง {parsed.planting_date} แต่ต้องเป็น "
                        f"{planting_date:%d/%m}"
                    )

                schedule_shot = graph_shot = None
                if task.capture_screenshot:
                    if not screenshot_dir:
                        raise ControlsNotConfiguredError(
                            "capture_screenshot=True แต่ไม่ได้ระบุ screenshot_dir"
                        )
                    schedule_shot, graph_shot = self.capture_screenshots(
                        year, planting_date, screenshot_dir
                    )

                if attempt > 1:
                    logger.info(
                        "ปี %s วันปลูก %s สำเร็จหลังลองใหม่ครั้งที่ %s", year, planting_date, attempt
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
                last_exc = exc
                logger.warning(
                    "ปี %s วันปลูก %s รันไม่สำเร็จ (ครั้งที่ %s/%s): %s",
                    year, planting_date, attempt, self._MAX_CANDIDATE_ATTEMPTS, exc,
                )

        return CandidateRunResult(
            year=year, planting_date=planting_date, ok=False, error_message=str(last_exc)
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
        climate_file_for,
        rain_file_for,
        export_dir: Path,
        screenshot_dir: Optional[Path] = None,
        on_candidate_done=None,
        should_stop=None,
        pause_check=None,
    ) -> YearRunResult:
        """v0.5.21 — เปลี่ยนจากรับ climate_file/rain_file ตายตัว (เปิดครั้งเดียว
        ใช้ทั้งปี) เป็นรับ "ตัว resolve" (Callable[[เดือนปลูก], Path]) เรียกใหม่
        ก่อนทุกวันปลูก — เปิดตามความหมายจริง: ระบบ shift-year เดิม (StationIndex.
        resolve()) มี logic "เลือกไฟล์เดือน <= เดือนปลูก มากที่สุด ถ้าไม่มีเลย
        ในปีนั้นถอยไปปีก่อนหน้า" อยู่แล้ว แต่ runner.py เคย resolve แค่ครั้งเดียว
        จากเดือนแรกสุดของปีแล้วใช้ไฟล์เดียวซ้ำทุกวันปลูก (ยืนยันจาก screenshot
        ผู้ใช้เดิม) ทำให้วันปลูกเดือนหลังๆ ในปีเดียวกันอาจไม่ได้ใช้ไฟล์เดือนที่
        ตรงที่สุดทั้งที่มีอยู่จริง — ตอนนี้ runner.py ส่ง resolver ที่ behavior
        ต่างกันตาม settings.shift_year_per_candidate (ดู _run_years_inner):
        เปิด = resolve ใหม่ทุกวันปลูกตามเดือนจริง, ปิด = resolve ครั้งเดียวคงที่
        (พฤติกรรมเดิมทุกประการ) — open_climate_file/open_rain_file เองมี dedup
        "ข้ามถ้าไฟล์เป้าหมายเปิดอยู่แล้ว" อยู่แล้ว จึงเรียกซ้ำได้ฟรีไม่เปลืองเวลา
        ถ้าเดือนไม่เปลี่ยนจากวันปลูกก่อนหน้า"""
        candidates = []
        stopped = False
        for task in tasks:
            # v0.7.2 — "จุดปลอดภัย" ระหว่างวันปลูก (ไม่มี dialog เปิดค้าง) ให้ระบบ
            # แอบดูเดสก์ท็อปซ่อน park การทำงานไว้ก่อน SwitchDesktop ได้ — กันไม่ให้
            # การสลับจอไปชนจังหวะที่ CropWat กำลังจะโชว์ modal dialog (File->Open)
            # ซึ่งเป็นต้นเหตุ error "Cannot make a visible window modal" ตอน peek
            if pause_check is not None:
                pause_check()
            # เช็คคำสั่งหยุด "ก่อนเริ่มทุกวันปลูก" ไม่ใช่แค่ตอนขึ้นปีใหม่ (v0.2.3 —
            # เดิมกดหยุดแล้วต้องรอจนครบทั้งปีถึงหยุดจริง) — หยุดกลางวันปลูกที่
            # กำลังทำอยู่ไม่ได้ เพราะจะทิ้ง CropWat ค้างครึ่งทาง (dialog เปิดค้าง
            # ฯลฯ) แต่หยุดระหว่างวันปลูก = สถานะสะอาด รันต่อ/รันใหม่ได้เสมอ
            if should_stop is not None and should_stop():
                logger.info("ปี %s: ได้รับคำสั่งหยุด — หยุดหลังทำไป %s วันปลูก", year, len(candidates))
                stopped = True
                break
            try:
                climate_file = climate_file_for(task.planting_date.month)
                rain_file = rain_file_for(task.planting_date.month)
                self.open_climate_file(climate_file)
                self.open_rain_file(rain_file)
            except Exception as exc:  # noqa: BLE001 -- กันไม่ให้ 1 วันปลูก error ล้มทั้งปี
                logger.warning(
                    "ปี %s วันปลูก %s เปิดไฟล์ climate/rain ไม่สำเร็จ: %s",
                    year, task.planting_date, exc,
                )
                result = CandidateRunResult(
                    year=year, planting_date=task.planting_date, ok=False,
                    error_message=f"เปิดไฟล์ climate/rain ไม่สำเร็จ: {exc}",
                )
                candidates.append(result)
                if on_candidate_done is not None:
                    try:
                        on_candidate_done(result)
                    except Exception:  # noqa: BLE001
                        logger.exception("on_candidate_done callback ล้มเหลว (ไม่กระทบการรัน)")
                continue
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


# ------------------------------------------------------------------
# ทางออกฉุกเฉิน: บังคับปิด CropWat จากนอก engine instance ใดๆ (v1.0)
# ------------------------------------------------------------------
#
# ยืนยันจากรายงานผู้ใช้จริง: บางครั้ง CropWat ค้างแบบกดปิดเองไม่ได้เลย (แม้แต่
# กด X) ต้องเข้า Task Manager เท่านั้น — เดิมต้นเหตุหลักคือ watcher ของโหมด
# เบื้องหลังเก่าเหวี่ยง dialog ที่ไม่รู้จักออกนอกจอทั้งที่ยังเป็น modal อยู่
# (ระบบนั้นถูกถอดออกทั้งหมดใน v0.7.0 แล้ว — โหมดเดสก์ท็อปซ่อนไม่ต้องเหวี่ยง
# อะไรเลย) ปุ่มนี้ยังเก็บไว้เป็นทางออกฉุกเฉินสุดท้ายเผื่อ CropWat crash เอง
# (Access violation ฯลฯ ซึ่งเป็นบั๊กภายในของ CropWat)
#
# ฟังก์ชันนี้ไม่ต้องมี engine instance ที่ connect() สำเร็จอยู่ก่อน (เผื่อ state
# ฝั่งเราเองก็อาจพังไปแล้วจากการค้างเดียวกัน) — หา CropWat จาก title/class เดียว
# กับที่ connect() ใช้โดยตรง
#
# หมายเหตุ (v0.7.0): find_windows เห็นเฉพาะหน้าต่างบนเดสก์ท็อปของ thread ที่
# เรียก — CropWat บนเดสก์ท็อปซ่อนจะไม่ถูกเจอจากตรงนี้ runner.force_reset จัดการ
# ปิด session เดสก์ท็อปซ่อน (ผ่าน process handle ตรงๆ) แยกต่างหากอยู่แล้ว
def force_close_cropwat() -> int:
    """หา CropWat ที่เปิดอยู่ทั้งหมดแล้วสั่งปิดแบบบังคับ คืนจำนวน process ที่ปิด
    สำเร็จ (0 = ไม่เจอ CropWat เปิดอยู่เลย)

    v0.5.29 — ยืนยันจากผู้ใช้จริง: เวอร์ชันแรก (v0.5.25) ใช้ subprocess เรียก
    `taskkill` แต่ผู้ใช้กดปุ่มแล้ว CropWat ไม่ปิดจริง — สาเหตุที่เป็นไปได้มากสุด
    คือ .exe ที่ build ด้วย PyInstaller `--noconsole` (ไม่มี console/stdio ของ
    ตัวเอง) ทำให้ subprocess.run() spawn โปรเซสลูกแบบมี stdio ไม่ครบ ล้มเหลว
    เงียบๆ ด้วย OSError (WinError 6: The handle is invalid) ซึ่งโดน `except
    Exception: continue` ดักไว้แบบไม่มีร่องรอยให้เห็นเลย — เปลี่ยนมาใช้
    win32api.TerminateProcess() ตรงๆ แทน (ผ่าน pywin32 ที่มีอยู่แล้ว) ไม่ต้อง
    พึ่ง subprocess/console เลย เชื่อถือได้กว่าในแอปแบบ windowed"""
    handles = find_windows(
        title_re=controls.MAIN_WINDOW_TITLE_RE,
        class_name=controls.MAIN_WINDOW_CLASS_NAME,
        top_level_only=True,
    )
    pids: set[int] = set()
    for hwnd in handles:
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:  # noqa: BLE001 -- หน้าต่างตายไปพอดีระหว่างเช็ค = ข้าม
            continue
        if pid:
            pids.add(pid)

    PROCESS_TERMINATE = 0x0001
    killed = 0
    for pid in pids:
        handle = None
        try:
            handle = win32api.OpenProcess(PROCESS_TERMINATE, False, pid)
            win32api.TerminateProcess(handle, 1)
            killed += 1
        except Exception:  # noqa: BLE001 -- สิทธิ์ไม่พอ/process ตายไปพอดี = ไม่นับว่าสำเร็จ
            logger.exception("บังคับปิด CropWat (PID %s) ไม่สำเร็จ", pid)
        finally:
            if handle:
                win32api.CloseHandle(handle)
    return killed
