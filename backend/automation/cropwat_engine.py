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
import time
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

from pywinauto import Application
from pywinauto.findwindows import ElementAmbiguousError, ElementNotFoundError
from pywinauto.timings import TimeoutError as PywinautoTimeoutError

import cropwat_controls as controls
from automation.exceptions import (
    ControlsNotConfiguredError,
    CropWatNotRunningError,
    CropWatReportedError,
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


class CropWatEngine:
    """ห่อ pywinauto Application ของ CropWat ไว้ 1 ตัว ใช้ตลอดการรันหลายปีต่อกัน
    (เปิด CropWat ครั้งเดียว ไม่เปิดใหม่ทุกปี เพื่อความเร็วและลดจุดพัง)"""

    def __init__(self) -> None:
        self.app: Optional[Application] = None
        self.main_window = None

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
        ไม่ใช่ top-level window ของ process ทั้งที่หน้าตาดูเหมือนหน้าต่างแยก"""
        window = self.app.window(class_name=class_name, top_level_only=False)
        window.wait("exists enabled visible ready", timeout=10)
        window.set_focus()
        # เผื่อเวลาให้ Windows ประมวลผลการเปลี่ยน focus จริงๆ ก่อน — เจอจริงว่าถ้า
        # ยิงคำสั่งเมนูต่อทันทีโดยไม่รอเลย บางครั้ง CropWat ยังทำงานกับหน้าต่างเดิม
        # ที่ active อยู่ก่อนหน้า ไม่ใช่ตัวที่เพิ่ง set_focus() ไป
        time.sleep(0.3)
        return window

    def _open_file_via_dialog(self, file_path: Path, dialog_title_re, filename_field, open_button) -> None:
        """ทำ flow เปิดไฟล์ผ่าน Windows-style open dialog ที่เด้งขึ้นมาเป็น
        หน้าต่างแยกต่างหาก (ไม่ใช่ลูกของ CROPWAT หลัก) — ใช้ร่วมกันทั้ง climate/rain

        สำคัญ: ยืนยัน "อ่านค่ากลับ" หลังพิมพ์ path ว่าช่อง filename มีข้อความตรงกับ
        ที่ตั้งใจพิมพ์จริงๆ ก่อนจะกดปุ่ม Open เสมอ — เจอจริงว่าบางครั้งกด Open ไป
        ทั้งที่ช่องยังว่างอยู่ (dialog ยังโหลดไม่เสร็จตอนพิมพ์ หรือ control ที่ set
        ข้อความไม่ใช่ตัวที่ถูกต้อง) ทำให้ CropWat error เพราะไม่มีไฟล์ให้เปิดจริง
        ถ้าอ่านกลับมาไม่ตรง จะลองพิมพ์ใหม่อีกครั้งก่อนจะ fail แบบมีข้อความชัดเจน
        แทนที่จะกด Open ไปทั้งที่รู้อยู่แล้วว่าข้อมูลผิด"""
        if not file_path.exists():
            raise FileNotFoundError(f"ไม่พบไฟล์: {file_path}")

        self.main_window.menu_select("File->Open")

        dialog = self.app.window(title_re=dialog_title_re)
        dialog.wait("exists enabled visible ready", timeout=10)
        time.sleep(0.3)  # เผื่อเวลาให้ dialog พร้อมรับ input จริงๆ ก่อนพิมพ์

        target = str(file_path)
        field = dialog[filename_field]
        for attempt in range(2):
            field.click_input()
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

        dialog[open_button].click_input()

    # ------------------------------------------------------------------
    # Step 1a: เปิดไฟล์ Climate (.PED) ของปีนั้น
    # ------------------------------------------------------------------
    def open_climate_file(self, file_path: Path) -> None:
        self._require_connected()
        cfg = controls.CLIMATE_SCREEN
        self._focus_mdi_child(cfg.window_class_name)
        self._open_file_via_dialog(
            Path(file_path),
            cfg.file_dialog_title_re,
            cfg.file_dialog_filename_field,
            cfg.file_dialog_open_button,
        )
        self._raise_if_error_dialog(f"เปิดไฟล์ climate {file_path}")

    # ------------------------------------------------------------------
    # Step 1b: เปิดไฟล์ Rain (.CRD) ของปีนั้น
    # ------------------------------------------------------------------
    def open_rain_file(self, file_path: Path) -> None:
        self._require_connected()
        cfg = controls.RAIN_SCREEN
        self._focus_mdi_child(cfg.window_class_name)
        self._open_file_via_dialog(
            Path(file_path),
            cfg.file_dialog_title_re,
            cfg.file_dialog_filename_field,
            cfg.file_dialog_open_button,
        )
        self._raise_if_error_dialog(f"เปิดไฟล์ rain {file_path}")

    # ------------------------------------------------------------------
    # ตรวจสอบว่า Crop/Soil เปิดอยู่แล้วหรือยัง — เรียกแค่ "ครั้งเดียวต่อ session"
    # ก่อนเริ่มวนหลายปี ไม่ใช่ทุกปี/ทุกวันปลูก
    #
    # สำคัญ (ยืนยันจากการทดสอบจริงแล้ว): แค่ "focus" หน้าต่างให้เจอเท่านั้น พอ —
    # ห้ามสั่ง File->Open ไฟล์ซ้ำเข้าไปในหน้าต่างที่มันเปิดอยู่แล้ว เพราะ CropWat
    # แจ้ง error กลับมา (คาดว่าไม่ยอมให้เปิดไฟล์เดิมที่ active อยู่ซ้ำ) — ไฟล์
    # crop/soil คงที่ตลอด batch ตาม spec อยู่แล้ว ไม่มีเหตุผลต้องสั่งเปิดใหม่เลย
    # แค่เช็คว่ามีหน้าต่างพร้อมใช้งานจริงก็พอ
    # ------------------------------------------------------------------
    def _verify_module_open(self, window_class_name: str, module_label: str) -> None:
        self._require_connected()
        try:
            self._focus_mdi_child(window_class_name)
        except (ElementNotFoundError, PywinautoTimeoutError) as exc:
            raise CropWatNotRunningError(
                f"ยังไม่ได้เปิดไฟล์ {module_label} ใน CropWat — กรุณาเปิดเอง "
                f"(File → Open) ก่อนกด \"เริ่มรันทั้งหมด\" (ระบบยังไม่รองรับสร้าง/"
                f"เปิดไฟล์ {module_label} อัตโนมัติตั้งแต่ศูนย์)"
            ) from exc

    def ensure_crop_soil_open(self, crop_file: Path, soil_file: Path) -> None:
        """เช็คว่า Crop/Soil เปิดอยู่แล้ว (ต้องเปิดเองใน CropWat ไว้ก่อน) — ไม่ได้
        เปิดไฟล์ใหม่ใดๆ เลย พารามิเตอร์ crop_file/soil_file รับไว้เผื่ออนาคตอยาก
        validate ว่าไฟล์ที่เปิดอยู่ตรงกับที่ตั้งค่าไว้จริงไหม (ยังไม่ได้ทำ)"""
        del crop_file, soil_file  # ยังไม่ได้ใช้ตรวจสอบอะไร แค่กันชื่อไว้เผื่ออนาคต
        self._verify_module_open(controls.CROP_SCREEN.window_class_name, "crop")
        self._verify_module_open(controls.SOIL_SCREEN.window_class_name, "soil")

    # ------------------------------------------------------------------
    # Step 2: ตั้งวันปลูกใน module Crop (crop file เองคงที่ทุกปี ไม่ต้องเปิดใหม่)
    # ------------------------------------------------------------------
    def set_planting_date(self, planting_date: date) -> None:
        self._require_connected()
        cfg = controls.CROP_SCREEN
        crop_window = self._focus_mdi_child(cfg.window_class_name)

        date_str = planting_date.strftime("%d/%m")
        field = crop_window.child_window(class_name=cfg.planting_date_field_class_name)
        field.set_edit_text(date_str)
        if cfg.confirm_key:
            field.type_keys(cfg.confirm_key)

        self._raise_if_error_dialog(f"ตั้งวันปลูก {date_str}")

    # ------------------------------------------------------------------
    # Step 3: สั่งคำนวณ Crop Water Requirements — ยืนยันจาก inspect_menu.py แล้ว
    # ว่าเป็น 2 ขั้นตอนแยกกัน ต้องรันอันนี้ก่อนเสมอ (Irrigation Scheduling ต้องพึ่ง
    # ผลนี้)
    # ------------------------------------------------------------------
    def calculate(self) -> None:
        self._require_connected()
        cfg = controls.CALCULATE
        self.main_window.menu_select(cfg.crop_water_requirements_menu_path)

        deadline = time.monotonic() + cfg.calculate_timeout_seconds
        while time.monotonic() < deadline:
            error_message = self._poll_error_dialog()
            if error_message:
                raise CropWatReportedError(
                    f"CropWat แจ้ง error ระหว่างคำนวณ Crop Water Requirements: {error_message}"
                )
            time.sleep(0.5)

    # ------------------------------------------------------------------
    # Step 4: สั่งคำนวณ Irrigation Scheduling — นี่คือสิ่งที่ "เปิดหน้า Irrigation
    # Schedule" จริงๆ ในเมนู ไม่มีคำสั่ง "เปิดหน้า" แยกต่างหาก การสั่งคำนวณนี้
    # จะทำให้หน้าต่างผลลัพธ์ (class TCropScheduleform) โผล่ขึ้นมาเป็นผลพลอยได้เลย
    # ------------------------------------------------------------------
    def open_irrigation_schedule(self) -> None:
        self._require_connected()
        cfg = controls.CALCULATE
        self.main_window.menu_select(cfg.irrigation_scheduling_menu_path)

        deadline = time.monotonic() + cfg.calculate_timeout_seconds
        while time.monotonic() < deadline:
            error_message = self._poll_error_dialog()
            if error_message:
                raise CropWatReportedError(
                    f"CropWat แจ้ง error ระหว่างคำนวณ Irrigation Scheduling: {error_message}"
                )
            time.sleep(0.5)

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

        schedule_window = self._focus_mdi_child(cfg.window_class_name)
        # ยืนยันจากผู้ใช้แล้ว: ต้องสลับไปดู "Daily soil moisture balance" ก่อนพิมพ์เสมอ
        schedule_window[cfg.table_format_daily_soil_moisture_radio].click_input()

        self.main_window.menu_select(cfg.print_menu_path)

        options = self.app.window(title_re=cfg.print_options_dialog_title_re)
        options.wait("exists enabled visible ready", timeout=10)
        options[cfg.print_options_ascii_file_radio].click_input()
        commas_checkbox = options[cfg.print_options_use_commas_checkbox].wrapper_object()
        if not commas_checkbox.get_check_state():
            commas_checkbox.click_input()
        irrigation_checkbox = options[cfg.print_options_irrigation_schedule_checkbox].wrapper_object()
        if not irrigation_checkbox.get_check_state():
            irrigation_checkbox.click_input()
        options[cfg.print_options_ok_button].click_input()

        save_dialog = self.app.window(title_re=cfg.print_save_dialog_title_re)
        save_dialog.wait("exists enabled visible ready", timeout=10)
        save_dialog[cfg.print_save_dialog_filename_field].set_edit_text(str(target_file))
        save_dialog[cfg.print_save_dialog_save_button].click_input()

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

        self._focus_mdi_child(cfg.window_class_name)
        schedule_path = screenshot_dir / f"{stem}_schedule.png"
        self.main_window.capture_as_image().save(schedule_path)

        self._focus_mdi_child(cfg.graph_window_class_name)
        graph_path = screenshot_dir / f"{stem}_graph.png"
        self.main_window.capture_as_image().save(graph_path)

        return schedule_path, graph_path

    # ------------------------------------------------------------------
    # ตรวจจับ error/warning dialog (ใช้ทั้งเช็ค inline และ poll หลัง calculate)
    # ------------------------------------------------------------------
    def _poll_error_dialog(self) -> Optional[str]:
        cfg = controls.ERROR_DIALOG
        try:
            dialog = self.app.window(title_re=cfg.title_re)
            if not dialog.exists(timeout=0):
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
                dialog[cfg.dismiss_button].click_input()
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
        cfg = controls.ERROR_DIALOG
        time.sleep(cfg.poll_timeout_seconds)
        message = self._poll_error_dialog()
        if message:
            raise CropWatReportedError(f"CropWat แจ้ง error ระหว่าง {context}: {message}")

    @staticmethod
    def _wait_for_file(path: Path, timeout: float) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if path.exists() and path.stat().st_size > 0:
                return True
            time.sleep(0.5)
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
    ) -> YearRunResult:
        try:
            self.open_climate_file(climate_file)
            self.open_rain_file(rain_file)
        except Exception as exc:  # noqa: BLE001 -- กันไม่ให้ 1 ปี error ล้มทั้ง batch
            logger.warning("ปี %s เปิดไฟล์ climate/rain ไม่สำเร็จ: %s", year, exc)
            return YearRunResult(
                year=year, ok=False, error_message=f"เปิดไฟล์ climate/rain ไม่สำเร็จ: {exc}"
            )

        candidates = [
            self.run_candidate_planting_date(year, task, export_dir, screenshot_dir)
            for task in tasks
        ]
        failed = [c for c in candidates if not c.ok]
        ok = not failed
        error_message = (
            None
            if ok
            else f"{len(failed)}/{len(candidates)} วันปลูกที่ทดลองล้มเหลว "
            f"(เช่น {failed[0].planting_date:%d/%m}: {failed[0].error_message})"
        )
        return YearRunResult(year=year, ok=ok, candidates=candidates, error_message=error_message)
