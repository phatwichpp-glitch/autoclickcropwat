"""
cropwat_controls.py
====================
ค่าคงที่ทั้งหมดที่อ้างอิงถึง control/เมนู/หน้าต่างจริงของ CropWat 8.0

ทำไมต้องแยกไฟล์นี้ออกมาต่างหาก:
- automation/cropwat_engine.py จะ "ไม่มี string ชื่อ control ฝังอยู่ในโค้ด logic เลย"
  ทุกอย่างอ้างอิงมาจากไฟล์นี้เท่านั้น เวลาต้องแก้ (เช่น เปลี่ยน Windows,
  CropWat คนละ build, ชื่อปุ่มไม่ตรง) แก้ที่เดียวจบ ไม่ต้องไล่หาในฟังก์ชัน
- ตอนนี้ยังไม่รู้ชื่อ control จริงทุกจุด (รอผลจาก inspect_cropwat.py /
  inspect_menu.py) ค่าที่ยังเป็น "None" คือยังต้องกรอกเพิ่ม

สิ่งที่ยืนยันแล้วจากการ inspect จริง (ดู log การสนทนาประกอบ):
- CropWat เป็น MDI app (Delphi/VCL 32-bit) — หน้าต่างย่อยแต่ละโมดูล (Climate/ETo,
  Rain, Crop, Soil, Irrigation Schedule) เป็น MDI child ซ้อนอยู่ใน MDIClient
  ของหน้าต่างหลัก ไม่ใช่หน้าต่างแยก
- แถบไอคอนซ้าย (Climate/ETo, Rain, Crop, Soil, CWR, Crop Pattern, Schedule,
  Scheme) ไม่ใช่ control ที่ pywinauto มองเห็นแยกชิ้น (เป็นปุ่มวาดเอง ไม่มี HWND
  ของตัวเอง) ห้ามพยายามอ้างอิงชื่อ control ของปุ่มพวกนี้ — ใช้วิธี focus ที่ตัว
  MDI child window ด้วย class_name แทน (class_name คงที่ ไม่เปลี่ยนตามไฟล์ที่โหลด
  ต่างจาก title ที่เปลี่ยนตามไฟล์)
- Climate/ETo MDI child window: class_name="TDayEToPMForm"
- Rain MDI child window: class_name="TDayRainForm"
- Crop MDI child window: class_name="TCropForm" — ช่องวันปลูกคือ Edit ตัวเดียวที่
  เป็น class "TMaskEdit" ในหน้าต่างนี้ (ค่าอื่นๆ ในฟอร์มเป็น TEdit/TRealEdit/
  TIntegerEdit หมด ไม่ชนกัน)
- Irrigation Schedule MDI child window: class_name="TCropScheduleform" —
  โผล่ขึ้นมาเองหลังสั่ง Calculations->Irrigation Scheduling ไม่มีเมนูเปิดแยก
- เมนู "File -> Open" เป็นเมนู generic ใช้เปิดไฟล์ให้กับ "MDI child ที่กำลัง
  active อยู่" ตอนนั้น — ต้อง set_focus() ที่ MDI child เป้าหมายก่อนเรียกเมนูนี้
  เสมอ ถึงจะได้ dialog เปิดไฟล์ประเภทที่ถูกต้อง (climate หรือ rain)
- Dialog เปิดไฟล์ (File -> Open) เป็นหน้าต่างแยกต่างหาก ไม่ใช่ลูกของหน้าต่าง
  CROPWAT หลัก — เวลา inspect ต้องเลือกหมายเลขของตัว dialog เอง ไม่ใช่เลขของ
  หน้าต่าง CROPWAT
- คำนวณจริงๆ เป็น 2 ขั้นตอนแยกกันในเมนู Calculations: "Crop Water Requirements"
  ก่อนเสมอ แล้วค่อย "Irrigation Scheduling" (ต้องพึ่งผล CWR)
- ไม่มีเมนู "Export" มีแต่ "File -> Print" ที่ผู้ใช้ยืนยันว่ามีโหมด print-to-file
  จริง (รอ inspect dialog ที่เด้งขึ้นตอนกด Print)

วิธีกรอกค่าที่เหลือ:
1. รัน inspect_cropwat.py / inspect_menu.py ตามคำแนะนำที่ให้ในแชท
2. เอาผลลัพธ์มาเทียบกับ field ที่ยังเป็น None ด้านล่าง แล้วแทนที่ด้วยค่าจริง
3. รัน `python -c "import cropwat_controls; print(cropwat_controls.require_configured())"`
   ต้องได้ [] ถึงจะเริ่มรันจริงได้
"""

from __future__ import annotations

from dataclasses import dataclass, fields


# ---------------------------------------------------------------------------
# หน้าต่างหลักของ CropWat
# ---------------------------------------------------------------------------

# regex ที่ใช้หา title ของหน้าต่างหลัก CropWat ตอน connect ด้วย pywinauto
# ยืนยันจาก inspect_cropwat.py แล้ว: title จริงคือ "CROPWAT - Session: untitled"
# (ส่วนท้ายเปลี่ยนตามชื่อไฟล์ session ที่เปิด) ใช้ regex กว้างๆ ให้ครอบคลุมทุก session
MAIN_WINDOW_TITLE_RE: str | None = r"CROPWAT.*"

# ยืนยันจาก inspect_cropwat.py แล้ว: หน้าต่างหลักจริงมี class นี้ ("TMainForm")
# จำเป็นต้องใช้คู่กับ title_re เสมอ — เจอจากการทดสอบจริงว่า title_re เพียวๆ
# match ได้ 2 หน้าต่าง (ElementAmbiguousError) คาดว่ามีหน้าต่างซ่อน/helper อีกตัว
# ที่ title ขึ้นต้นด้วย "CROPWAT" เหมือนกัน — ระบุ class_name เพิ่มกันความกำกวม
MAIN_WINDOW_CLASS_NAME: str | None = "TMainForm"

# ชื่อ backend ของ pywinauto ที่ใช้ต่อกับโปรแกรม ("uia" หรือ "win32")
# ยืนยันแล้ว: CropWat เป็นแอป Delphi (VCL) เก่า — uia backend ใช้ไม่ได้
# ("UIAWrapper' object has no attribute 'print_control_identifiers'")
# ต้องใช้ win32 backend เท่านั้น
#
# หมายเหตุ: pywinauto เตือนว่า "32-bit application should be automated using
# 32-bit Python" (CropWat เป็นโปรแกรม 32-bit แต่ python ที่ใช้เป็น 64-bit)
# ถ้าเจอปัญหาควบคุมไม่เสถียร (เช่น set_edit_text ไม่ทำงาน, ส่ง message ไม่ถึง)
# ให้ลองสลับไปใช้ Python 32-bit สำหรับรัน backend แทน
PYWINAUTO_BACKEND: str = "win32"


# ---------------------------------------------------------------------------
# ขั้นที่ 1a: หน้า Climate/ETo (เปิดไฟล์ .PED)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ClimateScreenControls:
    # ยืนยันแล้ว: MDI child window ของ Climate/ETo มี class คงที่นี้เสมอ
    # (title เปลี่ยนไปตามไฟล์ที่โหลด เช่น "Daily ETo Penman-Monteith - ...PED")
    window_class_name: str | None = "TDayEToPMForm"
    # เมนูสร้างฟอร์มเปล่า (ยืนยันจาก cropwat_menu.txt) — ใช้เฉพาะตอนที่ยังไม่มี
    # หน้าต่างโมดูลนี้เลย (CropWat เพิ่งเปิดมาเปล่าๆ): แถบไอคอนซ้ายเป็นปุ่มวาดเอง
    # ไม่มี HWND สั่งคลิกแบบ message ไม่ได้ File->New คือทางเดียวที่ทำให้ MDI child
    # โผล่ขึ้นมาผ่านเมนู แล้วค่อย File->Open ไฟล์จริงทับ (prompt "Save changes?"
    # ที่เด้งตอนเปิดทับ ระบบตอบ No ให้เองแล้ว — ดู _answer_no_to_save_prompt)
    new_menu_path: str | None = "File->New->Climate / ETo->Daily ETo Penman Monteith"
    # เมนู generic ที่ใช้เปิดไฟล์ — ต้อง set_focus() ที่ window_class_name ก่อนเรียก
    open_menu_path: str | None = "File->Open"
    # ยืนยันแล้ว: เป็น Windows common file-open dialog มาตรฐาน (class "#32770")
    # เป็นหน้าต่างแยกจาก CROPWAT หลัก ไม่ใช่ MDI child — สังเกตว่า "Files of type"
    # โชว์ "All ETo files" อัตโนมัติ ยืนยันว่า focus ที่ TDayEToPMForm ก่อนแล้วค่อย
    # เรียก File->Open ทำให้ได้ dialog ประเภทที่ถูกต้อง
    file_dialog_title_re: str | None = r"Open"
    # มี Edit control เดียวในทั้ง dialog (ช่อง "File name:") ใช้ class_name จับได้เลย
    file_dialog_filename_field: str | None = "Edit"
    file_dialog_open_button: str | None = "&Open"


CLIMATE_SCREEN = ClimateScreenControls()


# ---------------------------------------------------------------------------
# ขั้นที่ 1b: หน้า Rain (เปิดไฟล์ .CRD)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class RainScreenControls:
    # ยืนยันแล้ว: MDI child window ของ Rain มี class คงที่นี้เสมอ
    window_class_name: str | None = "TDayRainForm"
    # เหตุผลเดียวกับ CLIMATE_SCREEN.new_menu_path (ยืนยันจาก cropwat_menu.txt)
    new_menu_path: str | None = "File->New->Rain->Daily"
    open_menu_path: str | None = "File->Open"
    # เหมือน CLIMATE_SCREEN — Windows common file-open dialog มาตรฐานตัวเดียวกัน
    # (ยังไม่ได้ inspect แยกตอนโฟกัส Rain แต่โครงสร้าง dialog เป็นมาตรฐานเดียวกัน
    # กับ Climate แน่นอน เพราะเป็น Windows GetOpenFileName ตัวเดียวกัน)
    file_dialog_title_re: str | None = r"Open"
    file_dialog_filename_field: str | None = "Edit"
    file_dialog_open_button: str | None = "&Open"


RAIN_SCREEN = RainScreenControls()


# ---------------------------------------------------------------------------
# ขั้นที่ 2: หน้า Crop (ตั้งวันปลูก)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class CropScreenControls:
    # ยืนยันแล้ว: MDI child window ของ Crop มี class คงที่นี้เสมอ
    # (title มีชื่อไฟล์ crop ต่อท้าย เช่น "Dry crop - ...\MAIZE_120.CRO" —
    # แต่ crop file คงที่ทุกปีตาม spec เลยไม่ต้องเปิดไฟล์ใหม่ทุกปี)
    window_class_name: str | None = "TCropForm"
    # ยืนยันแล้ว: ช่องวันปลูกคือ Edit ตัวเดียวใน TCropForm ที่เป็น class "TMaskEdit"
    # (ค่าอื่นในฟอร์มเป็น TEdit/TRealEdit/TIntegerEdit ไม่ชนกัน) จับด้วย class_name
    # ได้เลย ไม่ต้องพึ่ง title เพราะ title คือค่าวันที่ปัจจุบัน (เปลี่ยนทุกปี)
    planting_date_field_class_name: str | None = "TMaskEdit"
    # คาดว่า TMaskEdit จะ commit ค่าเองตอน focus หลุด — ยังไม่ยืนยัน 100%
    # ถ้าทดสอบแล้วไม่ commit ให้เปลี่ยนเป็น key อื่น เช่น "{ENTER}"
    confirm_key: str | None = "{TAB}"

    # "File->New->Crop->Dry crop" สร้างฟอร์มนิยามพืช "เปล่า" (ไม่ใช่ทางลัดเปิดไฟล์
    # — ยืนยันจากการทดสอบจริง) แต่นั่นแหละคือประโยชน์ของมัน (v0.2.0): ใช้ทำให้
    # หน้าต่างโมดูล Crop โผล่ขึ้นมาตอน CropWat เพิ่งเปิดมาเปล่าๆ แล้วค่อยสั่ง
    # File->Open ไฟล์ .CRO จริงทับ — ผู้ใช้ไม่ต้องเปิด crop/soil เองอีกต่อไป
    # (prompt "Save changes?" ตอนเปิดทับ ระบบตอบ No ให้เอง)
    new_menu_path: str | None = "File->New->Crop->Dry crop"
    # Windows common file-open dialog มาตรฐานเดียวกับ Climate/Rain (ยืนยันแล้วว่า
    # โครงสร้าง dialog เหมือนกันทุกจุดที่เคย inspect — ใช้ค่าเดียวกันได้เลย)
    file_dialog_title_re: str | None = r"Open"
    file_dialog_filename_field: str | None = "Edit"
    file_dialog_open_button: str | None = "&Open"


CROP_SCREEN = CropScreenControls()


# ---------------------------------------------------------------------------
# ขั้นที่ 2b: หน้า Soil (เปิดไฟล์ .SOI — คงที่ทุกปีเหมือน Crop)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class SoilScreenControls:
    # ยืนยันแล้วจาก inspect_cropwat.py: MDI child window ของ Soil มี class นี้เสมอ
    window_class_name: str | None = "Tsoilform"
    # เหตุผลเดียวกับ CROP_SCREEN.new_menu_path — สร้างฟอร์มเปล่าเพื่อให้หน้าต่าง
    # โมดูลโผล่ แล้วค่อยเปิดไฟล์ .SOI จริงทับ (ยืนยันเมนูจาก cropwat_menu.txt)
    new_menu_path: str | None = "File->New->Soil"
    file_dialog_title_re: str | None = r"Open"
    file_dialog_filename_field: str | None = "Edit"
    file_dialog_open_button: str | None = "&Open"


SOIL_SCREEN = SoilScreenControls()


# ---------------------------------------------------------------------------
# ขั้นที่ 3: สั่งคำนวณ (Calculate)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class CalculateControls:
    # ยืนยันจาก inspect_menu.py แล้ว: "คำนวณ" จริงๆ เป็น 2 ขั้นตอนแยกกันในเมนู
    # Calculations ต้องรันเรียงลำดับ CWR ก่อนเสมอ แล้วค่อย Irrigation Scheduling
    # (Irrigation Scheduling ต้องมีผล CWR อยู่แล้วถึงจะคำนวณได้)
    crop_water_requirements_menu_path: str | None = "Calculations->Crop Water Requirements"
    irrigation_scheduling_menu_path: str | None = "Calculations->Irrigation Scheduling"
    # หมายเหตุ: เคยมี calculate_timeout_seconds=30 สำหรับ polling loop รอคำนวณ —
    # เอาออกแล้ว (v0.1.11) เพราะ loop นั้นไม่มีทางออกเมื่อสำเร็จ ทำให้เสียเวลา 30 วิ
    # เต็มทุกรอบคำนวณโดยเปล่าประโยชน์ (CropWat คำนวณเสร็จแทบทันที) — ตอนนี้ใช้
    # ERROR_DIALOG.poll_timeout_seconds เช็ค error รอบเดียวสั้นๆ แทน


CALCULATE = CalculateControls()


# ---------------------------------------------------------------------------
# ขั้นที่ 4-5: หน้า Irrigation Schedule + Print (export)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class IrrigationScheduleControls:
    # ยืนยันแล้ว: MDI child window ของ Irrigation Schedule มี class คงที่นี้เสมอ
    # (title คือ "Crop irrigation schedule" คงที่ด้วย ไม่เปลี่ยนตามปี)
    # หน้าต่างนี้โผล่ขึ้นมาเองหลังสั่ง Calculations->Irrigation Scheduling
    # ไม่มีเมนูเปิดแยกต่างหาก
    window_class_name: str | None = "TCropScheduleform"

    # ยืนยันจากผู้ใช้แล้ว (สำคัญ — แก้จากที่เข้าใจผิดตอนแรก): ก่อนพิมพ์ต้องเลือก
    # radio button "Daily soil moisture balance" ใน groupbox "Table format" ของ
    # หน้าต่างนี้ก่อนเสมอ (ไม่ใช่ "Irrigation schedule") เพราะ workflow จริงของ
    # ผู้ใช้ใช้ตาราง Daily soil moisture balance เป็นข้อมูลหลักที่กรอกลง Excel
    table_format_daily_soil_moisture_radio: str | None = "Daily soil moisture balance"

    # ยืนยันจาก inspect_menu.py แล้ว: ไม่มีเมนู "Export" แต่มี "File -> Print"
    # ที่ผู้ใช้ยืนยันว่ามีโหมด print-to-file จริง
    print_menu_path: str | None = "File->Print"

    # ยืนยันแล้วจาก inspect_cropwat.py: "File -> Print" เปิด dialog ตัวเลือกก่อน
    # (class "TPrintForm", title "Print") ไม่ใช่ dialog เลือกไฟล์ตรงๆ — ต้องตั้งค่า
    # ในนี้ก่อนกด OK แล้วค่อยไปเจอ dialog เลือก path จริงอีกที (คาดว่า) — สอง
    # field ด้านล่างเป็นชื่อ dialog/control ของหน้าตัวเลือกนี้
    print_options_dialog_title_re: str | None = r"Print"
    # radio button เลือก "พิมพ์ออกเป็นไฟล์ ASCII" แทนที่จะพิมพ์เข้าเครื่องพิมพ์จริง
    print_options_ascii_file_radio: str | None = "ASCII file (text only)"
    # checkbox คั่นคอลัมน์ด้วย comma — ต้องติ๊กไว้เพื่อให้ได้ output ที่ parse เป็น
    # CSV ได้ง่าย (ไม่งั้นจะได้ fixed-width text แทน)
    print_options_use_commas_checkbox: str | None = "Use commas for column separation in tables"
    # checkbox ใน groupbox "Data to print" ที่ต้องติ๊กเฉพาะ "Irrigation schedule"
    # (ตัวอื่นต้องแน่ใจว่าไม่ติ๊ก ไม่งั้นข้อมูลโมดูลอื่นจะปนมาด้วยตอน parse)
    print_options_irrigation_schedule_checkbox: str | None = "Irrigation schedule"
    print_options_ok_button: str | None = "OK"

    # ยืนยันแล้ว: เป็น Windows common "Save As" dialog มาตรฐาน (class "#32770")
    # เหมือน Open dialog ทุกประการ — output เป็น *.txt (ไม่ใช่ .csv)
    print_save_dialog_title_re: str | None = r"Save As"
    print_save_dialog_filename_field: str | None = "Edit"
    print_save_dialog_save_button: str | None = "&Save"

    # ยืนยันแล้ว: หน้าต่างกราฟ "Irrigation scheduling graph" เป็น MDI child อีกตัว
    # ที่อยู่ใน MDIClient เดียวกัน (class คงที่ ใช้ screenshot ที่ 2 ต่อวันปลูก
    # ที่ทดลอง คู่กับ screenshot ของหน้าต่างนี้เอง)
    graph_window_class_name: str | None = "TCropScheduleGraph"
    # หน้าต่างกราฟ "ไม่ได้เปิดเอง" หลังคำนวณ — ต้องสั่งเปิดผ่านเมนูนี้ก่อน capture
    # (ยืนยันจาก cropwat_menu.txt) ครั้งแรกครั้งเดียวต่อ session พอ หลังจากนั้น
    # หน้าต่างค้างอยู่และอัปเดตตามผลคำนวณล่าสุดเอง
    graph_menu_path: str | None = "Charts->Irrigation Schedule"


IRRIGATION_SCHEDULE = IrrigationScheduleControls()


# ---------------------------------------------------------------------------
# Error / Warning dialogs ที่ CropWat อาจเด้งขึ้นมาระหว่างรัน
# (ใช้ตรวจจับ edge case ตาม spec หัวข้อ "กรณี edge case ที่ต้องรองรับ")
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ErrorDialogControls:
    # ยืนยันแล้ว: dialog error จริงของ CropWat คือ class "TMessageForm" title "Error"
    # (ลองสร้าง error โดยสั่งคำนวณตอนยังไม่มีไฟล์เปิดอยู่) ใส่ "Warning" ไว้ด้วย
    # เผื่อ CropWat ใช้ title นี้กับ dialog เตือนที่ไม่ถึงกับ error (ยังไม่ยืนยัน)
    title_re: str | None = r"Error|Warning"
    # สำคัญ: TMessageForm ไม่มี control แยกสำหรับข้อความ — มีแค่ปุ่ม OK ตัวเดียว
    # เป็นลูกของมัน ข้อความ error วาดตรงบนตัว dialog เอง ไม่ใช่ label แยก เลย
    # "อ่านข้อความละเอียดไม่ได้" ด้วยวิธี child_window ปกติ — ปล่อยเป็นค่าว่าง
    # ("") เพื่อบอก engine ให้ใช้ fallback (บันทึกแค่ว่า "มี error dialog ชื่อ X"
    # แทนข้อความเต็ม) ไม่ใช่ None เพราะ "" ถือว่ากรอกแล้ว (ไม่ block การรัน)
    message_text_control: str | None = ""
    dismiss_button: str | None = "OK"
    # เวลาที่รอเช็คว่ามี dialog error เด้งขึ้นมาไหมหลังแต่ละ step (วินาที) —
    # ลดจาก 2.0 เหลือ 0.8 (v0.2.0): จุดนี้ถูกเรียก ~4 ครั้งต่อ 1 วันปลูก คือ
    # ต้นทุนคงที่ก้อนใหญ่สุดที่เหลืออยู่ (2.0×4 = 8 วิ/วันปลูก) — Delphi ประมวลผล
    # คำสั่งเมนูแล้วเด้ง error แทบทันทีถ้าจะเด้ง 0.8 วิพอ และถ้า error โผล่ช้ากว่า
    # นั้นจริงๆ ก็ยังถูกจับได้ที่จุดเช็คของ step ถัดไปอยู่ดี (ตอบ No/กด OK ให้เสมอ)
    poll_timeout_seconds: float = 0.8


ERROR_DIALOG = ErrorDialogControls()


# ---------------------------------------------------------------------------
# ตัวช่วยเช็คว่ากรอกค่าที่จำเป็นครบหรือยัง ก่อนจะปล่อยให้ engine รันจริง
# ---------------------------------------------------------------------------

def _unfilled_fields(instance) -> list[str]:
    missing = []
    for f in fields(instance):
        value = getattr(instance, f.name)
        if value is None:
            missing.append(f"{type(instance).__name__}.{f.name}")
    return missing


def require_configured() -> list[str]:
    """
    คืน list ของ field ที่ยังไม่ได้กรอก (ยังเป็น None) ทั้งหมด
    ถ้า list ว่าง = กรอกครบพร้อมรันจริงแล้ว
    engine จะเรียกฟังก์ชันนี้ก่อนเริ่มรันเสมอ เพื่อ fail เร็วพร้อม error message ที่ชัดเจน
    แทนที่จะไปพังกลางทางตอน pywinauto หา control ไม่เจอ
    """
    missing: list[str] = []
    if MAIN_WINDOW_TITLE_RE is None:
        missing.append("MAIN_WINDOW_TITLE_RE")
    if MAIN_WINDOW_CLASS_NAME is None:
        missing.append("MAIN_WINDOW_CLASS_NAME")
    missing += _unfilled_fields(CLIMATE_SCREEN)
    missing += _unfilled_fields(RAIN_SCREEN)
    missing += _unfilled_fields(CROP_SCREEN)
    missing += _unfilled_fields(SOIL_SCREEN)
    missing += _unfilled_fields(CALCULATE)
    missing += _unfilled_fields(IRRIGATION_SCHEDULE)
    missing += _unfilled_fields(ERROR_DIALOG)
    return missing
