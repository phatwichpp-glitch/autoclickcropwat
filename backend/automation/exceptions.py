"""Exception ประเภทต่างๆ ที่ automation engine อาจโยนออกมา ใช้แยกแยะสาเหตุความล้มเหลว
ให้ layer ที่สูงกว่า (state/API) รู้ว่าจะรายงานสถานะปีนั้นว่าอย่างไร"""

from __future__ import annotations


class CropWatAutomationError(Exception):
    """base class ของ error ทั้งหมดที่เกิดจากการควบคุม CropWat"""


class ControlsNotConfiguredError(CropWatAutomationError):
    """cropwat_controls.py ยังกรอกไม่ครบ — ยังไม่พร้อมรันจริง"""


class CropWatNotRunningError(CropWatAutomationError):
    """หา process/หน้าต่างของ CropWat ไม่เจอบนเครื่อง"""


class StepTimeoutError(CropWatAutomationError):
    """รอ control หรือผลลัพธ์จาก CropWat นานเกินกำหนด"""


class CropWatReportedError(CropWatAutomationError):
    """CropWat เด้ง error/warning dialog ขึ้นมาระหว่างขั้นตอนใดขั้นตอนหนึ่ง"""
