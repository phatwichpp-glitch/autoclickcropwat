"""
inspect_cropwat.py
====================
สคริปต์นี้ไว้ "สแกน" หน้าต่างของ CropWat เพื่อดูว่าโปรแกรมมีปุ่ม/เมนู/ช่องกรอกข้อมูล
อะไรบ้าง และแต่ละอันมีชื่อ (control identifier) ว่าอะไร

เราต้องรู้ชื่อพวกนี้ก่อน ถึงจะเขียนสคริปต์อัตโนมัติสั่งคลิก/พิมพ์ให้ CropWat ได้ถูกจุด

วิธีใช้
-------
1. ติดตั้ง Python 3 บนเครื่อง (ถ้ายังไม่มี) จาก https://www.python.org/downloads/
   ตอนติดตั้ง ให้ติ๊ก "Add Python to PATH" ด้วย

2. เปิด Command Prompt แล้วรัน:
       pip install pywinauto

3. เปิดโปรแกรม CropWat ขึ้นมาก่อน (เปิดค้างไว้เฉยๆ ก็ได้ ไม่ต้องกดอะไร)

4. รันสคริปต์นี้:
       python inspect_cropwat.py

5. สคริปต์จะพิมพ์รายชื่อหน้าต่างทั้งหมดที่เปิดอยู่ในเครื่อง ให้เลือกหมายเลข
   ของหน้าต่าง CropWat (ปกติจะเห็นคำว่า CROPWAT หรือ CropWat 8.0 ในชื่อ)

6. มันจะพิมพ์ "ต้นไม้" ของ controls ทั้งหมดในหน้าต่างนั้นออกมา (อาจยาวมาก)
   ให้ copy ข้อความทั้งหมดที่ออกมา (หรือ >> ไปเก็บเป็นไฟล์ตามที่แนะนำด้านล่าง)
   แล้วส่งกลับมาให้ผมดู

TIP: ถ้าอยากได้ผลลัพธ์เป็นไฟล์แทนที่จะพิมพ์ยาวๆ ในหน้าจอ ให้รันแบบนี้แทน:
       python inspect_cropwat.py > cropwat_structure.txt
   แล้วส่งไฟล์ cropwat_structure.txt กลับมาแทน

หมายเหตุ: ให้ทำ "ทีละหน้าจอ" ที่เกี่ยวข้องกับงานเรา เช่น
  - หน้าจอหลัก (เมนูบนสุด)
  - หน้าจอ Climate/Rain (เปิดไฟล์ .CLI/.PEN)
  - หน้าจอ Crop (ตั้งวันปลูก)
  - หน้าจอ Irrigation Schedule (ผลลัพธ์ + ปุ่ม Export/Print)
รันสคริปต์นี้ใหม่ทุกครั้งที่เปลี่ยนไปหน้าจอที่ต้องการสแกน (เพราะ controls จะไม่เหมือนกัน)
"""

import sys

from pywinauto import Desktop

# บังคับ stdout/stderr เป็น UTF-8 เสมอ — ไม่งั้นเวลา redirect ผลลัพธ์ลงไฟล์
# (เช่น "> cropwat_structure.txt") บน Windows cmd.exe จะ error เพราะ codepage
# เริ่มต้นเข้ารหัสภาษาไทยไม่ได้
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def eprint(*args, **kwargs):
    """print ไปที่ stderr เสมอ — ใช้กับข้อความโต้ตอบ (เลือกหน้าต่าง/error) เพื่อให้
    ยังโชว์บนจอได้แม้ผู้ใช้จะ redirect stdout ไปไฟล์ด้วย "> output.txt"
    (ไม่งั้น prompt จะหายไปอยู่ในไฟล์ ทำให้ดูเหมือนสคริปต์ค้าง ทั้งที่จริงรอ input อยู่)"""
    print(*args, file=sys.stderr, **kwargs)


def list_open_windows():
    """แสดงรายชื่อหน้าต่างทั้งหมดที่เปิดอยู่บนเครื่อง"""
    windows = Desktop(backend="uia").windows()
    visible_windows = [w for w in windows if w.window_text().strip()]
    eprint("=" * 70)
    eprint("หน้าต่างที่เปิดอยู่บนเครื่องตอนนี้:")
    eprint("=" * 70)
    for i, w in enumerate(visible_windows):
        eprint(f"[{i}] {w.window_text()}")
    return visible_windows


def dump_window_tree(window):
    """พิมพ์โครงสร้าง controls ทั้งหมดในหน้าต่างที่เลือก (ผลลัพธ์จริงลง stdout)"""
    print("\n" + "=" * 70)
    print(f"โครงสร้างของหน้าต่าง: {window.window_text()}")
    print("=" * 70)
    try:
        window.print_control_identifiers(depth=None)
    except Exception as e:
        eprint(f"เกิดข้อผิดพลาดตอนอ่านโครงสร้างแบบ uia: {e}")
        eprint("กำลังลองใหม่ด้วย backend แบบ win32 (เผื่อโปรแกรมเก่าใช้ backend นี้)...")
        from pywinauto import Application
        app = Application(backend="win32").connect(handle=window.handle)
        app.window(handle=window.handle).print_control_identifiers(depth=None)


if __name__ == "__main__":
    windows = list_open_windows()
    if not windows:
        eprint("ไม่พบหน้าต่างที่เปิดอยู่เลย ลองเปิด CropWat ก่อนแล้วรันใหม่")
        exit()

    eprint("\nพิมพ์หมายเลข [ ] ของหน้าต่าง CropWat ที่ต้องการสแกน แล้วกด Enter: ", end="")
    choice = input()
    try:
        idx = int(choice.strip())
        dump_window_tree(windows[idx])
    except (ValueError, IndexError):
        eprint("หมายเลขไม่ถูกต้อง ลองรันสคริปต์ใหม่แล้วดูเลขให้ตรงนะครับ")
