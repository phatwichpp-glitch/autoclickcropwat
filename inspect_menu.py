"""
inspect_menu.py
================
เสริมจาก inspect_cropwat.py — ใช้ดูเมนูจริง (File / Climate/ETo / Rain / Crop ฯลฯ)
ของโปรแกรม Windows แบบเก่า (เช่น CropWat ที่เป็น Delphi VCL) เพราะเมนูแบบนี้
เป็น "native menu" ที่ไม่ใช่ child window เลยไม่โผล่ในผล print_control_identifiers()
ของ inspect_cropwat.py ต้องอ่านผ่าน pywinauto menu API แยกต่างหาก

วิธีใช้: เหมือน inspect_cropwat.py
    backend\\.venv\\Scripts\\python.exe inspect_menu.py
    (เปิด CropWat ค้างไว้ก่อน แล้วเลือกหมายเลขหน้าต่าง CROPWAT)

ผลลัพธ์จะพิมพ์ path ของเมนูทั้งหมด เช่น "File -> Open" "Climate/ETo -> ..."
เอา path พวกนี้ไปใส่ใน cropwat_controls.py ตรง field ...open_menu_path
"""

import sys

from pywinauto import Application, Desktop

# บังคับ stdout/stderr เป็น UTF-8 เสมอ — ไม่งั้นเวลา redirect ผลลัพธ์ลงไฟล์
# (เช่น "> cropwat_menu.txt") บน Windows cmd.exe จะ error เพราะ codepage
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
    windows = Desktop(backend="uia").windows()
    visible_windows = [w for w in windows if w.window_text().strip()]
    eprint("=" * 70)
    eprint("หน้าต่างที่เปิดอยู่บนเครื่องตอนนี้:")
    eprint("=" * 70)
    for i, w in enumerate(visible_windows):
        eprint(f"[{i}] {w.window_text()}")
    return visible_windows


def dump_menu_tree(menu, path=""):
    try:
        items = menu.items()
    except Exception as e:
        print(f"{path}  (อ่าน items ไม่ได้: {e})")
        return
    for item in items:
        try:
            text = item.text() or "(no text)"
        except Exception:
            text = "(อ่าน text ไม่ได้)"
        full_path = f"{path} -> {text}" if path else text
        print(full_path)
        try:
            sub = item.sub_menu()
            if sub is not None:
                dump_menu_tree(sub, full_path)
        except Exception:
            pass


if __name__ == "__main__":
    windows = list_open_windows()
    if not windows:
        eprint("ไม่พบหน้าต่างที่เปิดอยู่เลย ลองเปิด CropWat ก่อนแล้วรันใหม่")
        exit()

    eprint("\nพิมพ์หมายเลข [ ] ของหน้าต่าง CropWat ที่ต้องการสแกน แล้วกด Enter: ", end="")
    choice = input()
    try:
        idx = int(choice.strip())
        target = windows[idx]
    except (ValueError, IndexError):
        eprint("หมายเลขไม่ถูกต้อง ลองรันสคริปต์ใหม่แล้วดูเลขให้ตรงนะครับ")
        exit()

    app = Application(backend="win32").connect(handle=target.handle)
    win = app.window(handle=target.handle)

    print("\n" + "=" * 70)
    print(f"เมนูของหน้าต่าง: {win.window_text()}")
    print("=" * 70)
    try:
        menu = win.menu()
        if menu is None:
            print("หน้าต่างนี้ไม่มี native menu (menu() คืนค่า None)")
        else:
            dump_menu_tree(menu)
    except Exception as e:
        print(f"อ่านเมนูไม่ได้: {e}")
