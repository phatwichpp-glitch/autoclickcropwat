"""
launcher.py
===========
Entry point สำหรับ build เป็น .exe เดียวด้วย PyInstaller — คนละอันกับตอน dev
(ตอน dev ใช้ `uvicorn app:app --reload` ตรงๆ ได้ แต่ .exe ต้อง import แบบ object
ตรงๆ ไม่ใช่ string เพราะ frozen bundle resolve module string ไม่ได้เหมือน dev)

ทำ 2 อย่าง: (1) เปิดเบราว์เซอร์เข้า localhost ให้อัตโนมัติหลังจากรอ server พร้อม
(2) รัน uvicorn เอง — ผู้ใช้ปลายทางแค่ดับเบิลคลิก ไม่ต้องพิมพ์คำสั่งอะไรเลย

สำคัญ: เช็ค port ว่าง + ดัก error ทุกจุดก่อนปิดโปรแกรม แล้ว "ค้างหน้าต่างไว้ให้อ่าน"
(input() รอกด Enter) เพราะปกติ .exe ที่ดับเบิลคลิกจากหน้าต่างจะเป็น process ใหม่ที่
ไม่มี console ผูกอยู่ก่อน — ถ้า error แล้วปิดตัวเองทันทีจะเห็นแค่หน้าต่างดำวาบแล้วหาย
ไป อ่านข้อความ error ไม่ทัน (ปัญหาที่เจอจริงตอนทดสอบ)

Build ด้วย:
    pyinstaller --onefile --name CropWatAutoRunner ^
        --add-data "..\\frontend;frontend" ^
        --hidden-import pywinauto --hidden-import win32timezone ^
        launcher.py
(รันจากในโฟลเดอร์ backend/ — ดู build.bat ที่เตรียมไว้ให้)
"""

from __future__ import annotations

import logging
import socket
import sys
import threading
import time
import webbrowser

import uvicorn

import app as app_module

HOST = "127.0.0.1"
PORT = 8000


def _wait_for_exit() -> None:
    try:
        input("\nกด Enter เพื่อปิดหน้าต่างนี้...")
    except EOFError:
        pass


def _port_in_use(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        return sock.connect_ex((host, port)) == 0


def _open_browser_when_ready() -> None:
    # รอ server เริ่มก่อนค่อยเปิดเบราว์เซอร์ — กันเปิดแล้วเจอ "connection refused"
    time.sleep(1.5)
    webbrowser.open(f"http://{HOST}:{PORT}")


def main() -> None:
    logging.basicConfig(level=logging.INFO)
    print("=" * 60)
    print("CropWat Auto-runner")
    print("=" * 60)

    if _port_in_use(HOST, PORT):
        print(
            f"\n[ผิดพลาด] พอร์ต {PORT} มีโปรแกรมอื่นใช้งานอยู่แล้ว\n"
            "สาเหตุที่เจอบ่อยที่สุด: CropWat Auto-runner เวอร์ชันเก่ายังรันค้างอยู่\n"
            "เบื้องหลัง (ปิดหน้าต่างไปแล้วแต่ process ไม่ตายจริง)\n\n"
            "วิธีแก้: เปิด Task Manager (Ctrl+Shift+Esc) หา 'CropWatAutoRunner.exe'\n"
            "แล้วกด End Task ให้หมดทุกตัว จากนั้นลองเปิดโปรแกรมนี้ใหม่อีกครั้ง"
        )
        _wait_for_exit()
        sys.exit(1)

    try:
        threading.Thread(target=_open_browser_when_ready, daemon=True).start()
        print(f"\nกำลังเริ่มระบบที่ http://{HOST}:{PORT} ...")
        print("(เบราว์เซอร์จะเปิดให้อัตโนมัติ — อย่าปิดหน้าต่างนี้ระหว่างใช้งาน)\n")
        # ส่ง app object ตรงๆ (ไม่ใช่ string "app:app") เพราะใน .exe ที่ build แบบ
        # --onefile จะ resolve import string ไม่ได้เหมือนตอนรันด้วย python ปกติ
        uvicorn.run(app_module.app, host=HOST, port=PORT, log_level="info")
    except Exception:
        logging.exception("เกิดข้อผิดพลาดที่ไม่คาดคิด")
        _wait_for_exit()
        sys.exit(1)

    # มาถึงตรงนี้แปลว่า server ปิดตัวเองแบบปกติ (เช่นกด Ctrl+C) — ค้างไว้เผื่อมี
    # log ท้ายๆ ที่อยากอ่านก่อนหน้าต่างหาย
    _wait_for_exit()


if __name__ == "__main__":
    main()
