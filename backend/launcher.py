"""
launcher.py
===========
Entry point สำหรับ build เป็น .exe เดียวด้วย PyInstaller — คนละอันกับตอน dev
(ตอน dev ใช้ `uvicorn app:app --reload` ตรงๆ ได้ แต่ .exe ต้อง import แบบ object
ตรงๆ ไม่ใช่ string เพราะ frozen bundle resolve module string ไม่ได้เหมือน dev)

ทำ 2 อย่าง: (1) เปิดเบราว์เซอร์เข้า localhost ให้อัตโนมัติหลังจากรอ server พร้อม
(2) รัน uvicorn เอง — ผู้ใช้ปลายทางแค่ดับเบิลคลิก ไม่ต้องพิมพ์คำสั่งอะไรเลย

Build ด้วย:
    pyinstaller --onefile --name CropWatAutoRunner ^
        --add-data "..\\frontend;frontend" ^
        --hidden-import pywinauto --hidden-import win32timezone ^
        launcher.py
(รันจากในโฟลเดอร์ backend/ — ดู build.bat ที่เตรียมไว้ให้)
"""

from __future__ import annotations

import logging
import threading
import time
import webbrowser

import uvicorn

import app as app_module

HOST = "127.0.0.1"
PORT = 8000


def _open_browser_when_ready() -> None:
    # รอ server เริ่มก่อนค่อยเปิดเบราว์เซอร์ — กันเปิดแล้วเจอ "connection refused"
    time.sleep(1.5)
    webbrowser.open(f"http://{HOST}:{PORT}")


def main() -> None:
    logging.basicConfig(level=logging.INFO)
    threading.Thread(target=_open_browser_when_ready, daemon=True).start()
    # ส่ง app object ตรงๆ (ไม่ใช่ string "app:app") เพราะใน .exe ที่ build แบบ
    # --onefile จะ resolve import string ไม่ได้เหมือนตอนรันด้วย python ปกติ
    uvicorn.run(app_module.app, host=HOST, port=PORT, log_level="info")


if __name__ == "__main__":
    main()
