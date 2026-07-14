"""
launcher.py
===========
Entry point สำหรับ build เป็น .exe เดียวด้วย PyInstaller — คนละอันกับตอน dev
(ตอน dev ใช้ `uvicorn app:app --reload` ตรงๆ ได้ แต่ .exe ต้อง import แบบ object
ตรงๆ ไม่ใช่ string เพราะ frozen bundle resolve module string ไม่ได้เหมือน dev)

ทำ 2 อย่าง: (1) เปิดเบราว์เซอร์แบบ app mode (ไม่มี address bar/แท็บ) เข้า localhost
ให้อัตโนมัติหลังจากรอ server พร้อม (2) รัน uvicorn เอง — ผู้ใช้ปลายทางแค่ดับเบิลคลิก
ไม่ต้องพิมพ์คำสั่งอะไรเลย

สำคัญมาก: import "app", "uvicorn" ฯลฯ ที่ดึง pywinauto/fastapi/pydantic เข้ามาด้วย
**ต้องอยู่ในฟังก์ชันที่ถูก try/except ครอบ ไม่ใช่ import ระดับบนสุดของไฟล์** —
เจอจริงว่าถ้า import พวกนี้ fail (เช่น hidden import ขาดตอน build, DLL หาไม่เจอใน
เครื่องปลายทาง) มันจะ crash ตั้งแต่ก่อนโค้ดใน main() จะได้ทำงานเลย ซึ่งถ้า import
อยู่ระดับบนสุดของไฟล์ จะหลุดพ้นจาก try/except ทั้งหมดที่เขียนไว้ใน main() ทำให้
โปรแกรมปิดตัวเงียบๆ ไม่มีข้อความ error ให้เห็นเลย (อาการ "cmd โผล่มาแล้วดับ")

Build ด้วย:
    pyinstaller --onefile --name CropWatAutoRunner ^
        --add-data "..\\frontend;frontend" ^
        --hidden-import pywinauto --hidden-import win32timezone ^
        launcher.py
(รันจากในโฟลเดอร์ backend/ — ดู build.bat ที่เตรียมไว้ให้)
"""

from __future__ import annotations

import socket
import sys
import traceback

HOST = "127.0.0.1"
PORT = 8000

# ตำแหน่งที่ Edge/Chrome มักติดตั้งอยู่ (เผื่อไม่ได้อยู่ใน PATH) — เช็คแบบนี้ก่อน
# เพราะ Edge ติดตั้งมาให้ในตัวทุกเครื่อง Windows 10/11 อยู่แล้ว โอกาสเจอสูงมาก
_BROWSER_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]


def _wait_for_exit() -> None:
    try:
        input("\nกด Enter เพื่อปิดหน้าต่างนี้...")
    except EOFError:
        pass


def _port_in_use(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        return sock.connect_ex((host, port)) == 0


def _find_app_mode_browser() -> str | None:
    import shutil
    from pathlib import Path

    for path in _BROWSER_CANDIDATES:
        if Path(path).exists():
            return path
    for name in ("msedge", "chrome", "msedge.exe", "chrome.exe"):
        found = shutil.which(name)
        if found:
            return found
    return None


def _open_browser_when_ready() -> None:
    import os
    import subprocess
    import time
    import webbrowser
    from pathlib import Path

    # รอ server เริ่มก่อนค่อยเปิดเบราว์เซอร์ — กันเปิดแล้วเจอ "connection refused"
    time.sleep(1.5)
    url = f"http://{HOST}:{PORT}"
    browser_exe = _find_app_mode_browser()
    if browser_exe:
        # --app=URL เปิดเป็นหน้าต่างเปล่า ไม่มี address bar/แท็บ/ปุ่มเบราว์เซอร์เลย
        # หน้าตาเหมือนโปรแกรม native — ยังเป็นเว็บเหมือนเดิมแค่ไม่โชว์ chrome ของ
        # เบราว์เซอร์ให้เห็น ใช้ --user-data-dir แยกกันไม่ให้ชนกับ Edge/Chrome
        # โปรไฟล์หลักที่ผู้ใช้เปิดอยู่ปกติ
        profile_dir = Path(os.environ.get("TEMP", ".")) / "CropWatAutoRunner-browser-profile"
        try:
            subprocess.Popen([
                browser_exe,
                f"--app={url}",
                "--window-size=1320,880",
                f"--user-data-dir={profile_dir}",
            ])
            return
        except OSError:
            pass  # เปิดแบบ app mode ไม่ได้ → fallback ไปเปิดแท็บเบราว์เซอร์ปกติแทน
    webbrowser.open(url)


def _run_server() -> None:
    """ทุก import ที่ "เสี่ยง fail" (ดึง pywinauto/fastapi/pydantic ฯลฯ เข้ามาด้วย)
    อยู่ในนี้ทั้งหมด เพื่อให้ exception จากตรงนี้ถูก try/except ใน main() จับได้"""
    import logging
    import threading

    import uvicorn

    import app as app_module

    logging.basicConfig(level=logging.INFO)
    threading.Thread(target=_open_browser_when_ready, daemon=True).start()
    print(f"\nกำลังเริ่มระบบที่ http://{HOST}:{PORT} ...")
    print("(เบราว์เซอร์จะเปิดให้อัตโนมัติ — อย่าปิดหน้าต่างนี้ระหว่างใช้งาน)\n")
    # ส่ง app object ตรงๆ (ไม่ใช่ string "app:app") เพราะใน .exe ที่ build แบบ
    # --onefile จะ resolve import string ไม่ได้เหมือนตอนรันด้วย python ปกติ
    uvicorn.run(app_module.app, host=HOST, port=PORT, log_level="info")


def main() -> None:
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
        _run_server()
    except BaseException:
        # ครอบกว้างสุด (BaseException ไม่ใช่แค่ Exception) เพราะ error ที่เจอจริง
        # ระหว่างทดสอบเกิดตอน import ไลบรารีข้างใน ก่อนโค้ด logic จะได้รันเลยด้วยซ้ำ
        # — ต้องมั่นใจว่าไม่ว่า error จะมาจากจุดไหนก็ตาม หน้าต่างจะไม่ปิดตัวเงียบๆ
        print("\n[เกิดข้อผิดพลาด] รายละเอียด:\n")
        traceback.print_exc()
        _wait_for_exit()
        sys.exit(1)

    # มาถึงตรงนี้แปลว่า server ปิดตัวเองแบบปกติ (เช่นกด Ctrl+C) — ค้างไว้เผื่อมี
    # log ท้ายๆ ที่อยากอ่านก่อนหน้าต่างหาย
    _wait_for_exit()


if __name__ == "__main__":
    try:
        main()
    except BaseException:
        # ตาข่ายชั้นสุดท้าย เผื่อ error หลุดออกมาจากจุดที่คาดไม่ถึงจริงๆ (เช่น
        # ระหว่าง parse argument หรือ setup อื่นๆ ก่อนเข้า main() ด้วยซ้ำ)
        print("\n[เกิดข้อผิดพลาดร้ายแรง] รายละเอียด:\n")
        traceback.print_exc()
        _wait_for_exit()
        sys.exit(1)
