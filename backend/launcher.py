"""
launcher.py
===========
Entry point สำหรับ build เป็น .exe เดียวด้วย PyInstaller — คนละอันกับตอน dev
(ตอน dev ใช้ `uvicorn app:app` ตรงๆ ได้ แต่ .exe ต้อง import แบบ object ตรงๆ
ไม่ใช่ string เพราะ frozen bundle resolve module string ไม่ได้เหมือน dev)

v0.3.0 — build ด้วย --noconsole แล้ว (ไม่มีหน้าต่าง console ดำๆ อีกต่อไป ให้
ความรู้สึกเหมือนโปรแกรมทั่วไป) ผลที่ตามมาที่โค้ดนี้ต้องรองรับ:
- print() ไปที่ null เงียบๆ / input() ใช้ไม่ได้ → ห้ามพึ่ง console ในการแจ้ง
  error กับผู้ใช้ ใช้ MessageBox (ctypes) แทน + เขียน log ลงไฟล์ข้างตัว .exe
- การ "ปิดโปรแกรม" ย้ายไปอยู่ที่ปุ่ม ✕ บนแถบ overlay (ดู overlay.py) เพราะ
  ไม่มี console ให้ปิดแล้ว
- ดับเบิลคลิก .exe ซ้ำตอนที่ตัวเก่ายังรันอยู่ = เปิด "หน้าต่างโปรแกรม" ของตัว
  เดิมขึ้นมาใหม่เฉยๆ (ไม่ใช่ error) — พฤติกรรมเดียวกับโปรแกรมทั่วไป

สำคัญมาก: import "app", "uvicorn" ฯลฯ ที่ดึง pywinauto/fastapi/pydantic เข้ามาด้วย
**ต้องอยู่ในฟังก์ชันที่ถูก try/except ครอบ ไม่ใช่ import ระดับบนสุดของไฟล์** —
เจอจริงว่าถ้า import พวกนี้ fail (เช่น hidden import ขาดตอน build, DLL หาไม่เจอใน
เครื่องปลายทาง) มันจะ crash ตั้งแต่ก่อนโค้ดใน main() จะได้ทำงานเลย ซึ่งถ้า import
อยู่ระดับบนสุดของไฟล์ จะหลุดพ้นจาก try/except ทั้งหมดที่เขียนไว้ใน main()

Build: ดู build.bat (รันจากในโฟลเดอร์ backend/)
"""

from __future__ import annotations

import ctypes
import socket
import sys
import traceback
from pathlib import Path

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


def _exe_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def _error_log_path() -> Path:
    return _exe_dir() / "CropWatAutoRunner-error.log"


def _message_box(title: str, text: str, *, error: bool = True) -> None:
    """แจ้งผู้ใช้ผ่าน MessageBox ของ Windows — ช่องทางเดียวที่เหลือหลังตัด console
    ทิ้ง (MB_ICONERROR=0x10, MB_ICONINFORMATION=0x40, MB_TOPMOST=0x40000)"""
    icon = 0x10 if error else 0x40
    ctypes.windll.user32.MessageBoxW(None, text, title, icon | 0x40000)


def _report_fatal(context: str) -> None:
    """เขียน traceback ลงไฟล์ข้างตัว .exe + เด้ง MessageBox บอกผู้ใช้ว่าดูได้ที่ไหน"""
    detail = traceback.format_exc()
    log_path = _error_log_path()
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"\n===== {context} =====\n{detail}\n")
    except OSError:
        pass
    last_line = detail.strip().splitlines()[-1] if detail.strip() else "(ไม่มี)"
    _message_box(
        "CropWat Auto-runner — เกิดข้อผิดพลาด",
        f"{context}\n\nรายละเอียดถูกบันทึกไว้ที่:\n{log_path}\n\nสรุป error:\n{last_line}",
    )


def _port_in_use(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        return sock.connect_ex((host, port)) == 0


def _find_app_mode_browser() -> str | None:
    import shutil

    for path in _BROWSER_CANDIDATES:
        if Path(path).exists():
            return path
    for name in ("msedge", "chrome", "msedge.exe", "chrome.exe"):
        found = shutil.which(name)
        if found:
            return found
    return None


# ต้องตรงกับ <title> ใน frontend/index.html เป๊ะ — Chrome/Edge --app mode ใช้
# page title เป็น window title เสมอ ใช้จับหน้าต่างที่เปิดค้างอยู่แล้ว
APP_WINDOW_TITLE = "CropWat Auto-runner"


def _find_app_window_hwnd() -> int | None:
    """หา window handle ของ "หน้าต่างโปรแกรม" ที่เปิดค้างอยู่แล้ว (ถ้ามี) — v0.5.14
    บั๊กที่เจอจากผู้ใช้จริง: กดเปิดหน้าต่างจาก tray icon/overlay แล้วได้ browser
    ใหม่ซ้อนทุกครั้ง ไม่เคยลิงก์กับหน้าต่างเดิมที่เปิดอยู่แล้ว เพราะเดิมเรียก
    webbrowser.open() ตรงๆ ซึ่งไม่เช็คอะไรเลย เปิดแท็บ/หน้าต่างใหม่เสมอ"""
    user32 = ctypes.windll.user32
    found: list[int] = []

    def _cb(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        if buf.value != APP_WINDOW_TITLE:
            return True
        class_buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_buf, 256)
        if class_buf.value == "Chrome_WidgetWin_1":  # class ของหน้าต่าง Chromium (Edge/Chrome)
            found.append(hwnd)
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)
    user32.EnumWindows(WNDENUMPROC(_cb), 0)
    return found[0] if found else None


def _bring_to_front(hwnd: int) -> None:
    """ดึงหน้าต่างที่เปิดอยู่แล้วมาข้างหน้า — Windows บล็อก SetForegroundWindow
    เฉยๆ จาก process ที่ไม่ได้ active อยู่ก่อน (มักได้แค่ไอคอนกระพริบที่ taskbar
    ไม่ได้ดึงมาข้างหน้าจริง) ต้อง AttachThreadInput ก่อนเสมอ — เทคนิคเดียวกับที่
    ยืนยันแล้วว่าได้ผลจริงใน automation/cropwat_engine.py::_real_set_focus"""
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    SW_RESTORE = 9
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)
    target_tid = user32.GetWindowThreadProcessId(hwnd, None)
    cur_tid = kernel32.GetCurrentThreadId()
    attached = False
    try:
        if target_tid and target_tid != cur_tid:
            attached = bool(user32.AttachThreadInput(cur_tid, target_tid, True))
        user32.SetForegroundWindow(hwnd)
        user32.BringWindowToTop(hwnd)
    finally:
        if attached:
            user32.AttachThreadInput(cur_tid, target_tid, False)


def _minimize_app_window() -> None:
    """ย่อหน้าต่างโปรแกรมลง Tray (v0.8.0) — เรียกจาก /api/window/minimize ตอนกด
    ปุ่ม "ย่อไปที่ Tray" บนหน้าเว็บ เรียกคืนได้จากเมนู tray icon ("เปิดหน้าต่าง
    โปรแกรม") ซึ่งใช้ _bring_to_front (มี SW_RESTORE ในตัวอยู่แล้วถ้า minimize ไว้)"""
    hwnd = _find_app_window_hwnd()
    if hwnd:
        SW_MINIMIZE = 6
        ctypes.windll.user32.ShowWindow(hwnd, SW_MINIMIZE)


def _launch_app_window() -> None:
    """เปิด "หน้าต่างโปรแกรม" (เบราว์เซอร์แบบ app mode ไม่มี address bar/แท็บ) —
    ใช้ทั้งตอนเริ่มโปรแกรม, ตอนดับเบิลคลิก .exe ซ้ำ, และตอนกดจาก tray icon/overlay
    เพื่อเรียกหน้าต่างกลับมา — เช็คก่อนเสมอว่ามีหน้าต่างเปิดค้างอยู่แล้วไหม ถ้ามี
    แค่ดึงมาข้างหน้า ไม่เปิดซ้อนใหม่"""
    import os
    import subprocess
    import webbrowser

    existing = _find_app_window_hwnd()
    if existing:
        _bring_to_front(existing)
        return

    url = f"http://{HOST}:{PORT}"
    browser_exe = _find_app_mode_browser()
    if browser_exe:
        # --user-data-dir แยกโปรไฟล์ ไม่ชนกับ Edge/Chrome หลักที่ผู้ใช้เปิดปกติ
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
            pass  # เปิดแบบ app mode ไม่ได้ → fallback เปิดแท็บเบราว์เซอร์ปกติแทน
    webbrowser.open(url)


def _open_browser_when_ready() -> None:
    import time

    # รอ server เริ่มก่อนค่อยเปิดหน้าต่าง — กันเปิดแล้วเจอ "connection refused"
    time.sleep(1.5)
    _launch_app_window()


def _run_server(open_browser: bool = True) -> None:
    """ทุก import ที่ "เสี่ยง fail" (ดึง pywinauto/fastapi/pydantic ฯลฯ เข้ามาด้วย)
    อยู่ในนี้ทั้งหมด เพื่อให้ exception จากตรงนี้ถูก try/except ใน main() จับได้"""
    import logging
    import threading

    import uvicorn

    # ไม่มี console ให้ดู log แล้ว — เขียนลงไฟล์ข้างตัว .exe แทน (เวลาผู้ใช้เจอ
    # ปัญหา ขอไฟล์ CropWatAutoRunner.log มาดูได้เลย)
    if getattr(sys, "frozen", False):
        logging.basicConfig(
            filename=str(_exe_dir() / "CropWatAutoRunner.log"),
            filemode="w",
            level=logging.INFO,
            format="%(asctime)s %(name)s %(levelname)s: %(message)s",
            encoding="utf-8",
        )
    else:
        logging.basicConfig(level=logging.INFO)

    import app as app_module

    if open_browser:
        threading.Thread(target=_open_browser_when_ready, daemon=True).start()
    # ส่ง app object ตรงๆ (ไม่ใช่ string "app:app") เพราะใน .exe ที่ build แบบ
    # --onefile จะ resolve import string ไม่ได้เหมือนตอนรันด้วย python ปกติ
    uvicorn.run(app_module.app, host=HOST, port=PORT, log_level="info", log_config=None)


def _set_dpi_aware() -> None:
    """ประกาศว่า process นี้เป็น "DPI-aware" ก่อนสร้างหน้าต่างใดๆ เลย — v0.5.16
    บั๊กที่เจอจากผู้ใช้จริง: overlay (tkinter) ขึ้นไม่เต็มข้อความบนจอความละเอียด
    สูง เพราะไม่ประกาศ DPI awareness มาก่อน Windows เลยเดา DPI ให้ (system DPI
    aware อัตโนมัติในบางเครื่อง) ทำให้ Tk รู้ DPI จริงแล้วขยายฟอนต์ตาม แต่ขนาด
    หน้าต่าง (pixel ตายตัว) ไม่ขยายตาม ข้อความเลยล้นกรอบ — ประกาศเองให้ชัดเจนที่
    Per-Monitor V2 (แม่นสุด, Windows 10 1703+) พร้อม fallback ไล่ลงมาเผื่อรันบน
    Windows รุ่นเก่า ต้องเรียกตัวนี้ก่อนสร้างหน้าต่างแรกเสมอ (ดู overlay.py ที่
    คำนวณ scale factor จาก DPI จริงมาขยายขนาด/padding ให้พอดีสัดส่วนอีกที)"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
        return
    except Exception:  # noqa: BLE001
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:  # noqa: BLE001
        pass


def main() -> None:
    _set_dpi_aware()
    if _port_in_use(HOST, PORT):
        # มีตัวเดิมรันอยู่แล้ว (เช่น ผู้ใช้ปิดหน้าต่างไปเฉยๆ แล้วดับเบิลคลิก .exe
        # ใหม่) — แค่เรียกหน้าต่างโปรแกรมของตัวเดิมขึ้นมาก็พอ ไม่ใช่ error
        _launch_app_window()
        sys.exit(0)

    # "--updated" = เปิดหลังอัปเดตอัตโนมัติ (ดู updater.py) — หน้าต่างเดิมของ
    # ผู้ใช้ยังเปิดอยู่และจะ reload ตัวเองเมื่อ backend กลับมา ไม่ต้องเปิดใหม่ซ้ำ
    # (ยืนยันจากผู้ใช้: เปิดซ้ำจะได้ 2 หน้าต่างซ้อนกัน งงว่าอันไหนจริง)
    try:
        _run_server(open_browser="--updated" not in sys.argv)
    except BaseException:
        # ครอบกว้างสุด (BaseException) เพราะ error ที่เจอจริงระหว่างทดสอบเกิดตอน
        # import ไลบรารีข้างใน ก่อนโค้ด logic จะได้รันเลยด้วยซ้ำ
        _report_fatal("เริ่มระบบไม่สำเร็จ")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except BaseException:
        # ตาข่ายชั้นสุดท้าย เผื่อ error หลุดจากจุดที่คาดไม่ถึงจริงๆ
        _report_fatal("เกิดข้อผิดพลาดร้ายแรง")
        sys.exit(1)
