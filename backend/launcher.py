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
import os
import socket
import sys
import traceback
from pathlib import Path

HOST = "127.0.0.1"
# override พอร์ตได้ผ่าน env var — ไว้ทดสอบ .exe ตัวใหม่เคียงข้างตัวจริงที่รันอยู่
# (พอร์ตชนกันจะกลายเป็น "เรียกหน้าต่างตัวเดิม" แทนที่จะเปิดตัวใหม่)
PORT = int(os.environ.get("CROPWAT_AUTORUNNER_PORT", "8000"))

# หน้าต่างโปรแกรมแบบ standalone (pywebview/WebView2, v0.10.0) — เดิมเปิดผ่าน
# Edge app-mode ซึ่งไอคอน taskbar เป็นโลโก้ Edge ดูไม่เป็นมืออาชีพ (feedback
# ผู้ใช้จริง) โหมดนี้หน้าต่างเป็นของ process เราเอง ไอคอน = app.ico ที่ฝังใน
# .exe — None = ยังไม่ได้เปิด/ใช้โหมดเบราว์เซอร์ fallback อยู่
_webview_window = None

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
        # Chrome_WidgetWin_1 = หน้าต่าง Chromium (โหมดเบราว์เซอร์ fallback),
        # WindowsForms* = หน้าต่าง standalone webview (v0.10.0)
        if class_buf.value == "Chrome_WidgetWin_1" or class_buf.value.startswith("WindowsForms"):
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
    """ย่อหน้าต่างโปรแกรมลงถาดระบบ — เรียกจาก /api/window/minimize ตอนกดปุ่ม
    "ย่อหน้าต่าง" บนหน้าเว็บ เรียกคืนได้จากเมนู tray icon ("เปิดหน้าต่างโปรแกรม")
    โหมด standalone ใช้ hide() = หายไปอยู่ tray จริงๆ (ไม่ค้างบน taskbar)"""
    w = _webview_window
    if w is not None:
        try:
            w.hide()
            return
        except Exception:  # noqa: BLE001 -- ตกไปทาง fallback ด้านล่าง
            pass
    hwnd = _find_app_window_hwnd()
    if hwnd:
        SW_MINIMIZE = 6
        ctypes.windll.user32.ShowWindow(hwnd, SW_MINIMIZE)


def _launch_app_window() -> None:
    """เปิด/เรียกคืน "หน้าต่างโปรแกรม" — ใช้ทั้งตอนเริ่มโปรแกรม และตอนกดจาก tray
    icon/overlay — โหมด standalone (v0.10.0): show หน้าต่าง webview ตัวเดิมของ
    process นี้ ถ้าไม่ได้ใช้โหมดนั้น fallback เป็นเบราว์เซอร์ app-mode แบบเดิม
    (เช็คก่อนเสมอว่ามีหน้าต่างเปิดค้างอยู่แล้วไหม ถ้ามีแค่ดึงมาข้างหน้า)"""
    import os
    import subprocess
    import webbrowser

    w = _webview_window
    if w is not None:
        try:
            w.show()
            w.restore()
            return
        except Exception:  # noqa: BLE001 -- หน้าต่าง webview มีปัญหา → fallback เบราว์เซอร์
            pass

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

    if _try_webview_mode(app_module, uvicorn):
        return

    # ---- โหมดเบราว์เซอร์ (fallback เดิม — เครื่องที่ไม่มี WebView2 runtime) ----
    if open_browser:
        threading.Thread(target=_open_browser_when_ready, daemon=True).start()
    # ส่ง app object ตรงๆ (ไม่ใช่ string "app:app") เพราะใน .exe ที่ build แบบ
    # --onefile จะ resolve import string ไม่ได้เหมือนตอนรันด้วย python ปกติ
    uvicorn.run(app_module.app, host=HOST, port=PORT, log_level="info", log_config=None)


def _assets_icon_path() -> Path:
    """ตำแหน่ง assets/app.ico — ตอน build เป็น .exe ไฟล์ถูกแตกไปที่ sys._MEIPASS
    (เหมือน _assets_dir ใน overlay.py)"""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)) / "assets" / "app.ico"
    return Path(__file__).parent / "assets" / "app.ico"


def _apply_window_icon() -> None:
    """ตั้งไอคอนหน้าต่าง standalone (title bar + taskbar) เป็นโลโก้โปรแกรม —
    WinForms ของ pywebview ไม่ตั้งให้เอง (เห็นเป็นไอคอนดีฟอลต์) ต้องยิง
    WM_SETICON เองหลังหน้าต่างโชว์แล้ว"""
    try:
        ico = _assets_icon_path()
        if not ico.is_file():
            return
        hwnd = _find_app_window_hwnd()
        if not hwnd:
            return
        user32 = ctypes.windll.user32
        IMAGE_ICON, LR_LOADFROMFILE = 1, 0x10
        WM_SETICON, ICON_SMALL, ICON_BIG = 0x80, 0, 1
        for size, which in ((16, ICON_SMALL), (32, ICON_BIG)):
            h = user32.LoadImageW(None, str(ico), IMAGE_ICON, size, size, LR_LOADFROMFILE)
            if h:
                user32.SendMessageW(hwnd, WM_SETICON, which, h)
    except Exception:  # noqa: BLE001 -- ไอคอนเป็นของเสริม พังแล้วห้ามล้มหน้าต่าง
        pass


def _try_webview_mode(app_module, uvicorn) -> bool:
    """โหมดหน้าต่าง standalone (v0.10.0): รัน uvicorn ใน daemon thread แล้วเปิด
    หน้าต่าง WebView2 ของตัวเองบน main thread (pywebview ต้องการ GUI loop บน
    main thread) — ไอคอน taskbar เป็นโลโก้โปรแกรม (จาก app.ico ที่ฝังใน .exe)
    ไม่ใช่โลโก้ Edge อีกต่อไป

    ปิดหน้าต่าง (X) = ซ่อนลงถาดระบบ โปรแกรมทำงานต่อ (pattern เดียวกับแอปที่มี
    tray icon ทั่วไป) — ปิดจริงทำจากเมนู tray "ปิดโปรแกรม" — คืน False ถ้าเปิด
    โหมดนี้ไม่ได้ (ไม่มี WebView2 runtime ฯลฯ) เพื่อ fallback เป็นเบราว์เซอร์"""
    import logging
    import threading
    import time

    logger = logging.getLogger("launcher")
    try:
        import webview
    except Exception:  # noqa: BLE001 -- import พังในบางเครื่อง = ใช้เบราว์เซอร์แทน
        logger.warning("โหลด pywebview ไม่ได้ — ใช้โหมดเบราว์เซอร์แทน", exc_info=True)
        return False

    threading.Thread(
        target=lambda: uvicorn.run(
            app_module.app, host=HOST, port=PORT, log_level="info", log_config=None
        ),
        daemon=True,
        name="uvicorn-server",
    ).start()

    # รอ server พร้อมก่อนค่อยเปิดหน้าต่าง — กันหน้าขาว "connection refused"
    deadline = time.monotonic() + 20
    while not _port_in_use(HOST, PORT):
        if time.monotonic() > deadline:
            raise RuntimeError("เริ่ม web server ภายในโปรแกรมไม่สำเร็จ (timeout 20 วินาที)")
        time.sleep(0.25)

    global _webview_window
    try:
        win = webview.create_window(
            APP_WINDOW_TITLE,
            f"http://{HOST}:{PORT}",
            width=1320,
            height=880,
            background_color="#0a101e",  # สีพื้นธีมหลัก — กันแฟลชขาวตอนเปิด
        )

        def _on_closing():
            # ปิดหน้าต่าง = ซ่อนลงถาดระบบ (return False = ยกเลิกการปิดจริง)
            try:
                win.hide()
            except Exception:  # noqa: BLE001
                pass
            return False

        def _on_shown(*_args):
            _apply_window_icon()

        win.events.closing += _on_closing
        win.events.shown += _on_shown
        _webview_window = win
        # private_mode=False สำคัญมาก: ค่า default (True) จะล้าง localStorage ทุก
        # ครั้งที่ปิดโปรแกรม → ติ๊ก "ไม่ต้องแสดงคู่มืออีก" ไม่เคยจำ — เก็บ profile
        # ไว้ที่โฟลเดอร์เดียวกับ browser-profile เดิมข้างๆ กัน
        import os

        storage = Path(os.environ.get("LOCALAPPDATA", str(_exe_dir()))) / "CropWatAutoRunner-webview"
        webview.start(private_mode=False, storage_path=str(storage))  # block จนหน้าต่างถูก destroy จริง
    except Exception:  # noqa: BLE001 -- WebView2 runtime ไม่มี/พัง = ใช้เบราว์เซอร์แทน
        _webview_window = None
        logger.warning("เปิดหน้าต่าง standalone ไม่สำเร็จ — ใช้โหมดเบราว์เซอร์แทน", exc_info=True)
        threading.Thread(target=_open_browser_when_ready, daemon=True).start()

    # GUI loop จบ (ปกติไม่เกิดเพราะ closing ถูกยกเลิกเสมอ) แต่ server ต้องอยู่ต่อ
    # — ค้าง main thread ไว้ ให้ปิดโปรแกรมผ่านเมนู tray (os._exit) เท่านั้น
    while True:
        time.sleep(3600)


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


def _activate_existing_instance() -> None:
    """มีตัวเดิมรันอยู่แล้ว (ดับเบิลคลิก .exe ซ้ำ) — ขอให้ process เดิมโชว์
    หน้าต่างของมันเองผ่าน API (v0.10.0: หน้าต่าง standalone เป็นของ process เดิม
    เราสร้าง/ดึงแทนไม่ได้ และถ้าถูกซ่อนลง tray อยู่ EnumWindows ก็มองไม่เห็น) —
    ตัวเดิมเป็นเวอร์ชันเก่าที่ไม่มี endpoint นี้ก็ fallback วิธีหา hwnd แบบเดิม"""
    import urllib.request

    try:
        req = urllib.request.Request(f"http://{HOST}:{PORT}/api/window/show", method="POST")
        with urllib.request.urlopen(req, timeout=3) as resp:
            if resp.status == 200:
                return
    except Exception:  # noqa: BLE001
        pass
    _launch_app_window()


def main() -> None:
    _set_dpi_aware()
    if _port_in_use(HOST, PORT):
        # มีตัวเดิมรันอยู่แล้ว (เช่น ผู้ใช้ปิดหน้าต่างไปเฉยๆ แล้วดับเบิลคลิก .exe
        # ใหม่) — แค่เรียกหน้าต่างโปรแกรมของตัวเดิมขึ้นมาก็พอ ไม่ใช่ error
        _activate_existing_instance()
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
