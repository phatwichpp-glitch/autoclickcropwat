"""
overlay.py
==========
3 อย่างที่ทำให้ใช้งานได้โดยไม่ต้องเปิดหน้าเว็บเลย:

1. **System tray icon** (v0.5.13) — สัญญาณถาวรว่า "โปรแกรมทำงานอยู่เบื้องหลัง"
   (ยืนยันจากผู้ใช้: ปิดหน้าต่างเว็บแล้วงงว่าโปรแกรมยังทำงานอยู่ไหม เพราะไม่มี
   อะไรบอก) ไอคอนเปลี่ยนเป็นจุดเขียวตอนกำลังรัน คลิกซ้าย/ดับเบิลคลิกเปิดหน้าต่าง
   โปรแกรม คลิกขวาเปิดเมนู (เปิดหน้าต่าง/เริ่มรัน/หยุด/ปิดโปรแกรม)

2. **แถบ progress ลอย (overlay)** — หน้าต่างเล็กๆ ไร้ขอบ อยู่บนสุดเสมอ (topmost)
   มุมขวาล่างจอ ลากย้ายได้ แสดง progress ระดับ "วันปลูก" แบบ real-time — แก้ปัญหา
   "เปิด CropWat บังจอแล้วดู progress ไม่ได้" เพราะ overlay ลอยทับทุกหน้าต่าง
   รวมถึง CropWat ด้วย
   v0.5.13: เดิม overlay ค้างอยู่ตลอดเวลาแม้ไม่ได้รันอะไร ทำให้สับสนกับ tray icon
   ว่าอันไหนคือสัญญาณ "ทำงานอยู่" ตัวจริง — เปลี่ยนให้ overlay โผล่เฉพาะตอนกำลังรัน/
   กำลังหยุด/เพิ่งจบ (ค้างไว้ 15 วิให้เห็นผลสรุป) แล้วซ่อนตัวเอง ส่วน tray icon
   เป็นสัญญาณ "ทำงานอยู่เบื้องหลัง" เพียงหนึ่งเดียวที่ค้างตลอดเวลา
   v0.8.0: ผู้ใช้ feedback ว่า overlay ดูเกะกะกว่า tray — เพิ่มปุ่มซ่อน overlay
   เอง (🗕) เรียกคืนได้จากเมนู tray ("แสดงแถบ Progress ลอย") และเพิ่ม tooltip
   ของ tray icon ให้บอกปี/%/ETA ระหว่างรัน ให้ tray เป็นช่องทางดู progress ที่
   ครบเทียบเท่า overlay ได้โดยไม่ต้องเปิดหน้าต่างไหนเลย

3. **Global hotkeys** — Ctrl+Alt+F9 เริ่มรัน / Ctrl+Alt+F10 หยุด กดได้จากทุกที่
   แม้ CropWat หรือโปรแกรมอื่นกำลัง focus อยู่ (RegisterHotKey ระดับ OS)

ใช้ tkinter (มากับ Python อยู่แล้ว) + pystray (tray icon) รันใน thread แยกของ
ตัวเอง — กติกาสำคัญของ tkinter: ทุกคำสั่ง tk ต้องมาจาก thread เดียวกับที่สร้าง
root เท่านั้น ดังนั้น thread อื่น (hotkey/engine/tray) ห้ามแตะ widget ตรงๆ สื่อสาร
ผ่าน threading.Event / run_state แล้วให้ refresh loop (root.after) ฝั่ง tk อ่านเอาเอง
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import logging
import os
import threading
import time
import webbrowser
from pathlib import Path

logger = logging.getLogger("overlay")

URL = "http://127.0.0.1:8000"
# grace period ที่ overlay ยังค้างโชว์ผลสรุป "จบแล้ว X/Y" หลังรันเสร็จ ก่อนซ่อนตัวเอง
_DONE_GRACE_SECONDS = 15

# ตั้งโดย engine ตอนกำลัง capture screenshot — overlay ต้องซ่อนตัวชั่วคราว ไม่งั้น
# ตัวเองจะติดไปในภาพ (capture ถ่ายจากพิกเซลจริงบนจอ) — ดู capture_screenshots
capture_pause = threading.Event()

# v0.8.0 — ผู้ใช้ให้ feedback ว่า overlay ลอยดูเกะกะกว่า tray icon ("ดู Progress
# ผ่าน System Tray จะดูไม่เกะกะเท่า Overlay") — เพิ่มปุ่มซ่อน overlay เอง (บนแถบ)
# + toggle จากเมนู tray (ดู _toggle_overlay) ตั้ง set() แล้ว overlay ซ่อนตัวเองใน
# รอบ refresh ถัดไปแม้กำลังรันอยู่ ตั้งใจไม่ persist ข้ามการเปิดโปรแกรมใหม่ (ค่า
# เริ่มต้นทุกครั้งคือ "แสดง")
overlay_hidden = threading.Event()


def _assets_dir() -> Path:
    """เหมือน _frontend_dir() ใน app.py — ตอน build เป็น .exe ไฟล์ assets/ ที่
    bundle ไปด้วย (--add-data "assets;assets") จะถูกแตกไปไว้ที่ sys._MEIPASS"""
    import sys

    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)) / "assets"
    return Path(__file__).parent / "assets"


def _format_eta(seconds) -> str:
    """เหมือน formatEta() ใน frontend/app.js — จัดวินาที -> ข้อความไทยอ่านง่าย"""
    if seconds is None or seconds < 1:
        return ""
    total_min = round(seconds / 60)
    if total_min < 1:
        return "เหลือ <1 นาที"
    h, m = divmod(total_min, 60)
    if h > 0:
        return f"เหลือ {h}ชม.{m}น." if m else f"เหลือ {h}ชม."
    return f"เหลือ {m}น."


def _open_app_window() -> None:
    """เปิด/ดึงหน้าต่างโปรแกรมมาข้างหน้า — v0.5.14 บั๊กที่เจอจากผู้ใช้จริง: ปุ่ม
    ⚙ ของ overlay และเมนู tray icon เดิมเรียก webbrowser.open() ตรงๆ ซึ่งไม่เช็ค
    อะไรเลย เปิด browser/แท็บใหม่ซ้อนทุกครั้งไม่เคยลิงก์กับหน้าต่างเดิม — ใช้
    launcher._launch_app_window() ตัวเดียวกับตอนเริ่มโปรแกรมแทน (เช็คก่อนว่ามี
    หน้าต่างเปิดค้างอยู่แล้วไหม มีก็แค่ดึงมาข้างหน้า)"""
    try:
        from launcher import _launch_app_window

        _launch_app_window()
    except Exception:  # noqa: BLE001 -- import/เรียกพัง ยังเปิดแท็บใหม่ได้เป็น fallback
        logger.exception("เปิดหน้าต่างโปรแกรมผ่าน launcher ไม่สำเร็จ — fallback เปิดแท็บใหม่")
        webbrowser.open(URL)

# --- global hotkeys -------------------------------------------------------
HOTKEY_START_ID = 1
HOTKEY_STOP_ID = 2
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
VK_F9 = 0x78
VK_F10 = 0x79
WM_HOTKEY = 0x0312

HOTKEY_HINT = "Ctrl+Alt+F9 เริ่ม · Ctrl+Alt+F10 หยุด"


def _start_run() -> None:
    import runner

    if not runner.start_default_run():
        logger.warning("เริ่มรันไม่ได้ (กำลังรันอยู่แล้ว หรือยังไม่ได้ตั้งค่าโฟลเดอร์ — กด ⚙)")


def _stop_run() -> None:
    from state import run_state

    run_state.request_stop()


def _quit_app() -> None:
    """ปิดโปรแกรมทั้งหมด — ใช้ MessageBoxW ดิบแทน tkinter.messagebox เพราะเรียก
    จาก thread ของ tray icon (ไม่ใช่ thread ของ tk root) ได้โดยไม่ต้องพึ่ง tk เลย
    (os._exit เพราะ uvicorn/engine/overlay/tray เป็น daemon thread ทั้งหมด ปิดตรงๆ
    ได้ ไม่มี state ที่ต้อง flush)"""
    import os

    from state import run_state

    snap = run_state.snapshot()
    if snap.overall_state.value == "running":
        MB_YESNO = 0x4
        MB_ICONQUESTION = 0x20
        IDYES = 6
        answer = ctypes.windll.user32.MessageBoxW(
            None,
            "กำลังรันอยู่ — หยุดและปิดโปรแกรมเลยหรือไม่?\n"
            "(ไฟล์ .txt ที่เสร็จแล้วอยู่ครบ เปิดใหม่แล้วรันต่อได้)",
            "CropWat Auto-runner",
            MB_YESNO | MB_ICONQUESTION,
        )
        if answer != IDYES:
            return

    logger.info("ผู้ใช้สั่งปิดโปรแกรมจาก tray icon")
    os._exit(0)


# --- system tray icon -------------------------------------------------------
# v0.5.13: สัญญาณถาวรเพียงหนึ่งเดียวว่า "โปรแกรมทำงานอยู่เบื้องหลัง" — ก่อนหน้านี้
# มีแต่ overlay ที่ค้างอยู่ตลอด ทำให้ผู้ใช้สับสนว่าปิดหน้าต่างเว็บไปแล้วโปรแกรมยัง
# ทำงานอยู่ไหม (ไม่มีอะไรบอกชัดๆ) — ไอคอนถาดระบบเป็น pattern มาตรฐานของ Windows
# ที่คนคุ้นเคยอยู่แล้ว: มีไอคอน = โปรแกรมมีชีวิตอยู่ ไม่มี = ปิดแล้ว
def _tray_icon_image(running: bool):
    """โหลด assets/app.ico (โลโก้จริงของโปรแกรม) ครั้งเดียว แล้วแปะจุดเขียวเล็กๆ
    มุมล่างขวาตอนกำลังรัน — ให้เห็นสถานะจากไอคอนได้ทันทีโดยไม่ต้อง hover ดู tooltip"""
    from PIL import Image, ImageDraw

    icon_path = _assets_dir() / "app.ico"
    try:
        base = Image.open(icon_path).convert("RGBA")
    except OSError:
        # ไม่เจอไฟล์ (ผิดปกติ แต่ไม่ควรทำให้ tray icon หายไปทั้งฟีเจอร์) — ใช้
        # สี่เหลี่ยมสีพื้นแทนเป็น fallback
        base = Image.new("RGBA", (64, 64), (28, 34, 48, 255))
    if not running:
        return base
    img = base.copy()
    draw = ImageDraw.Draw(img)
    w, h = img.size
    r = max(w, h) * 0.16
    cx, cy = w - r * 1.1, h - r * 1.1
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(62, 207, 142, 255), outline=(255, 255, 255, 255), width=2)
    return img


# ไอคอน tray ที่กำลังแสดงอยู่ — ให้ notify() ใช้ยิง Windows toast notification
# ได้จากทุกที่ในโปรแกรม (เช่น runner แจ้ง "รันเสร็จแล้ว") None = tray ยังไม่ขึ้น
_tray_icon = None


def notify(title: str, message: str) -> None:
    """ยิง Windows notification ผ่าน tray icon (v0.7.1 — จาก user-journey audit:
    เดิมรันเสร็จ/หยุด/พังแบบ "เงียบสนิท" ผู้ใช้ที่ไปทำงานอื่นอยู่ไม่มีทางรู้เลยว่า
    งานเสร็จหรือยัง ต้องคอยเปิดหน้าเว็บ/ชี้เมาส์ดู tooltip เอาเอง) — เรียกจาก
    thread ไหนก็ได้ ปลอดภัยแบบ no-op ถ้า tray ยังไม่พร้อม"""
    icon = _tray_icon
    if icon is None:
        return
    try:
        icon.notify(message, title)
    except Exception:  # noqa: BLE001 -- notification เป็นของเสริม พังแล้วห้ามล้มงานหลัก
        logger.debug("ยิง notification ไม่สำเร็จ", exc_info=True)


def _open_output_folder(_icon=None, _item=None) -> None:
    """เปิดโฟลเดอร์ผลลัพธ์ใน Explorer จากเมนู tray — ทางลัดที่ใช้บ่อยสุดหลังรันเสร็จ"""
    try:
        from config import load_settings

        out = load_settings().output_dir
        if out and Path(out).is_dir():
            os.startfile(out)  # type: ignore[attr-defined]
        else:
            notify("CropWat Auto-runner", "ยังไม่ได้ตั้งค่าโฟลเดอร์ผลลัพธ์ หรือโฟลเดอร์ยังไม่ถูกสร้าง")
    except Exception:  # noqa: BLE001
        logger.exception("เปิดโฟลเดอร์ผลลัพธ์จาก tray ไม่สำเร็จ")


def _take_screenshot(_icon=None, _item=None) -> None:
    """ถ่ายภาพหน้าจอ CropWat ตอนนี้จากเมนู tray (v0.8.0 — แทนที่ "เข้าดูเดสก์ท็อป
    ซ่อน" เดิมที่ถอดออกแล้วเพราะชน "Cannot make a visible window modal" ซ้ำๆ) —
    ใช้เช็คสถานะเฉยๆ แยกจาก screenshot ที่ต้องส่งงาน เปิดดูด้วยตัวดูภาพเริ่มต้น
    ของเครื่องทันทีที่ถ่ายเสร็จ"""
    import desktop_session

    try:
        import runner
        from config import load_settings

        if not runner.is_run_active():
            notify("CropWat Auto-runner", "ยังไม่ได้กำลังรันอยู่ — ถ่ายภาพได้เฉพาะระหว่างรันในโหมดเดสก์ท็อปซ่อนเท่านั้น")
            return
        path = desktop_session.request_peek_screenshot(load_settings().output_dir)
        os.startfile(str(path))  # type: ignore[attr-defined]
    except desktop_session.DesktopSessionError as exc:
        notify("CropWat Auto-runner", str(exc))
    except Exception:  # noqa: BLE001
        logger.exception("ถ่ายภาพหน้าจอจาก tray ไม่สำเร็จ")
        notify("CropWat Auto-runner", "ถ่ายภาพหน้าจอไม่สำเร็จ")


def _toggle_overlay(_icon=None, _item=None) -> None:
    """สลับซ่อน/แสดงแถบ progress ลอย — จากเมนู tray เท่านั้น (ตั้งใจ: พอ overlay
    ถูกซ่อนแล้วจะไม่มีปุ่มบนตัว overlay เองให้กดเรียกคืนอีก ต้องมาที่ tray)"""
    if overlay_hidden.is_set():
        overlay_hidden.clear()
    else:
        overlay_hidden.set()


def _tray_loop() -> None:
    global _tray_icon
    import pystray

    def _open_window(_icon=None, _item=None) -> None:
        _open_app_window()

    menu = pystray.Menu(
        pystray.MenuItem("เปิดหน้าต่างโปรแกรม", _open_window, default=True),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("▶ เริ่มประมวลผล", lambda _i, _m: _start_run()),
        pystray.MenuItem("⏹ หยุด", lambda _i, _m: _stop_run()),
        pystray.MenuItem("📸 Snapshot", _take_screenshot),
        pystray.MenuItem(
            "แสดงแถบความคืบหน้า",
            _toggle_overlay,
            checked=lambda _item: not overlay_hidden.is_set(),
        ),
        pystray.MenuItem("📂 เปิดโฟลเดอร์ผลลัพธ์", _open_output_folder),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("✕ ปิดโปรแกรม", lambda _i, _m: _quit_app()),
    )

    icon = pystray.Icon(
        "CropWatAutoRunner", _tray_icon_image(running=False), "CropWat Auto-runner", menu
    )
    _tray_icon = icon

    def _watch(icon: "pystray.Icon") -> None:
        """poll run_state ทุก 500ms แล้วสลับไอคอน (จุดเขียว) + tooltip ตามสถานะ —
        pystray รองรับตั้ง icon.icon/icon.title จาก thread อื่นได้ปลอดภัย

        v0.5.16 — บั๊กที่เจอจากผู้ใช้จริง (เข้าใจผิดว่า "tray icon ไม่ทำงาน"):
        เดิมอัปเดต title เฉพาะตอนสถานะ running "เปลี่ยน" (idle↔running) เท่านั้น
        ทำให้ตัวเลข X/Y ค้างอยู่ที่ค่าแรกตลอดการรัน (มักเป็น 0/0 เพราะ candidate_
        total ยังไม่ทันคำนวณตอนเพิ่งเริ่ม) ทั้งที่คลิกเมนูใช้งานได้ปกติทุกอย่าง —
        อัปเดต title ทุกรอบ poll ตอนกำลังรัน (ถูกแค่เปลี่ยนรูปไอคอนเท่านั้นที่ยัง
        เช็คเฉพาะตอน state เปลี่ยนจริง กันวาดรูปใหม่ทุก 500ms โดยไม่จำเป็น)"""
        from state import run_state

        last_running = None
        while True:
            try:
                snap = run_state.snapshot()
                running = snap.overall_state.value in ("running", "stopping")
                if running != last_running:
                    icon.icon = _tray_icon_image(running=running)
                    last_running = running
                if running:
                    # v0.8.0 — ผู้ใช้ชอบดู progress ผ่าน tray มากกว่า overlay
                    # ("ไม่เกะกะเท่า") เพิ่มปี + % ให้เห็นรายละเอียดในคำเดียวโดย
                    # ไม่ต้องเปิดหน้าต่างโปรแกรม (hover เมาส์บนไอคอนก็เห็นแล้ว)
                    total = snap.candidate_total
                    done = snap.candidate_done
                    pct_txt = f" ({done / total:.0%})" if total else ""
                    year_txt = f" ปี {snap.current_year}" if snap.current_year else ""
                    eta_txt = f" · {eta}" if (eta := _format_eta(snap.eta_seconds)) else ""
                    icon.title = (
                        f"CropWat Auto-runner —{year_txt} {done}/{total}{pct_txt}{eta_txt}"
                    )
                else:
                    icon.title = "CropWat Auto-runner — พร้อมใช้งาน"
            except Exception:  # noqa: BLE001 -- poll พังรอบไหนข้ามรอบนั้น อย่าให้ tray ตาย
                logger.exception("tray icon watcher ล้มเหลว")
            time.sleep(0.5)

    threading.Thread(target=_watch, args=(icon,), daemon=True, name="tray-watcher").start()
    icon.run()


def _hotkey_loop() -> None:
    """RegisterHotKey ผูกกับ thread ที่เรียกมัน — ต้องลงทะเบียนและวน GetMessage
    ใน thread เดียวกันนี้เท่านั้น (WM_HOTKEY ถูกส่งเข้า message queue ของ thread)"""
    user32 = ctypes.windll.user32
    if not user32.RegisterHotKey(None, HOTKEY_START_ID, MOD_CONTROL | MOD_ALT, VK_F9):
        logger.warning("ลงทะเบียน hotkey Ctrl+Alt+F9 ไม่สำเร็จ (อาจชนกับโปรแกรมอื่น)")
    if not user32.RegisterHotKey(None, HOTKEY_STOP_ID, MOD_CONTROL | MOD_ALT, VK_F10):
        logger.warning("ลงทะเบียน hotkey Ctrl+Alt+F10 ไม่สำเร็จ (อาจชนกับโปรแกรมอื่น)")

    msg = ctypes.wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
        if msg.message == WM_HOTKEY:
            if msg.wParam == HOTKEY_START_ID:
                logger.info("hotkey: เริ่มรัน")
                _start_run()
            elif msg.wParam == HOTKEY_STOP_ID:
                logger.info("hotkey: หยุด")
                _stop_run()


# --- overlay window -------------------------------------------------------

def _overlay_loop() -> None:
    import tkinter as tk

    from state import run_state

    BG = "#1c2230"
    FG = "#e8ecf5"
    DIM = "#8a93a8"
    BAR_BG = "#323a4d"
    BAR_FG = "#4da3ff"
    BAR_DONE = "#3ecf8e"

    root = tk.Tk()

    # v0.5.16 — บั๊กที่เจอจากผู้ใช้จริง: overlay ขึ้นไม่เต็มข้อความบนจอความ
    # ละเอียดสูง (150%/200% scaling) — launcher.py ประกาศ process เป็น "DPI-aware"
    # แล้ว (กัน Windows bitmap-stretch ทั้งบานแบบเบลอ) แต่นั่นทำให้ Tk รู้ DPI จริง
    # แล้วขยายฟอนต์ตาม ในขณะที่ขนาดหน้าต่าง/padding เดิมเป็น pixel ตายตัว (คิดที่
    # 96 DPI) ข้อความเลยล้นกรอบ — คำนวณ scale จาก DPI จริงของจอ แล้วขยายทุกค่าที่
    # เป็น "pixel ดิบ" (ขนาดหน้าต่าง/padding/ความสูงแถบ) ตาม ส่วนขนาดฟอนต์ (pt)
    # ปล่อยให้ Tk จัดการเองเพราะ DPI-aware แล้วมันคำนวณถูกต้องอยู่แล้ว
    try:
        scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
    except Exception:  # noqa: BLE001
        scale = 1.0

    def px(n: float) -> int:
        return round(n * scale)

    # v0.10.1 — บั๊กจากผู้ใช้จริง: ข้อความสถานะยาวขึ้น ("กำลังดำเนินการ 746/5760
    # วันปลูก (13%) · ปี 1986 · เหลือ 6ชม.3น.") + ปุ่มเพิ่มเป็น 6 ปุ่ม แต่ความ
    # กว้างยังตายตัว 400px เท่าเดิม → ปุ่มโดนดันหลุดขอบ (กด 🗕 พับแถบไม่ได้เลย)
    # — ขยายเป็น 560px + เปลี่ยนลำดับ pack ให้ปุ่มจองพื้นที่ก่อนข้อความ (ดูด้านล่าง)
    W, H = px(560), px(62)

    root.overrideredirect(True)  # ไร้ขอบ/title bar
    root.attributes("-topmost", True)  # ลอยบนสุดเสมอ แม้ CropWat จะ active
    root.attributes("-alpha", 0.94)
    root.configure(bg=BG)
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"{W}x{H}+{sw - W - px(16)}+{sh - H - px(56)}")
    root.withdraw()  # v0.5.13: เริ่มต้นซ่อนไว้ก่อน — โผล่เฉพาะตอนกำลังรัน/เพิ่งจบ
    # (tray icon คือสัญญาณ "ทำงานอยู่เบื้องหลัง" ถาวรตัวเดียว ดู _tray_loop)

    # ลากย้ายได้ทั้งแถบ (ไม่มี title bar ให้จับ)
    drag = {"x": 0, "y": 0}

    def _drag_start(event):
        drag["x"], drag["y"] = event.x, event.y

    def _drag_move(event):
        root.geometry(f"+{event.x_root - drag['x']}+{event.y_root - drag['y']}")

    top_row = tk.Frame(root, bg=BG)
    top_row.pack(fill="x", padx=px(10), pady=(px(7), px(2)))

    status_label = tk.Label(
        top_row, text="พร้อมใช้งาน", bg=BG, fg=FG, font=("Segoe UI", 9, "bold"), anchor="w"
    )

    def _mk_button(parent, text, command, fg=FG):
        btn = tk.Label(
            parent, text=text, bg=BG, fg=fg, font=("Segoe UI", 10), cursor="hand2", padx=px(4)
        )
        btn.bind("<Button-1>", lambda _e: (command(), "break")[1])
        return btn

    # v0.5.13: ใช้ _quit_app() ตัวเดียวกับ tray icon (MessageBoxW ดิบ ไม่ใช่
    # tkinter.messagebox) กันโค้ดถามยืนยัน "ปิดโปรแกรม" ซ้ำซ้อน 2 ที่
    def _hide_overlay_button() -> None:
        overlay_hidden.set()

    # สำคัญ (v0.10.1): pack "ปุ่มก่อนข้อความ" เสมอ — tkinter จัดสรรพื้นที่ตาม
    # ลำดับ pack ถ้าข้อความมาก่อนแล้วยาวเกิน ปุ่มทั้งแถวจะโดนดันหลุดขอบหน้าต่าง
    # (บั๊กที่เจอจริง: กด 🗕 ไม่ได้เพราะปุ่มอยู่นอกจอ) กลับลำดับแล้วข้อความจะโดน
    # ตัดท้ายแทน ปุ่มไม่มีวันหาย
    _mk_button(top_row, "✕", _quit_app, fg=DIM).pack(side="right")
    _mk_button(top_row, "⚙", _open_app_window, fg=DIM).pack(side="right")
    # v0.8.0: ซ่อน overlay เองจากปุ่มบนแถบ — เรียกคืนได้จากเมนู tray เท่านั้น
    # (ตั้งใจไม่มีปุ่ม "แสดง" บน overlay เพราะซ่อนไปแล้วก็ไม่มีอะไรให้กด)
    _mk_button(top_row, "🗕", _hide_overlay_button, fg=DIM).pack(side="right")
    # v0.7.1 (user-journey audit): ปุ่ม Snapshot เช็คสถานะจาก overlay ตรงๆ — จุดที่
    # ผู้ใช้มองระหว่างรันคือ overlay อยู่แล้ว ไม่ต้องอ้อมไปเปิดหน้าเว็บ
    _mk_button(top_row, "📸", _take_screenshot, fg=DIM).pack(side="right")
    _mk_button(top_row, "⏹", _stop_run, fg="#ff7b7b").pack(side="right")
    _mk_button(top_row, "▶", _start_run, fg=BAR_DONE).pack(side="right")
    status_label.pack(side="left", fill="x", expand=True)

    # tkinter ไม่มี event bubbling — bind ที่ root อย่างเดียวจะลากได้เฉพาะตอนจับ
    # พื้นหลังเปล่าๆ ต้อง bind ที่ทุก widget ที่กินพื้นที่ด้วยถึงจะลากได้ทั้งแถบ
    for w in (root, top_row, status_label):
        w.bind("<Button-1>", _drag_start)
        w.bind("<B1-Motion>", _drag_move)

    track = tk.Frame(root, bg=BAR_BG, height=px(6))
    track.pack(fill="x", padx=px(10), pady=(px(2), px(2)))
    track.pack_propagate(False)
    fill = tk.Frame(track, bg=BAR_FG, height=px(6), width=0)
    fill.place(x=0, y=0, relheight=1.0)

    hint_label = tk.Label(
        root, text=HOTKEY_HINT, bg=BG, fg=DIM, font=("Segoe UI", 7), anchor="w"
    )
    hint_label.pack(fill="x", padx=px(10), pady=(0, px(5)))

    hidden = {"v": True}  # เริ่มต้นด้วย withdraw() ไว้แล้วด้านบน
    hide_deadline = {"v": 0.0}  # monotonic time — ค้างโชว์ถึงเวลานี้เป็นอย่างน้อย

    def _refresh():
        try:
            snap = run_state.snapshot()
            total = snap.candidate_total
            done = snap.candidate_done
            running = snap.overall_state.value == "running"
            stopping = snap.overall_state.value == "stopping"
            active = running or stopping

            # v0.5.13: โผล่เฉพาะตอนกำลังรัน/กำลังหยุด/เพิ่งจบ (ค้างไว้
            # _DONE_GRACE_SECONDS วิให้เห็นผลสรุป) — ตอน active ผลักเวลาหมดอายุไป
            # เรื่อยๆ พอเลิก active ปุ๊บ deadline จะเหลืออีก ~15 วิพอดีจากติ๊กล่าสุด
            if active:
                hide_deadline["v"] = time.monotonic() + _DONE_GRACE_SECONDS
            should_show = (
                (not capture_pause.is_set())
                and (not overlay_hidden.is_set())
                and (active or time.monotonic() < hide_deadline["v"])
            )

            if not should_show:
                if not hidden["v"]:
                    root.withdraw()
                    hidden["v"] = True
                root.after(400, _refresh)
                return

            if hidden["v"]:
                root.deiconify()
                root.attributes("-topmost", True)
                hidden["v"] = False

            if total > 0:
                pct = done / total
                track_width = track.winfo_width() or 1
                fill.configure(width=int(track_width * pct))
                year_txt = f" · ปี {snap.current_year}" if snap.current_year else ""
                if running:
                    eta_txt = f" · {eta}" if (eta := _format_eta(snap.eta_seconds)) else ""
                    status_label.configure(
                        text=f"กำลังดำเนินการ {done}/{total} วันปลูก ({pct:.0%}){year_txt}{eta_txt}"
                    )
                    fill.configure(bg=BAR_FG)
                elif stopping:
                    status_label.configure(text=f"กำลังหยุด... {done}/{total}")
                    fill.configure(bg="#ffb020")
                else:
                    status_label.configure(text=f"เสร็จสมบูรณ์ {done}/{total} วันปลูก")
                    fill.configure(bg=BAR_DONE)
            else:
                fill.configure(width=0)
                status_label.configure(text="กำลังดำเนินการ...")
        except Exception:  # noqa: BLE001 -- refresh พังรอบไหนข้ามรอบนั้น อย่าให้ overlay ตาย
            logger.exception("overlay refresh ล้มเหลว")
        root.after(400, _refresh)

    root.after(400, _refresh)
    root.mainloop()


def start_background_ui() -> None:
    """เรียกครั้งเดียวตอน backend startup — ปล่อย 3 daemon thread (hotkey +
    overlay + tray icon) แต่ละตัวครอบ try/except ของตัวเอง พังแล้วแค่เสียฟีเจอร์
    นั้นไป ระบบหลักรันต่อได้ (tray icon เป็นฟีเจอร์ใหม่ v0.5.13 — ถ้า pystray มี
    ปัญหาในบางเครื่อง overlay + hotkey ยังทำงานได้ตามปกติ)"""

    def _safe(target, name):
        def _run():
            try:
                target()
            except Exception:  # noqa: BLE001
                logger.exception("%s ล้มเหลว (ระบบหลักยังทำงานต่อได้)", name)

        threading.Thread(target=_run, daemon=True, name=name).start()

    _safe(_hotkey_loop, "hotkey-listener")
    _safe(_overlay_loop, "progress-overlay")
    _safe(_tray_loop, "tray-icon")
