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


def _assets_dir() -> Path:
    """เหมือน _frontend_dir() ใน app.py — ตอน build เป็น .exe ไฟล์ assets/ ที่
    bundle ไปด้วย (--add-data "assets;assets") จะถูกแตกไปไว้ที่ sys._MEIPASS"""
    import sys

    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)) / "assets"
    return Path(__file__).parent / "assets"

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


def _tray_loop() -> None:
    import pystray

    def _open_window(_icon=None, _item=None) -> None:
        webbrowser.open(URL)

    menu = pystray.Menu(
        pystray.MenuItem("เปิดหน้าต่างโปรแกรม", _open_window, default=True),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("▶ เริ่มรันทั้งหมด", lambda _i, _m: _start_run()),
        pystray.MenuItem("⏹ หยุด", lambda _i, _m: _stop_run()),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("✕ ปิดโปรแกรม", lambda _i, _m: _quit_app()),
    )

    icon = pystray.Icon(
        "CropWatAutoRunner", _tray_icon_image(running=False), "CropWat Auto-runner", menu
    )

    def _watch(icon: "pystray.Icon") -> None:
        """poll run_state ทุก 500ms แล้วสลับไอคอน (จุดเขียว) + tooltip ตามสถานะ —
        pystray รองรับตั้ง icon.icon/icon.title จาก thread อื่นได้ปลอดภัย"""
        from state import run_state

        last_running = None
        while True:
            try:
                snap = run_state.snapshot()
                running = snap.overall_state.value in ("running", "stopping")
                if running != last_running:
                    icon.icon = _tray_icon_image(running=running)
                    icon.title = (
                        f"CropWat Auto-runner — กำลังรัน {snap.candidate_done}/{snap.candidate_total}"
                        if running
                        else "CropWat Auto-runner — พร้อมใช้งาน"
                    )
                    last_running = running
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
    W, H = 400, 62

    root = tk.Tk()
    root.overrideredirect(True)  # ไร้ขอบ/title bar
    root.attributes("-topmost", True)  # ลอยบนสุดเสมอ แม้ CropWat จะ active
    root.attributes("-alpha", 0.94)
    root.configure(bg=BG)
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"{W}x{H}+{sw - W - 16}+{sh - H - 56}")
    root.withdraw()  # v0.5.13: เริ่มต้นซ่อนไว้ก่อน — โผล่เฉพาะตอนกำลังรัน/เพิ่งจบ
    # (tray icon คือสัญญาณ "ทำงานอยู่เบื้องหลัง" ถาวรตัวเดียว ดู _tray_loop)

    # ลากย้ายได้ทั้งแถบ (ไม่มี title bar ให้จับ)
    drag = {"x": 0, "y": 0}

    def _drag_start(event):
        drag["x"], drag["y"] = event.x, event.y

    def _drag_move(event):
        root.geometry(f"+{event.x_root - drag['x']}+{event.y_root - drag['y']}")

    top_row = tk.Frame(root, bg=BG)
    top_row.pack(fill="x", padx=10, pady=(7, 2))

    status_label = tk.Label(
        top_row, text="พร้อมเริ่ม", bg=BG, fg=FG, font=("Segoe UI", 9, "bold"), anchor="w"
    )
    status_label.pack(side="left", fill="x", expand=True)

    # tkinter ไม่มี event bubbling — bind ที่ root อย่างเดียวจะลากได้เฉพาะตอนจับ
    # พื้นหลังเปล่าๆ ต้อง bind ที่ทุก widget ที่กินพื้นที่ด้วยถึงจะลากได้ทั้งแถบ
    for w in (root, top_row, status_label):
        w.bind("<Button-1>", _drag_start)
        w.bind("<B1-Motion>", _drag_move)

    def _mk_button(parent, text, command, fg=FG):
        btn = tk.Label(
            parent, text=text, bg=BG, fg=fg, font=("Segoe UI", 10), cursor="hand2", padx=4
        )
        btn.bind("<Button-1>", lambda _e: (command(), "break")[1])
        return btn

    # v0.5.13: ใช้ _quit_app() ตัวเดียวกับ tray icon (MessageBoxW ดิบ ไม่ใช่
    # tkinter.messagebox) กันโค้ดถามยืนยัน "ปิดโปรแกรม" ซ้ำซ้อน 2 ที่
    _mk_button(top_row, "✕", _quit_app, fg=DIM).pack(side="right")
    _mk_button(top_row, "⚙", lambda: webbrowser.open(URL), fg=DIM).pack(side="right")
    _mk_button(top_row, "⏹", _stop_run, fg="#ff7b7b").pack(side="right")
    _mk_button(top_row, "▶", _start_run, fg=BAR_DONE).pack(side="right")

    track = tk.Frame(root, bg=BAR_BG, height=6)
    track.pack(fill="x", padx=10, pady=(2, 2))
    track.pack_propagate(False)
    fill = tk.Frame(track, bg=BAR_FG, height=6, width=0)
    fill.place(x=0, y=0, relheight=1.0)

    hint_label = tk.Label(
        root, text=HOTKEY_HINT, bg=BG, fg=DIM, font=("Segoe UI", 7), anchor="w"
    )
    hint_label.pack(fill="x", padx=10, pady=(0, 5))

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
            should_show = (not capture_pause.is_set()) and (
                active or time.monotonic() < hide_deadline["v"]
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
                    status_label.configure(
                        text=f"กำลังรัน {done}/{total} วันปลูก ({pct:.0%}){year_txt}"
                    )
                    fill.configure(bg=BAR_FG)
                elif stopping:
                    status_label.configure(text=f"กำลังหยุด... {done}/{total}")
                    fill.configure(bg="#ffb020")
                else:
                    status_label.configure(text=f"จบแล้ว {done}/{total} วันปลูก")
                    fill.configure(bg=BAR_DONE)
            else:
                fill.configure(width=0)
                status_label.configure(text="กำลังรัน...")
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
