"""
overlay.py
==========
2 อย่างที่ทำให้ใช้งานได้โดยไม่ต้องเปิดหน้าเว็บเลย (v0.2.0):

1. **แถบ progress ลอย (overlay)** — หน้าต่างเล็กๆ ไร้ขอบ อยู่บนสุดเสมอ (topmost)
   มุมขวาล่างจอ ลากย้ายได้ แสดง progress ระดับ "วันปลูก" แบบ real-time พร้อมปุ่ม
   ▶ เริ่ม / ⏹ หยุด / ⚙ เปิดหน้าตั้งค่า — แก้ปัญหา "เปิด CropWat บังจอแล้วดู
   progress ไม่ได้" เพราะ overlay ลอยทับทุกหน้าต่างรวมถึง CropWat ด้วย

2. **Global hotkeys** — Ctrl+Alt+F9 เริ่มรัน / Ctrl+Alt+F10 หยุด กดได้จากทุกที่
   แม้ CropWat หรือโปรแกรมอื่นกำลัง focus อยู่ (RegisterHotKey ระดับ OS)

ใช้ tkinter (มากับ Python อยู่แล้ว ไม่ต้องลง dependency เพิ่ม) รันใน thread แยก
ของตัวเอง — กติกาสำคัญของ tkinter: ทุกคำสั่ง tk ต้องมาจาก thread เดียวกับที่สร้าง
root เท่านั้น ดังนั้น thread อื่น (hotkey/engine) ห้ามแตะ widget ตรงๆ สื่อสารผ่าน
threading.Event / run_state แล้วให้ refresh loop (root.after) ฝั่ง tk อ่านเอาเอง
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import logging
import threading
import webbrowser

logger = logging.getLogger("overlay")

URL = "http://127.0.0.1:8000"

# ตั้งโดย engine ตอนกำลัง capture screenshot — overlay ต้องซ่อนตัวชั่วคราว ไม่งั้น
# ตัวเองจะติดไปในภาพ (capture ถ่ายจากพิกเซลจริงบนจอ) — ดู capture_screenshots
capture_pause = threading.Event()

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

    def _quit_app():
        """ปิดโปรแกรมทั้งหมด — ตั้งแต่ build แบบ --noconsole (v0.3.0) ไม่มีหน้าต่าง
        console ให้ปิดแล้ว ปุ่มนี้คือทางปิดโปรแกรมทางเดียวที่มองเห็นได้ ถ้ากำลังรัน
        อยู่ถามยืนยันก่อนกันมือลั่น (os._exit เพราะ uvicorn/engine เป็น daemon
        thread ทั้งหมด ปิดตรงๆ ได้ ไม่มี state ที่ต้อง flush)"""
        from tkinter import messagebox

        snap = run_state.snapshot()
        if snap.overall_state.value == "running":
            if not messagebox.askyesno(
                "CropWat Auto-runner",
                "กำลังรันอยู่ — หยุดและปิดโปรแกรมเลยหรือไม่?\n"
                "(ไฟล์ .txt ที่เสร็จแล้วอยู่ครบ เปิดใหม่แล้วรันต่อได้)",
            ):
                return
        import os

        logger.info("ผู้ใช้สั่งปิดโปรแกรมจากปุ่ม ✕ บน overlay")
        os._exit(0)

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

    hidden = {"v": False}

    def _refresh():
        try:
            # ซ่อนตัวตอน engine กำลัง capture screenshot (ดู capture_pause ด้านบน)
            if capture_pause.is_set():
                if not hidden["v"]:
                    root.withdraw()
                    hidden["v"] = True
            else:
                if hidden["v"]:
                    root.deiconify()
                    root.attributes("-topmost", True)
                    hidden["v"] = False

                snap = run_state.snapshot()
                total = snap.candidate_total
                done = snap.candidate_done
                running = snap.overall_state.value == "running"
                stopping = snap.overall_state.value == "stopping"

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
                    status_label.configure(
                        text="กำลังรัน..." if running else "พร้อมเริ่ม (▶ หรือ Ctrl+Alt+F9)"
                    )
        except Exception:  # noqa: BLE001 -- refresh พังรอบไหนข้ามรอบนั้น อย่าให้ overlay ตาย
            logger.exception("overlay refresh ล้มเหลว")
        root.after(400, _refresh)

    root.after(400, _refresh)
    root.mainloop()


def start_background_ui() -> None:
    """เรียกครั้งเดียวตอน backend startup — ปล่อย 2 daemon thread (hotkey + overlay)
    แต่ละตัวครอบ try/except ของตัวเอง พังแล้วแค่เสีย feature นั้นไป ระบบหลักรันต่อได้"""

    def _safe(target, name):
        def _run():
            try:
                target()
            except Exception:  # noqa: BLE001
                logger.exception("%s ล้มเหลว (ระบบหลักยังทำงานต่อได้)", name)

        threading.Thread(target=_run, daemon=True, name=name).start()

    _safe(_hotkey_loop, "hotkey-listener")
    _safe(_overlay_loop, "progress-overlay")
