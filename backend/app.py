"""
app.py
======
FastAPI entrypoint — REST API + WebSocket ให้ frontend (localhost) เรียกใช้

รันด้วย:
    uvicorn app:app --host 127.0.0.1 --port 8000
(รันจากภายในโฟลเดอร์ backend/ — อย่าใช้ --reload ตอนใช้งานจริง เพราะ background
thread ที่คุม pywinauto จะไม่รอด reload)

หน้าเว็บ (frontend/) ถูก mount ไว้ที่ "/" อยู่แล้ว เปิด http://127.0.0.1:8000 ได้เลย
"""

from __future__ import annotations

import asyncio
import logging
import sys
import time
from pathlib import Path

from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

import os_dialogs
import runner
from config import (
    Settings,
    excel_path,
    load_settings,
    save_settings,
    screenshot_dir,
    word_path,
)
from file_engine import paths as file_paths
from models import RunRequest, ScanResult, StateSnapshot, StationScan, YearRunStatus
from state import run_state

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("app")


def _frontend_dir() -> Path:
    """ตอนถูก build เป็น .exe ด้วย PyInstaller (--onefile) ไฟล์ frontend/ ที่ bundle
    ไปด้วยจะถูกแตกไปไว้ที่ sys._MEIPASS (temp dir ชั่วคราวต่อการรันแต่ละครั้ง) —
    ต่างจากตอน dev ที่ frontend/ อยู่เป็นโฟลเดอร์พี่น้องของ backend/ ตามปกติ"""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)) / "frontend"
    return Path(__file__).parent.parent / "frontend"


FRONTEND_DIR = _frontend_dir()

app = FastAPI(title="CropWat Auto-runner")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# v0.5.22 — บั๊กที่เจอจากผู้ใช้จริง: หลังอัปเดตหลายรอบในเครื่องเดียวกัน (browser
# profile เดียวกันสะสม cache มานาน) หน้าเว็บค้างที่ "กำลังเชื่อมต่อ..." ตลอด ทั้ง
# ที่ backend เองทำงานปกติสมบูรณ์ (ยืนยันจาก log: ไม่มี request ขอ app.js/api/ws
# เลยแม้แต่ครั้งเดียว — เบราว์เซอร์ใช้ของที่ cache ไว้เก่าโดยไม่ยิง request ใหม่)
# เครื่องเพื่อนที่เพิ่งลงครั้งแรก (ไม่มี cache เก่า) กลับใช้งานได้ปกติ — ยืนยันว่า
# เป็นเรื่อง cache ไม่ใช่บั๊ก logic จริง — StaticFiles ของ Starlette ไม่ตั้ง
# Cache-Control มาให้เอง อาศัย heuristic caching ของเบราว์เซอร์เอง ซึ่งอันตราย
# มากสำหรับโปรแกรมที่ self-update บ่อย (เนื้อไฟล์เปลี่ยนแต่ URL เดิมทุกครั้ง) —
# บังคับไม่ให้ cache ไฟล์เว็บเลยทุกก้อน กันปัญหานี้เกิดซ้ำกับผู้ใช้คนอื่น
@app.middleware("http")
async def _no_cache_headers(request, call_next):
    response = await call_next(request)
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.on_event("startup")
async def on_startup() -> None:
    run_state.bind_loop(asyncio.get_running_loop())
    settings = load_settings()
    run_state.init_years(settings.default_start_year, settings.default_end_year)

    # แถบ progress ลอย + global hotkeys (Ctrl+Alt+F9/F10) — daemon threads ที่
    # พังได้โดยไม่กระทบระบบหลัก (เช่น เครื่องที่ tkinter มีปัญหา ก็แค่ไม่มี overlay)
    import overlay

    overlay.start_background_ui()


# ---------------------------------------------------------------------------
# Config (หน้าตั้งค่า path ไฟล์ต่างๆ — ขั้นที่ 5 จะทำหน้า UI มาเรียก endpoint นี้)
# ---------------------------------------------------------------------------

@app.get("/api/config", response_model=Settings)
async def get_config() -> Settings:
    return load_settings()


@app.put("/api/config", response_model=Settings)
async def update_config(settings: Settings) -> Settings:
    save_settings(settings)
    return settings


# ---------------------------------------------------------------------------
# Folder picker + เปิดโฟลเดอร์ใน Explorer (v1.0 UX pass — ลดการพิมพ์ path เอง
# ซึ่งเป็นจุดพังบ่อยสำหรับผู้ใช้ที่ไม่ถนัดคอมพิวเตอร์)
# ---------------------------------------------------------------------------

class BrowseFolderRequest(BaseModel):
    initial_dir: str = ""


@app.post("/api/browse-folder")
async def browse_folder(req: BrowseFolderRequest) -> dict:
    path = await asyncio.to_thread(os_dialogs.pick_folder, req.initial_dir)
    return {"path": path}


class OpenFolderRequest(BaseModel):
    path: str


@app.post("/api/open-folder")
async def open_folder(req: OpenFolderRequest) -> dict:
    try:
        await asyncio.to_thread(os_dialogs.open_in_explorer, req.path)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    return {"ok": True}


@app.post("/api/browse-file")
async def browse_file(req: BrowseFolderRequest) -> dict:
    path = await asyncio.to_thread(os_dialogs.pick_file, req.initial_dir)
    return {"path": path}


@app.post("/api/launch-cropwat")
async def launch_cropwat() -> dict:
    """เปิดโปรแกรม CropWat 8.0 ให้จาก path ที่ตั้งไว้ในหน้าตั้งค่า (cropwat_exe_path)
    — ไม่ auto-detect เอง (ผู้ใช้ต้องเลือก path ก่อนครั้งแรกผ่านปุ่ม Browse)"""
    settings = load_settings()
    if not settings.cropwat_exe_path:
        raise HTTPException(400, "ยังไม่ได้ตั้งค่า path โปรแกรม CropWat 8.0 — ไปที่แท็บตั้งค่าก่อน")
    try:
        await asyncio.to_thread(os_dialogs.launch_exe, settings.cropwat_exe_path)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc))
    return {"ok": True}


# ---------------------------------------------------------------------------
# Scan โฟลเดอร์ต้นทาง (หน้าตั้งค่า) — หา climate/rain station, crop, soil ให้เอง
# และรายงานว่าปีไหนมีไฟล์ครบ/ขาดบ้าง ก่อนเริ่มรันจริง
# ---------------------------------------------------------------------------

def _scan_station(root: Path, prefix: str, index_fn) -> tuple[list[str], StationScan | None, list[str]]:
    folders = file_paths.find_station_folders(root, prefix)
    names = [str(p) for p in folders]
    if len(folders) != 1:
        return names, None, []
    index = index_fn(folders[0])
    years = index.available_years()
    if not years:
        return names, StationScan(folder=str(folders[0]), years=[], missing_years=[]), []
    missing = [y for y in range(years[0], years[-1] + 1) if y not in years]
    return names, StationScan(folder=str(folders[0]), years=years, missing_years=missing), []


@app.get("/api/scan", response_model=ScanResult)
async def scan_input_dir() -> ScanResult:
    settings = load_settings()
    if not settings.input_dir:
        return ScanResult(errors=["ยังไม่ได้ตั้งค่าโฟลเดอร์ต้นทาง"])

    root = Path(settings.input_dir)
    if not root.is_dir():
        return ScanResult(errors=[f"ไม่พบโฟลเดอร์: {root}"])

    climate_folders, climate_scan, _ = _scan_station(root, "Clim_", file_paths.index_climate_station)
    rain_folders, rain_scan, _ = _scan_station(root, "Rain_", file_paths.index_rain_station)
    crop_files = [str(p) for p in file_paths.find_files_by_extension(root, ".cro")]
    soil_files = [str(p) for p in file_paths.find_files_by_extension(root, ".soi")]

    return ScanResult(
        climate_station_folders=climate_folders,
        rain_station_folders=rain_folders,
        crop_files=crop_files,
        soil_files=soil_files,
        climate=climate_scan,
        rain=rain_scan,
    )


# ---------------------------------------------------------------------------
# Status (Dashboard หลัก)
# ---------------------------------------------------------------------------

@app.get("/api/status", response_model=StateSnapshot)
async def get_status() -> StateSnapshot:
    return run_state.snapshot()


@app.get("/api/output-progress")
async def output_progress() -> dict:
    """ความคืบหน้าไฟล์ output แบบละเอียด (ทำถึงไหนแล้ว/เหลืออะไร) — สแกนไฟล์ .txt
    จริงเทียบกับแผนปัจจุบัน กดเริ่มรันอีกครั้งระบบจะทำต่อจากจุดที่ค้างให้เอง"""
    settings = load_settings()
    return await asyncio.to_thread(runner.scan_output_progress, settings)


@app.get("/api/shift-year-check")
async def shift_year_check() -> dict:
    """v0.5.24 — ตามคำขอผู้ใช้: ต้องสแกนตรวจว่าทุกวันปลูกที่วางแผนไว้มีไฟล์
    climate/rain "ก่อนหน้า" ให้ใช้จริงไหม (ห้ามเดาใช้ไฟล์เดือนอนาคตแทน) ก่อนเริ่ม
    รันทุกครั้ง — frontend เรียกจุดนี้ก่อนกด "เริ่มรันทั้งหมด" เสมอ ถ้าเจอปัญหา
    ต้องเตือน + ขอความยินยอมจากผู้ใช้ก่อน ไม่ปล่อยรันเงียบๆ"""
    settings = load_settings()
    problems = await asyncio.to_thread(runner.check_shift_year_coverage, settings)
    return {"ok": len(problems) == 0, "problems": problems}


# ---------------------------------------------------------------------------
# Run control
# ---------------------------------------------------------------------------

@app.post("/api/run/start", response_model=StateSnapshot)
async def start_run(req: RunRequest) -> StateSnapshot:
    if req.start_year > req.end_year:
        raise HTTPException(400, "start_year ต้องไม่มากกว่า end_year")
    if runner.is_run_active():
        raise HTTPException(409, "มีการรันอยู่แล้ว กรุณารอให้เสร็จหรือกดหยุดก่อน")

    settings = load_settings()
    run_state.init_years(req.start_year, req.end_year)
    years = list(range(req.start_year, req.end_year + 1))
    runner.start_run(years, settings)
    return run_state.snapshot()


@app.post("/api/run/retry", response_model=StateSnapshot)
async def retry_errors() -> StateSnapshot:
    if runner.is_run_active():
        raise HTTPException(409, "มีการรันอยู่แล้ว กรุณารอให้เสร็จหรือกดหยุดก่อน")

    # v0.5.15 — บั๊ก UX ที่เจอจากผู้ใช้จริง: ปุ่มนี้เดิมชื่อ "รันปีที่ค้างใหม่" แต่
    # ทำงานเฉพาะปีสถานะ "มีปัญหา" (error) เท่านั้น — ปีที่แค่ "ยังไม่เคยรันถึง"
    # (เช่น รันค้างไว้กลางทางแล้วหยุด) ไม่ใช่ error จึงไม่ถูกเลือก ทำให้กดแล้วขึ้น
    # "ไม่มีปีที่มีปัญหา" ทั้งที่เห็นชัดว่ายังมีปีค้างอีกเพียบ — ข้อความนี้ต้องบอก
    # ทางออกที่ถูกต้องชัดเจน (กด "เริ่มรันทั้งหมด" แทน ระบบจะข้ามที่เสร็จแล้วเอง)
    error_years = run_state.get_error_years()
    if not error_years:
        raise HTTPException(
            400,
            "ไม่มีปีที่สถานะ 'มีปัญหา' ให้รันใหม่ (ปุ่มนี้ใช้เฉพาะปีที่ error เท่านั้น) "
            "— ถ้าอยากทำต่อจากปีที่ยังไม่เสร็จ ให้กด \"เริ่มรันทั้งหมด\" แทน "
            "ระบบจะข้ามวันปลูกที่เสร็จแล้วให้เองอัตโนมัติ",
        )

    settings = load_settings()
    runner.start_run(error_years, settings)
    return run_state.snapshot()


@app.post("/api/run/year/{year}", response_model=StateSnapshot)
async def run_one_year(year: int) -> StateSnapshot:
    if runner.is_run_active():
        raise HTTPException(409, "มีการรันอยู่แล้ว กรุณารอให้เสร็จหรือกดหยุดก่อน")

    settings = load_settings()
    runner.start_run([year], settings)
    return run_state.snapshot()


@app.post("/api/run/stop", response_model=StateSnapshot)
async def stop_run() -> StateSnapshot:
    run_state.request_stop()
    return run_state.snapshot()


@app.post("/api/run/force-close-cropwat")
async def force_close_cropwat() -> dict:
    """ทางออกฉุกเฉิน (v1.0): "หยุด" ปกติเป็นการขอความร่วมมือ (thread เช็ค flag
    ระหว่าง candidate/ปี) ใช้ไม่ได้ตอน CropWat ค้างสนิทจริงๆ (กด X เองก็ไม่ติด) —
    endpoint นี้บังคับปิด process CropWat โดยตรง + คืนสถานะฝั่งเราเป็น IDLE ทันที
    ไม่ต้องปิด-เปิดโปรแกรม CropWatAutoRunner ทั้งตัวเพื่อรีเซ็ต (ดู runner.force_reset)"""
    killed = await asyncio.to_thread(runner.force_reset)
    return {"ok": True, "killed": killed}


# ---------------------------------------------------------------------------
# เดสก์ท็อปซ่อน (v0.8.0): ถ่ายภาพหน้าจอตอนนี้ — แทนที่การสลับจอเข้าไปดูสด (v0.7.x)
# ซึ่งชน "Cannot make a visible window modal" ซ้ำๆ จนต้องถอดออกทั้งหมด ใช้
# PrintWindow ถ่ายภาพนิ่งแทน ปลอดภัยกว่ามาก (ไม่แตะ SwitchDesktop เลย) เก็บภาพ
# แยกโฟลเดอร์จาก screenshot ที่ต้องส่งงานเสมอ (ดู desktop_session.PEEK_SCREENSHOT_DIR_NAME)
# ---------------------------------------------------------------------------

@app.post("/api/desktop/screenshot")
async def desktop_screenshot() -> dict:
    """ขอถ่ายภาพหน้าจอ CropWat ปัจจุบัน (สำหรับเช็คเฉยๆ ไม่เกี่ยวกับงานที่ต้องส่ง) —
    ใช้ได้เฉพาะระหว่างรันในโหมดเดสก์ท็อปซ่อน"""
    import desktop_session

    if not runner.is_run_active():
        raise HTTPException(
            409, "ยังไม่ได้กำลังรันอยู่ — ถ่ายภาพได้เฉพาะระหว่างรันในโหมดเดสก์ท็อปซ่อนเท่านั้น"
        )
    settings = load_settings()
    try:
        await asyncio.to_thread(desktop_session.request_peek_screenshot, settings.output_dir)
    except desktop_session.DesktopSessionError as exc:
        raise HTTPException(404, str(exc)) from exc
    return {"ok": True, "url": f"/api/desktop/screenshot-image?t={int(time.time())}"}


@app.get("/api/desktop/screenshot-image")
async def desktop_screenshot_image() -> FileResponse:
    import desktop_session

    settings = load_settings()
    path = Path(settings.output_dir) / desktop_session.PEEK_SCREENSHOT_DIR_NAME / "latest.png"
    if not path.exists():
        raise HTTPException(404, "ยังไม่มีภาพหน้าจอที่ถ่ายไว้")
    return FileResponse(path, media_type="image/png")


def _find_latest_screenshot(settings: Settings):
    """หาไฟล์ภาพหน้าจอ "ล่าสุด" ข้ามทั้งสองแหล่ง: ภาพงานส่ง (screenshots/) และ
    ภาพตรวจสอบสด (_peek_screenshots/latest.png) — คืน (path, kind) หรือ (None, None)
    ใช้ mtime ตัดสิน เพราะ engine เขียนไฟล์เสร็จแล้วเท่านั้นถึงจะโผล่ในโฟลเดอร์
    (v0.9.0 — หน้าจอแสดงภาพอัตโนมัติ: โชว์ภาพใหม่ทุกครั้งที่ระบบ capture ตามปฏิทิน
    หรือผู้ใช้กดถ่ายสด โดย frontend คอยเทียบ mtime ผ่าน /api/screenshots/latest)"""
    import desktop_session

    candidates: list[tuple[float, Path, str]] = []
    if settings.output_dir:
        work_dir = screenshot_dir(settings)
        if work_dir.is_dir():
            for p in work_dir.glob("*.png"):
                try:
                    candidates.append((p.stat().st_mtime, p, "work"))
                except OSError:
                    continue
        peek = Path(settings.output_dir) / desktop_session.PEEK_SCREENSHOT_DIR_NAME / "latest.png"
        if peek.is_file():
            try:
                candidates.append((peek.stat().st_mtime, peek, "live"))
            except OSError:
                pass
    if not candidates:
        return None, None
    _mtime, path, kind = max(candidates, key=lambda c: c[0])
    return path, kind


@app.get("/api/screenshots/latest")
async def screenshots_latest() -> dict:
    """metadata ของภาพหน้าจอล่าสุด (งานส่งหรือตรวจสอบสด อันที่ใหม่กว่า) — frontend
    เรียกซ้ำระหว่างรันเพื่อรีเฟรชหน้าจอแสดงภาพอัตโนมัติ เทียบ mtime ก่อนโหลดภาพจริง"""
    settings = load_settings()
    path, kind = await asyncio.to_thread(_find_latest_screenshot, settings)
    if path is None:
        return {"exists": False}
    mtime = path.stat().st_mtime
    return {
        "exists": True,
        "name": path.name,
        "kind": kind,
        "mtime": mtime,
        "url": f"/api/screenshots/latest-image?t={int(mtime * 1000)}",
    }


@app.get("/api/screenshots/latest-image")
async def screenshots_latest_image() -> FileResponse:
    settings = load_settings()
    path, _kind = await asyncio.to_thread(_find_latest_screenshot, settings)
    if path is None:
        raise HTTPException(404, "ยังไม่มีภาพหน้าจอ")
    return FileResponse(path, media_type="image/png")


@app.get("/api/error-shots/{name}")
async def error_shot(name: str) -> FileResponse:
    """เสิร์ฟภาพ error dialog ที่ระบบถ่ายเก็บไว้อัตโนมัติ (v0.11.0) — ให้หน้าเว็บ
    เปิดดูว่า CropWat แจ้ง error อะไรตอนวันปลูกนั้นพัง (กัน path traversal ด้วย)"""
    if "/" in name or "\\" in name or ".." in name:
        raise HTTPException(400, "ชื่อไฟล์ไม่ถูกต้อง")
    settings = load_settings()
    path = Path(settings.output_dir) / "_error_dialogs" / name
    if not path.is_file():
        raise HTTPException(404, "ไม่พบภาพ error")
    return FileResponse(path, media_type="image/png")


@app.get("/api/cropwat-exe")
async def cropwat_exe_status() -> dict:
    """สถานะการค้นหาโปรแกรม CropWat 8.0 (v0.10.2) — ให้หน้าตั้งค่าแสดงว่า "ระบบ
    หาเจอเองแล้วที่ไหน" ตอบข้อสงสัยยอดฮิต "ต้องตั้ง path เองไหม" ได้ทันทีโดยไม่
    ต้องเปิดคู่มือ: ติดตั้งที่ตำแหน่งมาตรฐาน = ไม่ต้องตั้งอะไรเลย"""
    import desktop_session

    settings = load_settings()
    path = await asyncio.to_thread(
        desktop_session.resolve_cropwat_exe, settings.cropwat_exe_path
    )
    if path is None:
        return {"found": False, "path": None, "source": None}
    configured_valid = bool(settings.cropwat_exe_path) and Path(settings.cropwat_exe_path).is_file()
    return {
        "found": True,
        "path": str(path),
        "source": "configured" if configured_valid else "auto",
    }


@app.post("/api/window/minimize")
async def window_minimize() -> dict:
    """ย่อหน้าต่างโปรแกรมลงถาดระบบ (v0.8.0) — เปิดกลับได้จากเมนู tray icon
    ("เปิดหน้าต่างโปรแกรม")"""
    import launcher

    await asyncio.to_thread(launcher._minimize_app_window)
    return {"ok": True}


@app.post("/api/window/show")
async def window_show() -> dict:
    """โชว์หน้าต่างโปรแกรมของ process นี้ (v0.10.0) — เรียกจาก instance ที่สอง
    ตอนผู้ใช้ดับเบิลคลิก .exe ซ้ำ (หน้าต่าง standalone เป็นของ process นี้
    instance อื่นสร้าง/ดึงแทนไม่ได้ โดยเฉพาะตอนถูกซ่อนลงถาดระบบอยู่)"""
    import launcher

    await asyncio.to_thread(launcher._launch_app_window)
    return {"ok": True}


# ---------------------------------------------------------------------------
# เฟส 2: สร้าง/อัปเดต Excel — อ่าน .txt ทั้งหมดที่มีอยู่จริง เขียนทับ sheet Result
# ใหม่ทั้งแผ่น เรียกซ้ำได้ทุกเมื่อ ไม่ต้องรอเฟส 1 (รัน CropWat) เสร็จก่อน
# ---------------------------------------------------------------------------

@app.post("/api/build-excel")
async def build_excel() -> dict:
    settings = load_settings()
    if not settings.output_dir:
        raise HTTPException(400, "ยังไม่ได้ตั้งค่าโฟลเดอร์ผลลัพธ์")
    try:
        years_written = await asyncio.to_thread(runner.build_excel, settings)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    except PermissionError as exc:
        # เคสที่เจอบ่อยแน่นอน: ผู้ใช้เปิด Result.xlsx ค้างอยู่ใน Excel ตอนกดปุ่ม —
        # Windows lock ไฟล์ไว้ทำให้เขียนทับไม่ได้ ต้องบอกวิธีแก้ตรงๆ ไม่ใช่ 500
        raise HTTPException(
            409,
            f"เขียนไฟล์ {excel_path(settings).name} ไม่ได้ — ถ้าเปิดไฟล์นี้ค้างอยู่ใน "
            "Excel ให้ปิดก่อนแล้วกดใหม่อีกครั้ง",
        ) from exc
    return {"years_written": years_written, "path": str(excel_path(settings))}


@app.get("/api/download")
async def download_excel_master() -> FileResponse:
    settings = load_settings()
    path = excel_path(settings)
    if not settings.output_dir or not path.exists():
        raise HTTPException(404, "ยังไม่พบไฟล์ Result.xlsx (ลองกด 'สร้าง/อัปเดต Excel' ก่อน)")
    return FileResponse(path, filename=path.name)


# ---------------------------------------------------------------------------
# ไฟล์ Word รวมภาพ screenshot (โครงสร้างเหมือนไฟล์ .docx ตัวอย่างจริงของผู้ใช้:
# บรรทัดวันที่ + ภาพตาราง + ภาพกราฟ ต่อ 1 วันปลูก) — สร้างซ้ำได้ทุกเมื่อเหมือน Excel
# ---------------------------------------------------------------------------

@app.post("/api/build-word")
async def build_word() -> dict:
    settings = load_settings()
    if not settings.output_dir:
        raise HTTPException(400, "ยังไม่ได้ตั้งค่าโฟลเดอร์ผลลัพธ์")
    try:
        candidates_written = await asyncio.to_thread(runner.build_word, settings)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    except PermissionError as exc:
        raise HTTPException(
            409,
            f"เขียนไฟล์ {word_path(settings).name} ไม่ได้ — ถ้าเปิดไฟล์นี้ค้างอยู่ใน "
            "Word ให้ปิดก่อนแล้วกดใหม่อีกครั้ง",
        ) from exc
    return {"candidates_written": candidates_written, "path": str(word_path(settings))}


@app.get("/api/download-word")
async def download_word_doc() -> FileResponse:
    settings = load_settings()
    path = word_path(settings)
    if not settings.output_dir or not path.exists():
        raise HTTPException(404, "ยังไม่พบไฟล์ Screenshots.docx (ลองกด 'สร้างไฟล์ Word' ก่อน)")
    return FileResponse(path, filename=path.name)


# ---------------------------------------------------------------------------
# เช็คอัปเดต + อัปเดตตัวเองจาก GitHub Releases (ดูรายละเอียดใน updater.py)
# ---------------------------------------------------------------------------

@app.get("/api/update/check")
async def update_check() -> dict:
    import updater

    return await asyncio.to_thread(updater.check_for_update)


@app.post("/api/update/apply")
async def update_apply() -> dict:
    import updater

    if runner.is_run_active():
        raise HTTPException(409, "กำลังรันอยู่ — กดหยุดหรือรอให้เสร็จก่อนค่อยอัปเดต")
    try:
        await asyncio.to_thread(updater.apply_update)
    except updater.UpdateError as exc:
        raise HTTPException(400, str(exc)) from exc
    return {"ok": True, "message": "กำลังอัปเดต — โปรแกรมจะปิดและเปิดใหม่เองอัตโนมัติ"}


# ---------------------------------------------------------------------------
# WebSocket: push สถานะแบบ real-time ทุกครั้งที่มีการเปลี่ยนแปลง
# ---------------------------------------------------------------------------

@app.websocket("/ws")
async def ws_status(websocket: WebSocket) -> None:
    await websocket.accept()
    queue = run_state.subscribe()
    try:
        await websocket.send_json(run_state.snapshot().model_dump())
        while True:
            snapshot = await queue.get()
            await websocket.send_json(snapshot.model_dump())
    except WebSocketDisconnect:
        pass
    finally:
        run_state.unsubscribe(queue)


# frontend เว็บหน้าเดียว — mount ทีหลังสุดเพื่อไม่ให้บัง /api และ /ws
app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")
