"""
state.py
========
เก็บสถานะการรันทั้งหมดไว้ใน memory (ไม่ persist ข้าม process — ถ้า backend restart
กลางคันจะรีเซ็ต ซึ่งรับได้เพราะเป็น local tool ที่ผู้ใช้ดูหน้าจออยู่ตลอด)

Automation engine รันอยู่ใน background thread (pywinauto เป็น blocking call ไม่เหมาะ
กับ asyncio event loop โดยตรง) ส่วน FastAPI/WebSocket รันอยู่บน asyncio event loop หลัก
ดังนั้น state.py ทำหน้าที่เป็นจุดกลางที่ thread ทั้งสองฝั่งคุยกันอย่างปลอดภัย (threading.Lock)
แล้ว "แจ้งเตือน" ฝั่ง asyncio ผ่าน call_soon_threadsafe ทุกครั้งที่มีการเปลี่ยนแปลง
"""

from __future__ import annotations

import asyncio
import threading
import time
from typing import Optional

from models import OverallRunState, StateSnapshot, YearRunStatus, YearStatus


class RunState:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._years: dict[int, YearStatus] = {}
        self._overall_state: OverallRunState = OverallRunState.IDLE
        self._current_year: Optional[int] = None
        self._start_year: Optional[int] = None
        self._end_year: Optional[int] = None
        self._stop_requested = False
        self._candidate_done = 0
        self._candidate_total = 0
        # v0.5.14: จับเวลาจริงตั้งแต่เริ่มรัน ใช้ประมาณ "เหลืออีกกี่นาที" จาก
        # ความเร็วเฉลี่ยจริงที่ทำได้ (ไม่ใช่ตัวเลขคงที่) — แม่นกว่าเดา เพราะ
        # ปรับตามความเร็วเครื่อง/ความเร็ว automation/จำนวน screenshot ที่ต้องถ่าย
        # ในแต่ละวันปลูกจริงๆ
        self._run_started_at: Optional[float] = None
        # v0.5.17 — บั๊กที่เจอจากผู้ใช้จริง (ETA เพี้ยน "1 นาที" แล้วพุ่งเป็น
        # "2 นาที" ทั้งที่ยังไม่เสร็จ): ตอน resume, candidate_done เริ่มจากเลขสูง
        # ทันที (เช่น 206) ที่ elapsed≈0 ทำให้คำนวณ "วินาที/วันปลูก" ต่ำเวอร์
        # เกินจริงตั้งแต่ติ๊กแรก แล้วค่อยๆ ไต่ขึ้นตามจริงพอวันปลูกใหม่ๆ (ที่ช้ากว่า
        # มาก) เริ่มเสร็จ — เก็บ "จำนวนที่เสร็จไปแล้วก่อนรอบนี้จะเริ่ม" (baseline)
        # ไว้แยก แล้วคำนวณความเร็วจากเฉพาะวันปลูกที่ "เสร็จจริงในรอบรันนี้" เท่านั้น
        self._baseline_done = 0

        # websocket subscribers ต้องถูกแตะจาก asyncio loop เท่านั้น
        self._subscribers: set[asyncio.Queue] = set()
        self._loop: Optional[asyncio.AbstractEventLoop] = None

    # --- เชื่อมกับ asyncio loop หลัก (เรียกตอน FastAPI startup) ---
    def bind_loop(self, loop: asyncio.AbstractEventLoop) -> None:
        self._loop = loop

    # --- setup ---
    def init_years(self, start_year: int, end_year: int) -> None:
        with self._lock:
            self._start_year = start_year
            self._end_year = end_year
            self._years = {
                year: YearStatus(year=year, status=YearRunStatus.QUEUED)
                for year in range(start_year, end_year + 1)
            }
        self._notify()

    # --- run lifecycle ---
    def begin_run(self) -> None:
        with self._lock:
            self._overall_state = OverallRunState.RUNNING
            self._stop_requested = False
            self._candidate_done = 0
            self._candidate_total = 0
            self._run_started_at = time.monotonic()
            self._baseline_done = 0
        self._notify()

    def set_baseline_done(self, done: int) -> None:
        """เรียกครั้งเดียวตอนเริ่มรัน (หลัง begin_run) — บอกว่ามีกี่วันปลูกที่
        "เสร็จไปแล้วจากรอบก่อนหน้า" ก่อนรอบนี้จะเริ่มทำอะไรเลย ใช้แยกคำนวณ ETA
        จากความเร็วที่ทำได้จริงในรอบนี้เท่านั้น ดู _estimate_eta_seconds"""
        with self._lock:
            self._baseline_done = done
        self._notify()

    def set_candidate_progress(self, done: int, total: int) -> None:
        """progress ระดับวันปลูก (ละเอียดกว่าระดับปี) — runner อัปเดตหลังจบแต่ละ
        วันปลูกที่ทดลอง ให้ bar เดินสม่ำเสมอ ไม่กระโดดทีละปี"""
        with self._lock:
            self._candidate_done = done
            self._candidate_total = total
        self._notify()

    def end_run(self) -> None:
        with self._lock:
            self._overall_state = OverallRunState.IDLE
            self._current_year = None
        self._notify()

    def request_stop(self) -> None:
        with self._lock:
            self._stop_requested = True
            self._overall_state = OverallRunState.STOPPING
        self._notify()

    def is_stop_requested(self) -> bool:
        with self._lock:
            return self._stop_requested

    def set_current_year(self, year: Optional[int]) -> None:
        with self._lock:
            self._current_year = year
        self._notify()

    # --- per-year updates ---
    def set_year_status(
        self,
        year: int,
        status: YearRunStatus,
        error_message: Optional[str] = None,
        exported_file: Optional[str] = None,
    ) -> None:
        with self._lock:
            existing = self._years.get(year, YearStatus(year=year))
            existing.status = status
            existing.error_message = error_message
            if exported_file is not None:
                existing.exported_file = exported_file
            existing.touch()
            self._years[year] = existing
        self._notify()

    def get_error_years(self) -> list[int]:
        with self._lock:
            return sorted(
                y for y, s in self._years.items() if s.status == YearRunStatus.ERROR
            )

    def get_years_in_range(self, start_year: int, end_year: int) -> list[int]:
        with self._lock:
            for year in range(start_year, end_year + 1):
                if year not in self._years:
                    self._years[year] = YearStatus(year=year)
            return list(range(start_year, end_year + 1))

    # --- snapshot for REST / WebSocket ---
    def snapshot(self) -> StateSnapshot:
        with self._lock:
            eta_seconds = self._estimate_eta_seconds()
            return StateSnapshot(
                overall_state=self._overall_state,
                current_year=self._current_year,
                start_year=self._start_year,
                end_year=self._end_year,
                years=sorted(self._years.values(), key=lambda s: s.year),
                candidate_done=self._candidate_done,
                candidate_total=self._candidate_total,
                eta_seconds=eta_seconds,
            )

    # ต้องมีวันปลูกที่ "เสร็จจริงในรอบรันนี้" อย่างน้อยกี่อันก่อนจะกล้าประมาณ ETA
    # — ตัวอย่างเดียวแกว่งง่ายเกินไป (เช่น วันปลูกแรกบังเอิญเร็ว/ช้าผิดปกติ)
    _MIN_SAMPLES_FOR_ETA = 2

    def _estimate_eta_seconds(self) -> Optional[float]:
        """v0.5.17 — เขียนใหม่หลังเจอบั๊กจากผู้ใช้จริง (ETA บอก "1 นาที" แล้วพุ่ง
        เป็น "2 นาที" ทั้งที่ยังไม่เสร็จ งงว่าคำนวณถอยหลังได้ไง): เดิม (v0.5.14)
        ใช้ candidate_done ทั้งหมดหาร elapsed ทั้งหมด — พอเป็นการ resume ที่มีของ
        เก่าเสร็จไปแล้วเยอะ (เช่น 206 วันปลูก) candidate_done จะกระโดดขึ้นสูง
        ทันทีที่ elapsed≈0 ทำให้ "วินาที/วันปลูก" ต่ำเวอร์เกินจริงตั้งแต่ติ๊กแรก
        (ดูเหมือนเร็วมาก) แล้วค่อยๆ ไต่ขึ้นตามจริงเมื่อวันปลูกใหม่ๆ ที่ช้ากว่ามาก
        เริ่มเสร็จทีละอัน — ตัวเลขเลย "เพิ่มขึ้น" ระหว่างรันซึ่งดูขัดสามัญสำนึก

        แก้ด้วยการนับความเร็วจากเฉพาะวันปลูกที่ "เสร็จจริงในรอบรันนี้" เท่านั้น
        (ไม่รวมของเก่าที่ resume มา ดู _baseline_done/set_baseline_done) และรอให้
        มีตัวอย่างจริงอย่างน้อย _MIN_SAMPLES_FOR_ETA ก่อนกล้าฟันธง กัน sample
        เดียวที่อาจเร็ว/ช้าผิดปกติทำให้เดาพลาดไปไกล — ต้องเรียกในนี้เท่านั้น (ไม่
        ล็อกซ้ำ — caller ถือ lock อยู่แล้ว) คืน None ถ้ายังประมาณไม่ได้"""
        newly_done = self._candidate_done - self._baseline_done
        if (
            self._overall_state != OverallRunState.RUNNING
            or self._run_started_at is None
            or newly_done < self._MIN_SAMPLES_FOR_ETA
            or self._candidate_total <= self._candidate_done
        ):
            return None
        elapsed = time.monotonic() - self._run_started_at
        seconds_per_candidate = elapsed / newly_done
        remaining = self._candidate_total - self._candidate_done
        return seconds_per_candidate * remaining

    # --- pub/sub for WebSocket (asyncio side only) ---
    def subscribe(self) -> asyncio.Queue:
        queue: asyncio.Queue = asyncio.Queue()
        self._subscribers.add(queue)
        return queue

    def unsubscribe(self, queue: asyncio.Queue) -> None:
        self._subscribers.discard(queue)

    def _notify(self) -> None:
        """เรียกได้จากทั้ง worker thread และ asyncio thread"""
        if self._loop is None:
            return
        snapshot = self.snapshot()
        self._loop.call_soon_threadsafe(self._push_to_subscribers, snapshot)

    def _push_to_subscribers(self, snapshot: StateSnapshot) -> None:
        for queue in list(self._subscribers):
            queue.put_nowait(snapshot)


# instance เดียวที่ใช้ร่วมกันทั้งแอป
run_state = RunState()
