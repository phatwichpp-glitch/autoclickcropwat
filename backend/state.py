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
            return StateSnapshot(
                overall_state=self._overall_state,
                current_year=self._current_year,
                start_year=self._start_year,
                end_year=self._end_year,
                years=sorted(self._years.values(), key=lambda s: s.year),
            )

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
