"""Bounded, single-flight background refreshes; no Streamlit calls in workers."""
from __future__ import annotations

import logging
import threading
import time
from concurrent.futures import Future, ThreadPoolExecutor
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable


@dataclass
class _Job:
    future: Future | None = None
    started: float = 0.0
    next_refresh: float = 0.0
    value: Any = None
    fetched: float | None = None
    error: str = ""
    timed_out: bool = False
    revision: int = 0
    touched: float = 0.0


@dataclass(frozen=True)
class RefreshState:
    value: Any
    pending: bool
    error: str
    revision: int
    age_seconds: float | None


@dataclass(frozen=True)
class _Outcome:
    value: Any
    error: str
    completed: float


class BackgroundRefresh:
    def __init__(self, *, ttl=60.0, stale_ttl=300.0, timeout=25.0,
                 retry_interval=60.0, max_entries=64, max_workers=4, clock=time.monotonic):
        self.ttl = ttl
        self.stale_ttl = stale_ttl
        self.timeout = timeout
        self.retry_interval = retry_interval
        self.max_entries = max_entries
        self.clock = clock
        self._jobs: dict[Any, _Job] = {}
        self._lock = threading.RLock()
        self._pool = ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="ubike-refresh")

    def _run(self, loader):
        try:
            return _Outcome(loader(), "", self.clock())
        except Exception as exc:
            return _Outcome(None, str(exc) or type(exc).__name__, self.clock())

    def poll(self, key, loader: Callable[[], Any], *, force=False) -> RefreshState:
        """Return immediately, including while a catalog/API request is stalled.

        A timed-out running request remains single-flight until it actually exits.
        Its late result is discarded; repeated polls never create duplicate workers.
        Idle entries may be evicted, but running futures never are.
        """
        with self._lock:
            now = self.clock()
            job = self._jobs.get(key)
            if job is None:
                if len(self._jobs) >= self.max_entries:
                    idle = [k for k, j in self._jobs.items() if j.future is None or j.future.done()]
                    if not idle:
                        return RefreshState(None, False, "資料服務忙碌，請稍後再試。", 0, None)
                    del self._jobs[min(idle, key=lambda k: self._jobs[k].touched)]
                job = self._jobs[key] = _Job()
            job.touched = now

            if job.future is not None:
                outcome = job.future.result() if job.future.done() and not job.future.cancelled() else None
                # Deadline wins over a late completion that was not yet observed.
                elapsed = (outcome.completed if outcome is not None else now) - job.started
                if not job.timed_out and elapsed >= self.timeout:
                    job.timed_out = True
                    job.error = f"即時資料查詢超過 {self.timeout:g} 秒，保留原有資料；稍後可重試。"
                    job.next_refresh = now + self.retry_interval
                    job.revision += 1
                    job.future.cancel()
                    logging.getLogger(__name__).warning("Background refresh timed out: %s", key)
                if job.future.done():
                    if not job.timed_out:
                        if outcome is not None and not outcome.error:
                            job.value = outcome.value
                            job.fetched = outcome.completed
                            job.error = ""
                            job.next_refresh = outcome.completed + self.ttl
                        else:
                            job.error = outcome.error if outcome is not None else "資料查詢已取消。"
                            logging.getLogger(__name__).warning("Background refresh failed: %s", job.error)
                            job.next_refresh = now + self.retry_interval
                        job.revision += 1
                    job.future = None

            if job.future is None and (force or now >= job.next_refresh):
                job.future = self._pool.submit(self._run, loader)
                job.started = now
                job.timed_out = False
                # Keep any previous error visible until a successful refresh.

            age = None if job.fetched is None else max(0.0, now - job.fetched)
            value = job.value if age is not None and age <= self.stale_ttl else None
            return RefreshState(deepcopy(value), job.future is not None and not job.timed_out,
                                job.error, job.revision, age)


live_refresh = BackgroundRefresh()
station_map_refresh = BackgroundRefresh(ttl=6 * 60 * 60, stale_ttl=24 * 60 * 60)
