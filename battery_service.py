from __future__ import annotations

import json
import logging
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, replace
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from config import (
    BATTERY_CACHE_TTL_SECONDS,
    BATTERY_MAX_ATTEMPTS,
    BATTERY_MAX_CONCURRENCY,
    BATTERY_REQUEST_TIMEOUT_SECONDS,
    BATTERY_RETRY_BACKOFF_SECONDS,
    BATTERY_STALE_TTL_SECONDS,
    DEFAULT_BATTERY_PRIORITY_THRESHOLD,
    DEFAULT_BATTERY_THRESHOLD,
    YOUBIKE_BATTERY_URL,
)
from station_service import get_station_catalog, match_station


class BatteryServiceError(RuntimeError):
    pass


@dataclass(frozen=True)
class BatterySnapshot:
    station_no: str
    station_name: str
    bikes: tuple[dict, ...]
    fetched_at: float
    source: str = "live"
    error: str | None = None
    api_latency_ms: int | None = None

    @property
    def age_seconds(self) -> float:
        return max(0.0, time.time() - self.fetched_at)

    def as_dict(self) -> dict:
        return {
            "station_no": self.station_no,
            "station_name": self.station_name,
            "bikes": [dict(item) for item in self.bikes],
            "fetched_at": self.fetched_at,
            "source": self.source,
            "error": self.error,
            "api_latency_ms": self.api_latency_ms,
            "age_seconds": self.age_seconds,
        }


_cache: dict[str, BatterySnapshot] = {}
_cache_guard = threading.RLock()
_station_locks: dict[str, threading.Lock] = {}
_station_locks_guard = threading.Lock()
_api_semaphore = threading.BoundedSemaphore(BATTERY_MAX_CONCURRENCY)


def _build_logger() -> logging.Logger:
    logger = logging.getLogger("ai_ubike.battery")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    log_dir = Path("logs")
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        handler: logging.Handler = RotatingFileHandler(
            log_dir / "battery.log",
            maxBytes=2_000_000,
            backupCount=3,
            encoding="utf-8",
        )
    except OSError:
        handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
    logger.propagate = False
    return logger


_logger = _build_logger()


def _station_lock(station_no: str) -> threading.Lock:
    with _station_locks_guard:
        lock = _station_locks.get(station_no)
        if lock is None:
            lock = threading.Lock()
            _station_locks[station_no] = lock
        return lock


def _extract_records(payload: Any) -> list[dict]:
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    if not isinstance(payload, dict):
        return []
    for key in ("retVal", "data", "items", "result", "results"):
        value = payload.get(key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
        if isinstance(value, dict):
            for nested_key in ("items", "data", "list", "results"):
                nested = value.get(nested_key)
                if isinstance(nested, list):
                    return [item for item in nested if isinstance(item, dict)]
    return []


def _safe_battery_power(value: Any) -> int | None:
    try:
        number = int(float(value))
    except (TypeError, ValueError, OverflowError):
        return None
    return max(0, min(100, number))


def _normalize_bikes(payload: Any) -> tuple[dict, ...]:
    bikes: list[dict] = []
    for record in _extract_records(payload):
        battery_power = _safe_battery_power(record.get("battery_power"))
        bike_no = str(record.get("bike_no") or "").strip()
        if not bike_no or battery_power is None:
            continue
        bikes.append(
            {
                "bike_no": bike_no,
                "pillar_no": str(record.get("pillar_no") or "").strip(),
                "battery_power": battery_power,
            }
        )
    return tuple(bikes)


def _is_retryable_http_status(status: int) -> bool:
    return status in {408, 425, 429, 500, 502, 503, 504}


def _request_station_battery(station_no: str) -> tuple[tuple[dict, ...], int]:
    last_error: BaseException | None = None
    latency_ms = 0
    for attempt in range(1, BATTERY_MAX_ATTEMPTS + 1):
        started = time.perf_counter()
        query = urlencode({"station_no": station_no})
        request = Request(
            f"{YOUBIKE_BATTERY_URL}?{query}",
            headers={
                "Accept": "application/json, text/plain, */*",
                "User-Agent": "AI-UBIKE/30 battery-service",
            },
        )
        try:
            with _api_semaphore:
                with urlopen(request, timeout=BATTERY_REQUEST_TIMEOUT_SECONDS) as response:
                    payload = json.loads(response.read().decode("utf-8"))
            latency_ms = int((time.perf_counter() - started) * 1000)
            return _normalize_bikes(payload), latency_ms
        except HTTPError as exc:
            latency_ms = int((time.perf_counter() - started) * 1000)
            last_error = exc
            retryable = _is_retryable_http_status(int(exc.code))
            _logger.warning(
                "[BATTERY_API] station_no=%s status=http_%s attempt=%s/%s latency_ms=%s retryable=%s",
                station_no, exc.code, attempt, BATTERY_MAX_ATTEMPTS, latency_ms, retryable,
            )
            if not retryable:
                break
        except (URLError, TimeoutError, json.JSONDecodeError, UnicodeDecodeError, OSError) as exc:
            latency_ms = int((time.perf_counter() - started) * 1000)
            last_error = exc
            _logger.warning(
                "[BATTERY_API] station_no=%s status=%s attempt=%s/%s latency_ms=%s",
                station_no, type(exc).__name__, attempt, BATTERY_MAX_ATTEMPTS, latency_ms,
            )

        if attempt < BATTERY_MAX_ATTEMPTS:
            delay_index = min(attempt - 1, len(BATTERY_RETRY_BACKOFF_SECONDS) - 1)
            time.sleep(BATTERY_RETRY_BACKOFF_SECONDS[delay_index])

    raise BatteryServiceError(f"電池 API 查詢失敗 station_no={station_no}: {last_error}")


def _fresh(snapshot: BatterySnapshot, now: float) -> bool:
    return now - snapshot.fetched_at <= BATTERY_CACHE_TTL_SECONDS


def _stale_usable(snapshot: BatterySnapshot, now: float) -> bool:
    return now - snapshot.fetched_at <= BATTERY_STALE_TTL_SECONDS


def get_cached_station_battery(station_no: str) -> dict | None:
    key = str(station_no or "").strip()
    if not key:
        return None
    with _cache_guard:
        snapshot = _cache.get(key)
    return snapshot.as_dict() if snapshot else None


def get_station_battery(
    station_no: str,
    *,
    station_name: str = "",
    force: bool = False,
) -> dict:
    """取得單站電池資料。

    - 30 秒內：共用 fresh cache
    - 同站同時查：per-station lock 合併成一個真正 API request
    - API 暫時失敗：5 分鐘內可回 stale cache
    """
    key = str(station_no or "").strip()
    if not key:
        raise BatteryServiceError("station_no 不可為空。")

    request_started_at = time.time()
    with _cache_guard:
        cached = _cache.get(key)
    if cached and not force and _fresh(cached, request_started_at):
        _logger.info("[BATTERY_API] station_no=%s cache=hit age=%.2fs", key, cached.age_seconds)
        return cached.as_dict()

    lock = _station_lock(key)
    with lock:
        now = time.time()
        with _cache_guard:
            cached = _cache.get(key)
        # single-flight：等待 lock 期間若別人已更新，force 請求也共用那次更新。
        if cached and (
            (not force and _fresh(cached, now))
            or (force and cached.fetched_at >= request_started_at)
        ):
            return cached.as_dict()

        try:
            bikes, latency_ms = _request_station_battery(key)
            snapshot = BatterySnapshot(
                station_no=key,
                station_name=str(station_name or (cached.station_name if cached else "")),
                bikes=bikes,
                fetched_at=time.time(),
                source="live",
                api_latency_ms=latency_ms,
            )
            with _cache_guard:
                _cache[key] = snapshot
            _logger.info(
                "[BATTERY_API] station=%s station_no=%s status=ok cache=miss bikes=%s latency_ms=%s",
                snapshot.station_name, key, len(bikes), latency_ms,
            )
            return snapshot.as_dict()
        except BatteryServiceError as exc:
            now = time.time()
            if cached and _stale_usable(cached, now):
                stale = replace(cached, source="stale_cache", error=str(exc))
                _logger.warning(
                    "[BATTERY_API] station=%s station_no=%s status=failed fallback=stale_cache cache_age=%.2fs",
                    cached.station_name, key, cached.age_seconds,
                )
                return stale.as_dict()
            _logger.error(
                "[BATTERY_API] station=%s station_no=%s status=failed fallback=none error=%s",
                station_name, key, exc,
            )
            raise


def refresh_station_battery(station_no: str, *, station_name: str = "") -> dict:
    return get_station_battery(station_no, station_name=station_name, force=True)


def get_station_battery_by_name(station_name: str, *, district: str = "", force: bool = False) -> dict:
    catalog = get_station_catalog()
    match = match_station(station_name, district=district, catalog=catalog)
    if match is None:
        raise BatteryServiceError(f"找不到可安全配對的 YouBike 場站：{station_name}")
    result = get_station_battery(
        str(match["station_no"]),
        station_name=str(match["station_name"]),
        force=force,
    )
    result["match"] = {
        "excel_station_name": station_name,
        "station_no": match["station_no"],
        "station_name": match["station_name"],
        "match_score": match.get("match_score"),
        "match_method": match.get("match_method"),
    }
    return result


def get_low_battery_bikes(
    station_no: str,
    *,
    station_name: str = "",
    threshold: int = DEFAULT_BATTERY_THRESHOLD,
    priority_threshold: int = DEFAULT_BATTERY_PRIORITY_THRESHOLD,
    force: bool = False,
) -> dict:
    threshold = max(0, min(100, int(threshold)))
    priority_threshold = max(0, min(threshold, int(priority_threshold)))
    result = get_station_battery(station_no, station_name=station_name, force=force)
    bikes = list(result.get("bikes") or [])
    low = [bike for bike in bikes if int(bike["battery_power"]) <= threshold]
    priority = [bike for bike in low if int(bike["battery_power"]) <= priority_threshold]
    low.sort(key=lambda bike: (str(bike.get("pillar_no") or ""), int(bike["battery_power"])))
    priority.sort(key=lambda bike: (str(bike.get("pillar_no") or ""), int(bike["battery_power"])))
    return {
        **result,
        "threshold": threshold,
        "priority_threshold": priority_threshold,
        "low_bikes": low,
        "priority_bikes": priority,
        "low_count": len(low),
        "priority_count": len(priority),
    }


def get_battery_cache_status(station_no: str | None = None) -> dict:
    now = time.time()
    with _cache_guard:
        if station_no:
            snapshot = _cache.get(str(station_no).strip())
            snapshots = [snapshot] if snapshot else []
        else:
            snapshots = list(_cache.values())
    rows = []
    for snapshot in snapshots:
        age = max(0.0, now - snapshot.fetched_at)
        rows.append(
            {
                "station_no": snapshot.station_no,
                "station_name": snapshot.station_name,
                "age_seconds": age,
                "state": "fresh" if age <= BATTERY_CACHE_TTL_SECONDS else (
                    "stale" if age <= BATTERY_STALE_TTL_SECONDS else "expired"
                ),
                "source": snapshot.source,
            }
        )
    return {"count": len(rows), "entries": rows}


def clear_battery_cache(station_no: str | None = None) -> None:
    with _cache_guard:
        if station_no:
            _cache.pop(str(station_no).strip(), None)
        else:
            _cache.clear()


def prefetch_station_batteries(
    stations: Iterable[dict],
    *,
    max_workers: int = BATTERY_MAX_CONCURRENCY,
) -> dict[str, dict]:
    """只預取 UI 真正需要顯示的小量候選站，例如智慧調度前 10 站。"""
    specs = []
    seen = set()
    for item in stations:
        station_no = str(item.get("station_no") or "").strip()
        if not station_no or station_no in seen:
            continue
        seen.add(station_no)
        specs.append((station_no, str(item.get("station_name") or item.get("name") or "")))

    output: dict[str, dict] = {}
    workers = max(1, min(int(max_workers), BATTERY_MAX_CONCURRENCY, len(specs) or 1))
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="battery-prefetch") as pool:
        future_map = {
            pool.submit(get_station_battery, station_no, station_name=station_name): station_no
            for station_no, station_name in specs
        }
        for future in as_completed(future_map):
            station_no = future_map[future]
            try:
                output[station_no] = future.result()
            except BatteryServiceError as exc:
                output[station_no] = {"station_no": station_no, "error": str(exc), "bikes": []}
    return output
