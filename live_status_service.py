from __future__ import annotations

import json
import logging
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, replace
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

from config import (
    APP_VERSION,
    BATTERY_RETRY_BACKOFF_SECONDS,
    LIVE_STATUS_BATCH_SIZE,
    LIVE_STATUS_CACHE_TTL_SECONDS,
    LIVE_STATUS_MAX_ATTEMPTS,
    LIVE_STATUS_MAX_CONCURRENCY,
    LIVE_STATUS_MISSING_RETRY_ROUNDS,
    LIVE_STATUS_REQUEST_TIMEOUT_SECONDS,
    LIVE_STATUS_STALE_TTL_SECONDS,
    YOUBIKE_LIVE_STATUS_URL,
)
from station_service import match_station


class LiveStatusServiceError(RuntimeError):
    pass


@dataclass(frozen=True)
class LiveStatusSnapshot:
    cache_key: tuple[str, ...]
    records: tuple[dict, ...]
    fetched_at_epoch: float
    event_id: str
    latest_source_time: str = ""
    request_count: int = 0
    failed_request_count: int = 0
    batch_round_count: int = 0
    missing_station_ids: tuple[str, ...] = ()
    unmatched_station_names: tuple[str, ...] = ()
    api_latency_ms: int = 0
    source: str = "server_live"
    error: str | None = None

    @property
    def age_seconds(self) -> float:
        return max(0.0, time.time() - self.fetched_at_epoch)

    def as_dict(self) -> dict:
        fetched_at = datetime.fromtimestamp(
            self.fetched_at_epoch,
            ZoneInfo("Asia/Taipei"),
        ).strftime("%Y/%m/%d %H:%M:%S")
        return {
            "ok": True,
            "records": [dict(item) for item in self.records],
            "fetched_at": fetched_at,
            "fetched_at_epoch": self.fetched_at_epoch,
            "latest_source_time": self.latest_source_time,
            "station_count": len(self.records),
            "requested_station_count": len(self.cache_key),
            "missing_station_count": len(self.missing_station_ids),
            "missing_station_ids": list(self.missing_station_ids),
            "unmatched_station_count": len(self.unmatched_station_names),
            "unmatched_station_names": list(self.unmatched_station_names),
            "request_batch_count": self.request_count,
            "request_count": self.request_count,
            "failed_request_count": self.failed_request_count,
            "batch_round_count": self.batch_round_count,
            "single_round_count": 0,
            "batch_size": LIVE_STATUS_BATCH_SIZE,
            "request_concurrency": LIVE_STATUS_MAX_CONCURRENCY,
            "elapsed_ms": self.api_latency_ms,
            "event_id": self.event_id,
            "source": "YouBike 官網公開接口（Python Server 共用快取）",
            "cache_source": self.source,
            "error": self.error,
            "age_seconds": self.age_seconds,
        }


_cache: dict[tuple[str, ...], LiveStatusSnapshot] = {}
_cache_guard = threading.RLock()
_request_locks: dict[tuple[str, ...], threading.Lock] = {}
_request_locks_guard = threading.Lock()
_api_semaphore = threading.BoundedSemaphore(LIVE_STATUS_MAX_CONCURRENCY)


def _build_logger() -> logging.Logger:
    logger = logging.getLogger("ai_ubike.live_status")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    log_dir = Path("logs")
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        handler: logging.Handler = RotatingFileHandler(
            log_dir / "live_status.log",
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


def _request_lock(cache_key: tuple[str, ...]) -> threading.Lock:
    with _request_locks_guard:
        lock = _request_locks.get(cache_key)
        if lock is None:
            lock = threading.Lock()
            _request_locks[cache_key] = lock
        return lock


def _first_nonempty(*values):
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return value
    return None


def _extract_items(payload: Any) -> list[dict]:
    containers = [payload]
    if isinstance(payload, dict):
        containers.extend((payload.get("retVal"), payload.get("data"), payload.get("result")))
    for container in containers:
        if isinstance(container, list):
            return [item for item in container if isinstance(item, dict)]
        if not isinstance(container, dict):
            continue
        for key in ("data", "items", "result", "stations", "retVal"):
            value = container.get(key)
            if isinstance(value, list):
                return [item for item in value if isinstance(item, dict)]
            if isinstance(value, dict):
                nested = value.get("data") or value.get("items")
                if isinstance(nested, list):
                    return [item for item in nested if isinstance(item, dict)]
    return []


def _nonnegative_int_or_none(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        return max(0, int(float(value)))
    except (TypeError, ValueError, OverflowError):
        return None


def _is_retryable_http_status(status: int) -> bool:
    return status in {408, 425, 429, 500, 502, 503, 504}


def _request_batch(station_numbers: list[str]) -> tuple[list[dict], int, int]:
    encoded_body = json.dumps({"station_no": station_numbers}).encode("utf-8")
    last_error: BaseException | None = None
    latency_ms = 0
    for attempt in range(1, LIVE_STATUS_MAX_ATTEMPTS + 1):
        started = time.perf_counter()
        request = Request(
            YOUBIKE_LIVE_STATUS_URL,
            data=encoded_body,
            method="POST",
            headers={
                "Accept": "application/json, text/plain, */*",
                "Content-Type": "application/json;charset=UTF-8",
                "Origin": "https://www.youbike.com.tw",
                "Referer": "https://www.youbike.com.tw/region/taitung/stations/",
                "User-Agent": f"AI-UBIKE/{APP_VERSION} live-status-service",
            },
        )
        try:
            with _api_semaphore:
                with urlopen(request, timeout=LIVE_STATUS_REQUEST_TIMEOUT_SECONDS) as response:
                    payload = json.loads(response.read().decode("utf-8-sig"))
            latency_ms = int((time.perf_counter() - started) * 1000)
            if isinstance(payload, dict) and payload.get("retCode") not in (None, 1, "1", True):
                raise LiveStatusServiceError(str(payload.get("retMsg") or "官方資料服務回傳失敗"))
            return _extract_items(payload), latency_ms, attempt
        except HTTPError as exc:
            latency_ms = int((time.perf_counter() - started) * 1000)
            last_error = exc
            retryable = _is_retryable_http_status(int(exc.code))
            _logger.warning(
                "[LIVE_STATUS_API] stations=%s status=http_%s attempt=%s/%s latency_ms=%s retryable=%s",
                len(station_numbers), exc.code, attempt, LIVE_STATUS_MAX_ATTEMPTS, latency_ms, retryable,
            )
            if not retryable:
                break
        except (URLError, TimeoutError, json.JSONDecodeError, UnicodeDecodeError, OSError) as exc:
            latency_ms = int((time.perf_counter() - started) * 1000)
            last_error = exc
            _logger.warning(
                "[LIVE_STATUS_API] stations=%s status=%s attempt=%s/%s latency_ms=%s",
                len(station_numbers), type(exc).__name__, attempt, LIVE_STATUS_MAX_ATTEMPTS, latency_ms,
            )
        except LiveStatusServiceError as exc:
            last_error = exc
            break

        if attempt < LIVE_STATUS_MAX_ATTEMPTS:
            delay_index = min(attempt - 1, len(BATTERY_RETRY_BACKOFF_SECONDS) - 1)
            time.sleep(BATTERY_RETRY_BACKOFF_SECONDS[delay_index])

    raise LiveStatusServiceError(f"即時車數 API 查詢失敗：{last_error}")


def _chunks(values: list[str]) -> list[list[str]]:
    return [
        values[start : start + LIVE_STATUS_BATCH_SIZE]
        for start in range(0, len(values), LIVE_STATUS_BATCH_SIZE)
    ]


def _fetch_parking_records(station_numbers: list[str]) -> tuple[dict[str, dict], dict]:
    pending = list(dict.fromkeys(station_numbers))
    found: dict[str, dict] = {}
    request_count = 0
    failed_request_count = 0
    total_latency_ms = 0
    last_errors: list[str] = []
    batch_round_count = 0

    for _round in range(max(1, LIVE_STATUS_MISSING_RETRY_ROUNDS)):
        batches = _chunks(pending)
        if not batches:
            break
        batch_round_count += 1
        workers = max(1, min(LIVE_STATUS_MAX_CONCURRENCY, len(batches)))
        with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="live-status") as pool:
            future_map = {pool.submit(_request_batch, batch): batch for batch in batches}
            for future in as_completed(future_map):
                request_count += 1
                try:
                    items, latency_ms, _attempts = future.result()
                    total_latency_ms += latency_ms
                    for item in items:
                        station_no = str(
                            _first_nonempty(item.get("station_no"), item.get("sno")) or ""
                        ).strip()
                        if station_no:
                            found[station_no] = item
                except LiveStatusServiceError as exc:
                    failed_request_count += 1
                    last_errors.append(str(exc))
        pending = [station_no for station_no in station_numbers if station_no not in found]

    if not found and last_errors:
        raise LiveStatusServiceError(last_errors[-1])
    return found, {
        "request_count": request_count,
        "failed_request_count": failed_request_count,
        "batch_round_count": batch_round_count,
        "missing_station_ids": pending,
        "api_latency_ms": total_latency_ms,
    }


def _prepare_station_records(stations: Iterable[dict]) -> tuple[list[dict], list[str]]:
    matched: list[dict] = []
    unmatched: list[str] = []
    seen: set[str] = set()
    for item in stations:
        name = str(item.get("name") or item.get("station_name") or "").strip()
        district = str(item.get("district") or "").strip()
        if not name:
            continue
        match = match_station(name, district=district)
        if match is None:
            unmatched.append(name)
            continue
        station_no = str(match.get("station_no") or "").strip()
        if not station_no or station_no in seen:
            continue
        seen.add(station_no)
        matched.append(match)
    return matched, unmatched


def _normalize_live_record(station: dict, parking: dict) -> dict:
    detail = _first_nonempty(parking.get("available_spaces_detail"), parking.get("sbi_detail"))
    if not isinstance(detail, dict):
        detail = {}
    source_update_time = str(
        _first_nonempty(
            parking.get("updated_at"),
            parking.get("mday"),
            parking.get("time"),
        )
        or ""
    ).strip()
    station_no = str(station.get("station_no") or station.get("station_id") or "").strip()
    return {
        "station_uid": station_no,
        "station_id": station_no,
        "station_name": str(station.get("station_name") or "").strip(),
        "station_key": str(station.get("station_key") or "").strip(),
        "service_status": _nonnegative_int_or_none(
            _first_nonempty(parking.get("status"), parking.get("act"), station.get("service_status"), 1)
        ),
        "general_bikes": _nonnegative_int_or_none(detail.get("yb2")),
        "electric_bikes": _nonnegative_int_or_none(detail.get("eyb")),
        "available_spaces": _nonnegative_int_or_none(
            _first_nonempty(parking.get("available_spaces"), parking.get("sbi"))
        ),
        "empty_spaces": _nonnegative_int_or_none(
            _first_nonempty(parking.get("empty_spaces"), parking.get("bemp"))
        ),
        "parking_spaces": _nonnegative_int_or_none(
            _first_nonempty(parking.get("parking_spaces"), parking.get("tot"))
        ),
        "source_update_time": source_update_time,
        "latitude": station.get("latitude"),
        "longitude": station.get("longitude"),
    }


def _fresh(snapshot: LiveStatusSnapshot, now: float) -> bool:
    return now - snapshot.fetched_at_epoch <= LIVE_STATUS_CACHE_TTL_SECONDS


def _stale_usable(snapshot: LiveStatusSnapshot, now: float) -> bool:
    return now - snapshot.fetched_at_epoch <= LIVE_STATUS_STALE_TTL_SECONDS


def get_live_status_for_stations(stations: Iterable[dict], *, force: bool = False) -> dict:
    matched, unmatched = _prepare_station_records(stations)
    station_numbers = [str(item["station_no"]) for item in matched]
    cache_key = tuple(sorted(station_numbers))
    if not cache_key:
        raise LiveStatusServiceError("沒有任何 Excel 場站可安全配對到 YouBike 場站清單。")

    request_started_at = time.time()
    with _cache_guard:
        cached = _cache.get(cache_key)
    if cached and not force and _fresh(cached, request_started_at):
        _logger.info(
            "[LIVE_STATUS_API] stations=%s cache=hit age=%.2fs",
            len(cache_key), cached.age_seconds,
        )
        return cached.as_dict()

    with _request_lock(cache_key):
        now = time.time()
        with _cache_guard:
            cached = _cache.get(cache_key)
        if cached and (
            (not force and _fresh(cached, now))
            or (force and cached.fetched_at_epoch >= request_started_at)
        ):
            return cached.as_dict()

        started = time.perf_counter()
        try:
            parking_by_station, metadata = _fetch_parking_records(station_numbers)
            records = tuple(
                _normalize_live_record(station, parking_by_station[station_no])
                for station in matched
                if (station_no := str(station["station_no"])) in parking_by_station
            )
            if not records:
                raise LiveStatusServiceError("YouBike 官網沒有回傳可用的即時車數。")
            source_times = [str(item.get("source_update_time") or "") for item in records]
            snapshot = LiveStatusSnapshot(
                cache_key=cache_key,
                records=records,
                fetched_at_epoch=time.time(),
                event_id=uuid.uuid4().hex,
                latest_source_time=max((value for value in source_times if value), default=""),
                request_count=int(metadata["request_count"]),
                failed_request_count=int(metadata["failed_request_count"]),
                batch_round_count=int(metadata["batch_round_count"]),
                missing_station_ids=tuple(metadata["missing_station_ids"]),
                unmatched_station_names=tuple(unmatched),
                api_latency_ms=int((time.perf_counter() - started) * 1000),
            )
            with _cache_guard:
                _cache[cache_key] = snapshot
            _logger.info(
                "[LIVE_STATUS_API] stations=%s returned=%s missing=%s unmatched=%s cache=miss "
                "requests=%s failed_requests=%s latency_ms=%s",
                len(cache_key), len(records), len(snapshot.missing_station_ids), len(unmatched),
                snapshot.request_count, snapshot.failed_request_count, snapshot.api_latency_ms,
            )
            return snapshot.as_dict()
        except LiveStatusServiceError as exc:
            now = time.time()
            if cached and _stale_usable(cached, now):
                stale = replace(cached, source="stale_cache", error=str(exc))
                _logger.warning(
                    "[LIVE_STATUS_API] stations=%s status=failed fallback=stale_cache cache_age=%.2fs error=%s",
                    len(cache_key), cached.age_seconds, exc,
                )
                return stale.as_dict()
            _logger.error(
                "[LIVE_STATUS_API] stations=%s status=failed fallback=none error=%s",
                len(cache_key), exc,
            )
            raise


def clear_live_status_cache() -> None:
    with _cache_guard:
        _cache.clear()


def get_live_status_cache_status() -> dict:
    now = time.time()
    with _cache_guard:
        snapshots = list(_cache.values())
    return {
        "count": len(snapshots),
        "entries": [
            {
                "station_count": len(snapshot.cache_key),
                "age_seconds": max(0.0, now - snapshot.fetched_at_epoch),
                "state": (
                    "fresh"
                    if now - snapshot.fetched_at_epoch <= LIVE_STATUS_CACHE_TTL_SECONDS
                    else "stale"
                    if now - snapshot.fetched_at_epoch <= LIVE_STATUS_STALE_TTL_SECONDS
                    else "expired"
                ),
            }
            for snapshot in snapshots
        ],
    }
