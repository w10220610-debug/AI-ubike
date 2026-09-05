from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Iterable

from battery_service import BatteryServiceError, get_station_battery_by_name
from config import BATTERY_MAX_CONCURRENCY


def query_station_for_ui(
    station_name: str,
    *,
    district: str = "",
    threshold: int = 89,
    priority_threshold: int = 40,
    force: bool = False,
) -> dict:
    threshold = max(0, min(100, int(threshold)))
    priority_threshold = max(0, min(threshold, int(priority_threshold)))
    result = get_station_battery_by_name(station_name, district=district, force=force)
    bikes = [dict(item) for item in result.get("bikes") or []]
    low = [item for item in bikes if int(item.get("battery_power", 101)) <= threshold]
    low.sort(key=lambda item: (str(item.get("pillar_no") or ""), int(item.get("battery_power", 101))))
    priority = [item for item in low if int(item.get("battery_power", 101)) <= priority_threshold]
    return {
        **result,
        "requested_name": station_name,
        "requested_district": district,
        "threshold": threshold,
        "priority_threshold": priority_threshold,
        "low_bikes": low,
        "priority_bikes": priority,
        "low_count": len(low),
        "priority_count": len(priority),
    }


def query_stations_for_ui(
    stations: Iterable[dict],
    *,
    threshold: int = 89,
    priority_threshold: int = 40,
    force_names: set[str] | None = None,
) -> dict[str, dict]:
    specs = []
    seen = set()
    for item in stations:
        name = str(item.get("name") or item.get("station_name") or "").strip()
        district = str(item.get("district") or "").strip()
        if not name or name in seen:
            continue
        seen.add(name)
        specs.append((name, district))
    force_names = force_names or set()
    output: dict[str, dict] = {}
    workers = max(1, min(BATTERY_MAX_CONCURRENCY, len(specs) or 1))
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="battery-ui") as pool:
        futures = {
            pool.submit(
                query_station_for_ui,
                name,
                district=district,
                threshold=threshold,
                priority_threshold=priority_threshold,
                force=name in force_names,
            ): name
            for name, district in specs
        }
        for future in as_completed(futures):
            name = futures[future]
            try:
                output[name] = future.result()
            except BatteryServiceError as exc:
                output[name] = {"requested_name": name, "error": str(exc), "bikes": [], "low_bikes": [], "priority_bikes": [], "low_count": 0, "priority_count": 0}
            except Exception as exc:  # UI 不應因單一站查詢失敗而整頁中斷
                output[name] = {"requested_name": name, "error": str(exc), "bikes": [], "low_bikes": [], "priority_bikes": [], "low_count": 0, "priority_count": 0}
    return output
