from __future__ import annotations

import json
import re
import threading
import time
import unicodedata
from difflib import SequenceMatcher
from functools import lru_cache
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from config import (
    STATION_CATALOG_TTL_SECONDS,
    STATION_MATCH_AMBIGUITY_MARGIN,
    STATION_MATCH_THRESHOLD,
    YOUBIKE_STATION_CATALOG_URL,
)


class StationServiceError(RuntimeError):
    pass


_catalog_lock = threading.Lock()
_catalog_cache: tuple[float, tuple[dict, ...]] | None = None


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
        containers.extend([payload.get("retVal"), payload.get("data"), payload.get("result")])
    for container in containers:
        if isinstance(container, list):
            return [item for item in container if isinstance(item, dict)]
        if isinstance(container, dict):
            for key in ("data", "items", "result", "stations", "retVal"):
                value = container.get(key)
                if isinstance(value, list):
                    return [item for item in value if isinstance(item, dict)]
                if isinstance(value, dict):
                    nested = value.get("data") or value.get("items")
                    if isinstance(nested, list):
                        return [item for item in nested if isinstance(item, dict)]
    return []


@lru_cache(maxsize=32768)
def normalize_station_name(raw: str) -> str:
    text = unicodedata.normalize("NFKC", str(raw or "")).strip().lower().replace("臺", "台")
    text = re.sub(
        r"^(?:youbike|ubike)\s*2\s*[.．]?\s*0\s*e?\s*[_\-－—:：]*\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )
    text = text.replace("公共自行車租賃站", "")
    return re.sub(r"[^0-9a-z\u3400-\u9fff]", "", text)


def _normalize_district(value: str) -> str:
    return normalize_station_name(value).replace("台東縣", "").replace("台東市", "台東市")


def _similarity(a: str, b: str) -> float:
    left = normalize_station_name(a)
    right = normalize_station_name(b)
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    if min(len(left), len(right)) >= 4 and (left in right or right in left):
        return 0.96
    return SequenceMatcher(None, left, right, autojunk=False).ratio()


def _request_catalog() -> tuple[dict, ...]:
    request = Request(
        YOUBIKE_STATION_CATALOG_URL,
        headers={"Accept": "application/json, text/plain, */*", "User-Agent": "AI-UBIKE/30 station-service"},
    )
    try:
        with urlopen(request, timeout=15) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (HTTPError, URLError, TimeoutError, OSError, json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise StationServiceError(f"YouBike 場站清單讀取失敗：{exc}") from exc

    records: list[dict] = []
    for item in _extract_items(payload):
        station_no = str(_first_nonempty(item.get("station_no"), item.get("sno"), item.get("station_id")) or "").strip()
        station_name = str(_first_nonempty(item.get("name_tw"), item.get("sna"), item.get("station_name")) or "").strip()
        if not station_no or not station_name:
            continue
        county = str(_first_nonempty(item.get("county_tw"), item.get("city_tw"), item.get("scity")) or "").strip()
        district = str(_first_nonempty(item.get("district_tw"), item.get("sarea")) or "").strip()
        address = str(_first_nonempty(item.get("address_tw"), item.get("ar")) or "").strip()
        records.append(
            {
                "station_no": station_no,
                "station_id": station_no,
                "station_name": station_name,
                "station_key": normalize_station_name(station_name),
                "county": county,
                "district": district,
                "address": address,
                "latitude": _first_nonempty(item.get("lat"), item.get("latitude")),
                "longitude": _first_nonempty(item.get("lng"), item.get("longitude")),
                "service_status": _first_nonempty(item.get("status"), item.get("act"), 1),
            }
        )
    if not records:
        raise StationServiceError("YouBike 場站清單沒有可辨識的資料。")
    return tuple(records)


def get_station_catalog(*, force: bool = False) -> tuple[dict, ...]:
    global _catalog_cache
    now = time.time()
    cached = _catalog_cache
    if cached and not force and now - cached[0] <= STATION_CATALOG_TTL_SECONDS:
        return cached[1]
    with _catalog_lock:
        cached = _catalog_cache
        now = time.time()
        if cached and not force and now - cached[0] <= STATION_CATALOG_TTL_SECONDS:
            return cached[1]
        records = _request_catalog()
        _catalog_cache = (time.time(), records)
        return records


def _district_bonus(record: dict, district: str) -> float:
    wanted = _normalize_district(district)
    if not wanted:
        return 0.0
    text = " ".join(
        str(record.get(key) or "")
        for key in ("county", "district", "address", "station_name")
    )
    normalized = _normalize_district(text)
    return 0.08 if wanted and wanted in normalized else 0.0


def match_station(
    station_name: str,
    *,
    district: str = "",
    catalog: Iterable[dict] | None = None,
) -> dict | None:
    records = tuple(catalog) if catalog is not None else get_station_catalog()
    wanted_key = normalize_station_name(station_name)
    if not wanted_key:
        return None

    exact = [record for record in records if str(record.get("station_key") or "") == wanted_key]
    if district and len(exact) > 1:
        district_matches = [record for record in exact if _district_bonus(record, district) > 0]
        if len(district_matches) == 1:
            chosen = dict(district_matches[0])
            chosen.update({"match_score": 1.08, "match_method": "exact+district"})
            return chosen
    if len(exact) == 1:
        chosen = dict(exact[0])
        chosen.update({"match_score": 1.0, "match_method": "exact"})
        return chosen
    if len(exact) > 1:
        return None

    ranked: list[tuple[float, dict]] = []
    for record in records:
        score = _similarity(station_name, str(record.get("station_name") or "")) + _district_bonus(record, district)
        if score >= STATION_MATCH_THRESHOLD:
            ranked.append((score, record))
    if not ranked:
        return None
    ranked.sort(key=lambda item: item[0], reverse=True)
    best_score, best = ranked[0]
    second_score = ranked[1][0] if len(ranked) > 1 else -1.0
    if best_score < 0.96 and second_score >= best_score - STATION_MATCH_AMBIGUITY_MARGIN:
        return None
    chosen = dict(best)
    chosen.update({"match_score": round(best_score, 4), "match_method": "fuzzy"})
    return chosen


def match_station_specs(stations: Iterable[dict]) -> list[dict]:
    catalog = get_station_catalog()
    output = []
    for item in stations:
        name = str(item.get("name") or item.get("station_name") or "").strip()
        district = str(item.get("district") or "").strip()
        if not name:
            continue
        matched = match_station(name, district=district, catalog=catalog)
        if matched is None:
            output.append({**item, "match_error": "unmatched"})
        else:
            output.append({**item, **matched})
    return output
