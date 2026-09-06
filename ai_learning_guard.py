from __future__ import annotations

import html as html_lib
import inspect
import json
import math
import os
import re
import threading
import time
import unicodedata
import uuid
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from streamlit_component_compat import install_component_declare_compat


# Install before legacy_ui.py is exec()'d by app.py.
install_component_declare_compat()


# Learning guardrails.
SUSPECTED_SINGLE_TYPE_DELTA = 6
SUSPECTED_TOTAL_ABS_DELTA = 8
MANUAL_EVENT_WINDOW_SECONDS = 30 * 60
MAX_MANUAL_EVENTS = 500
MAX_TRANSITION_RECORDS = 50000
PREDICTION_HORIZON_MINUTES = 60

# Shared learning pool.
SHARED_POOL_SYNC_SECONDS = 10.0
_CACHE_DIR = Path(__file__).resolve().parent / ".base_cache"
SHARED_LEARNING_POOL_PATH = _CACHE_DIR / "ai_learning_pool.shared.json"

# Shared dispatcher-location pool. Every active browser/session gets an anonymous
# device id. Fresh GPS fixes are merged here so either dispatcher can help mark a
# station change as manual intervention.
SHARED_DISPATCHER_LOCATION_PATH = _CACHE_DIR / "ai_dispatcher_locations.shared.json"
DISPATCHER_LOCATION_MAX_AGE_SECONDS = 180.0
DISPATCHER_CONFIRMED_RADIUS_METERS = 60.0
DISPATCHER_SUSPECTED_RADIUS_METERS = 140.0
DISPATCHER_AMBIGUITY_GAP_METERS = 25.0
DISPATCHER_MAX_USABLE_ACCURACY_METERS = 150.0
SHARED_GEOLOCATION_STATE_PREFIX = "shared_geolocation::"

_LATEST_LEARNING_RECORDS: list[dict] = []
_SHARED_POOL_CACHE: list[dict] = []
_SHARED_POOL_LAST_CHECK_MONOTONIC = 0.0
_SHARED_POOL_MTIME_NS: int | None = None
_SHARED_POOL_LOCK = threading.RLock()
_PUSHED_RECORD_IDS: set[str] = set()

_LOCATION_POOL_CACHE: dict[str, dict] = {}
_LOCATION_POOL_MTIME_NS: int | None = None
_LOCATION_POOL_LOCK = threading.RLock()


def _station_key(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).strip().replace("臺", "台")
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", text).lower()


def _int_or_none(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    if not math.isfinite(number):
        return None
    return max(0, int(number))


def _float_or_none(value: Any) -> float | None:
    try:
        number = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    return number if math.isfinite(number) else None


def _records_by_station(frame: pd.DataFrame) -> dict[str, dict]:
    if not isinstance(frame, pd.DataFrame) or frame.empty or "場站名稱" not in frame.columns:
        return {}
    output: dict[str, dict] = {}
    for row in frame.to_dict(orient="records"):
        station_name = str(row.get("場站名稱") or "").strip()
        key = _station_key(station_name)
        if not key:
            continue
        output[key] = {
            "station_name": station_name,
            "bike": _int_or_none(row.get("2.0 現況")),
            "ebike": _int_or_none(row.get("2.0E 現況")),
        }
    return output


def _record_identity(record: dict) -> str:
    record_id = str(record.get("record_id") or "").strip()
    if record_id:
        return record_id
    stable = (
        str(record.get("station_key") or _station_key(record.get("station_name"))),
        str(record.get("observed_at_epoch") or ""),
        str(record.get("bike_delta") or ""),
        str(record.get("ebike_delta") or ""),
        str(record.get("classification") or ""),
        str(record.get("source_event_id") or ""),
    )
    return "legacy:" + "|".join(stable)


def _sort_trim_records(records: list[dict] | None) -> list[dict]:
    deduped: dict[str, dict] = {}
    for item in records or []:
        if isinstance(item, dict):
            deduped[_record_identity(item)] = dict(item)

    def sort_key(item: dict) -> tuple[float, str]:
        observed = _float_or_none(item.get("observed_at_epoch")) or 0.0
        return observed, _record_identity(item)

    return sorted(deduped.values(), key=sort_key)[-MAX_TRANSITION_RECORDS:]


def _read_shared_pool(*, force: bool = False) -> list[dict]:
    global _SHARED_POOL_CACHE, _SHARED_POOL_LAST_CHECK_MONOTONIC, _SHARED_POOL_MTIME_NS

    now_mono = time.monotonic()
    with _SHARED_POOL_LOCK:
        if (
            not force
            and _SHARED_POOL_CACHE
            and now_mono - _SHARED_POOL_LAST_CHECK_MONOTONIC < SHARED_POOL_SYNC_SECONDS
        ):
            return [dict(item) for item in _SHARED_POOL_CACHE]

        _SHARED_POOL_LAST_CHECK_MONOTONIC = now_mono
        try:
            current_mtime_ns = SHARED_LEARNING_POOL_PATH.stat().st_mtime_ns
        except OSError:
            current_mtime_ns = None

        if (
            not force
            and _SHARED_POOL_CACHE
            and current_mtime_ns is not None
            and current_mtime_ns == _SHARED_POOL_MTIME_NS
        ):
            return [dict(item) for item in _SHARED_POOL_CACHE]

        try:
            raw = json.loads(SHARED_LEARNING_POOL_PATH.read_text(encoding="utf-8"))
            records = raw.get("records", []) if isinstance(raw, dict) else []
            loaded = _sort_trim_records(records if isinstance(records, list) else [])
        except (OSError, ValueError, TypeError):
            loaded = []

        _SHARED_POOL_CACHE = loaded
        _SHARED_POOL_MTIME_NS = current_mtime_ns
        return [dict(item) for item in loaded]


def _write_shared_pool(records: list[dict]) -> list[dict]:
    global _SHARED_POOL_CACHE, _SHARED_POOL_LAST_CHECK_MONOTONIC, _SHARED_POOL_MTIME_NS

    normalized = _sort_trim_records(records)
    SHARED_LEARNING_POOL_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": 2,
        "updated_at_epoch": time.time(),
        "max_records": MAX_TRANSITION_RECORDS,
        "records": normalized,
    }
    encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    temp_path = SHARED_LEARNING_POOL_PATH.with_name(
        f".{SHARED_LEARNING_POOL_PATH.name}.{uuid.uuid4().hex}.tmp"
    )
    try:
        temp_path.write_text(encoded, encoding="utf-8")
        os.replace(temp_path, SHARED_LEARNING_POOL_PATH)
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass

    try:
        _SHARED_POOL_MTIME_NS = SHARED_LEARNING_POOL_PATH.stat().st_mtime_ns
    except OSError:
        _SHARED_POOL_MTIME_NS = None
    _SHARED_POOL_CACHE = normalized
    _SHARED_POOL_LAST_CHECK_MONOTONIC = time.monotonic()
    return [dict(item) for item in normalized]


def sync_shared_learning_pool(
    incoming_records: list[dict] | None = None,
    *,
    force_read: bool = False,
) -> list[dict]:
    """Silently merge local learning into the shared AI pool."""
    global _LATEST_LEARNING_RECORDS

    with _SHARED_POOL_LOCK:
        existing = _read_shared_pool(force=force_read)
        incoming = [dict(item) for item in (incoming_records or []) if isinstance(item, dict)]
        if incoming:
            merged = _sort_trim_records([*existing, *incoming])
            existing_ids = {_record_identity(item) for item in existing}
            merged_ids = {_record_identity(item) for item in merged}
            shared = _write_shared_pool(merged) if merged_ids != existing_ids else merged
        else:
            shared = existing

        _LATEST_LEARNING_RECORDS = [dict(item) for item in shared]
        try:
            st.session_state["__ai_prediction_records__"] = [dict(item) for item in shared]
            st.session_state["__ai_shared_pool_count__"] = len(shared)
            st.session_state["__ai_shared_pool_synced_at__"] = time.time()
        except Exception:
            pass
        return [dict(item) for item in shared]


def _location_payload_is_valid(location: object) -> bool:
    if not isinstance(location, dict):
        return False
    lat = _float_or_none(location.get("latitude"))
    lon = _float_or_none(location.get("longitude"))
    return lat is not None and lon is not None and -90 <= lat <= 90 and -180 <= lon <= 180


def _session_device_id() -> str:
    try:
        device_id = str(st.session_state.get("__ai_dispatcher_device_id__") or "").strip()
        if not device_id:
            device_id = uuid.uuid4().hex
            st.session_state["__ai_dispatcher_device_id__"] = device_id
        return device_id
    except Exception:
        return "runtime-" + uuid.uuid4().hex


def _newest_session_location() -> dict | None:
    try:
        candidates: list[dict] = []
        for key, value in st.session_state.items():
            if not str(key).startswith(SHARED_GEOLOCATION_STATE_PREFIX):
                continue
            if _location_payload_is_valid(value):
                candidates.append(dict(value))
    except Exception:
        return None

    if not candidates:
        return None

    def updated(item: dict) -> float:
        return _float_or_none(item.get("updated_at")) or 0.0

    return max(candidates, key=updated)


def _read_dispatcher_location_pool(*, force: bool = False) -> dict[str, dict]:
    global _LOCATION_POOL_CACHE, _LOCATION_POOL_MTIME_NS

    with _LOCATION_POOL_LOCK:
        try:
            current_mtime_ns = SHARED_DISPATCHER_LOCATION_PATH.stat().st_mtime_ns
        except OSError:
            current_mtime_ns = None

        if (
            not force
            and _LOCATION_POOL_CACHE
            and current_mtime_ns is not None
            and current_mtime_ns == _LOCATION_POOL_MTIME_NS
        ):
            raw_entries = {key: dict(value) for key, value in _LOCATION_POOL_CACHE.items()}
        else:
            try:
                raw = json.loads(SHARED_DISPATCHER_LOCATION_PATH.read_text(encoding="utf-8"))
                entries = raw.get("devices", {}) if isinstance(raw, dict) else {}
                raw_entries = {
                    str(key): dict(value)
                    for key, value in (entries.items() if isinstance(entries, dict) else [])
                    if isinstance(value, dict)
                }
            except (OSError, ValueError, TypeError):
                raw_entries = {}
            _LOCATION_POOL_MTIME_NS = current_mtime_ns

        now_ts = time.time()
        cleaned: dict[str, dict] = {}
        for device_id, item in raw_entries.items():
            if not _location_payload_is_valid(item):
                continue
            updated_at = _float_or_none(item.get("updated_at")) or 0.0
            if updated_at <= 0 or now_ts - updated_at > DISPATCHER_LOCATION_MAX_AGE_SECONDS:
                continue
            cleaned[device_id] = dict(item)

        _LOCATION_POOL_CACHE = cleaned
        return {key: dict(value) for key, value in cleaned.items()}


def _write_dispatcher_location_pool(entries: dict[str, dict]) -> dict[str, dict]:
    global _LOCATION_POOL_CACHE, _LOCATION_POOL_MTIME_NS

    SHARED_DISPATCHER_LOCATION_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": 1,
        "updated_at_epoch": time.time(),
        "max_age_seconds": DISPATCHER_LOCATION_MAX_AGE_SECONDS,
        "devices": entries,
    }
    encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    temp_path = SHARED_DISPATCHER_LOCATION_PATH.with_name(
        f".{SHARED_DISPATCHER_LOCATION_PATH.name}.{uuid.uuid4().hex}.tmp"
    )
    try:
        temp_path.write_text(encoded, encoding="utf-8")
        os.replace(temp_path, SHARED_DISPATCHER_LOCATION_PATH)
    finally:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass

    try:
        _LOCATION_POOL_MTIME_NS = SHARED_DISPATCHER_LOCATION_PATH.stat().st_mtime_ns
    except OSError:
        _LOCATION_POOL_MTIME_NS = None
    _LOCATION_POOL_CACHE = {key: dict(value) for key, value in entries.items()}
    return {key: dict(value) for key, value in entries.items()}


def sync_dispatcher_location_pool(location: dict | None = None) -> list[dict]:
    """Share fresh dispatcher GPS fixes across active devices without UI controls."""
    device_id = _session_device_id()
    location = dict(location) if _location_payload_is_valid(location) else _newest_session_location()

    with _LOCATION_POOL_LOCK:
        entries = _read_dispatcher_location_pool(force=True)
        changed = False
        if _location_payload_is_valid(location):
            lat = float(location["latitude"])
            lon = float(location["longitude"])
            accuracy = max(0.0, _float_or_none(location.get("accuracy")) or 0.0)
            updated_at = _float_or_none(location.get("updated_at")) or time.time()
            previous = entries.get(device_id, {})
            incoming = {
                "device_id": device_id,
                "latitude": lat,
                "longitude": lon,
                "accuracy": accuracy,
                "updated_at": updated_at,
            }
            if previous != incoming:
                entries[device_id] = incoming
                changed = True

        if changed:
            entries = _write_dispatcher_location_pool(entries)

        result = [dict(item) for item in entries.values()]
        try:
            st.session_state["__ai_active_dispatcher_locations__"] = result
        except Exception:
            pass
        return result


def _coordinate_from_mapping(item: object) -> tuple[float, float] | None:
    if not isinstance(item, dict):
        return None

    lat_keys = ("latitude", "lat", "緯度", "場站緯度")
    lon_keys = ("longitude", "lng", "lon", "經度", "場站經度")
    lat = next((_float_or_none(item.get(key)) for key in lat_keys if item.get(key) is not None), None)
    lon = next((_float_or_none(item.get(key)) for key in lon_keys if item.get(key) is not None), None)
    if lat is None or lon is None or not (-90 <= lat <= 90 and -180 <= lon <= 180):
        return None
    return float(lat), float(lon)


def _station_locations_from_dataframe(frame: pd.DataFrame) -> dict[str, dict]:
    if not isinstance(frame, pd.DataFrame) or frame.empty or "場站名稱" not in frame.columns:
        return {}

    output: dict[str, dict] = {}
    for row in frame.to_dict(orient="records"):
        coordinates = _coordinate_from_mapping(row)
        if coordinates is None:
            continue
        name = str(row.get("場站名稱") or "").strip()
        if not name:
            continue
        output[name] = {"latitude": coordinates[0], "longitude": coordinates[1]}
    return output


def _normalize_station_locations(raw: object) -> dict[str, dict]:
    if not isinstance(raw, dict):
        return {}
    output: dict[str, dict] = {}
    for name, item in raw.items():
        coordinates = _coordinate_from_mapping(item)
        if coordinates is None:
            continue
        station_name = str(name or (item.get("station_name") if isinstance(item, dict) else "") or "").strip()
        if not station_name:
            continue
        output[station_name] = {"latitude": coordinates[0], "longitude": coordinates[1]}
    return output


def _runtime_station_locations(current_df: pd.DataFrame) -> dict[str, dict]:
    """Resolve the station-coordinate map already built by the legacy UI."""
    dataframe_locations = _station_locations_from_dataframe(current_df)
    if dataframe_locations:
        return dataframe_locations

    candidate_names = (
        "combined_station_locations",
        "prebuilt_station_locations",
        "station_locations",
    )
    frame = inspect.currentframe()
    try:
        frame = frame.f_back if frame else None
        for _ in range(8):
            if frame is None:
                break
            for namespace in (frame.f_locals, frame.f_globals):
                for name in candidate_names:
                    normalized = _normalize_station_locations(namespace.get(name))
                    if normalized:
                        return normalized
            frame = frame.f_back
    finally:
        del frame
    return {}


def _haversine_meters(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius = 6371000.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    value = (
        math.sin(dphi / 2.0) ** 2
        + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2.0) ** 2
    )
    return radius * 2.0 * math.atan2(math.sqrt(value), math.sqrt(max(0.0, 1.0 - value)))


def _location_station_matches(
    station_locations: dict[str, dict],
    dispatcher_locations: list[dict],
    observed_at: float,
) -> dict[str, dict]:
    """Map each active dispatcher to only their nearest station."""
    matches: dict[str, dict] = {}
    if not station_locations or not dispatcher_locations:
        return matches

    normalized_stations: list[tuple[str, str, float, float]] = []
    for station_name, location in station_locations.items():
        coordinates = _coordinate_from_mapping(location)
        if coordinates is None:
            continue
        normalized_stations.append(
            (station_name, _station_key(station_name), coordinates[0], coordinates[1])
        )

    for dispatcher in dispatcher_locations:
        if not _location_payload_is_valid(dispatcher):
            continue
        updated_at = _float_or_none(dispatcher.get("updated_at")) or 0.0
        age = max(0.0, observed_at - updated_at)
        if updated_at <= 0 or age > DISPATCHER_LOCATION_MAX_AGE_SECONDS:
            continue

        accuracy = max(0.0, _float_or_none(dispatcher.get("accuracy")) or 0.0)
        if accuracy > DISPATCHER_MAX_USABLE_ACCURACY_METERS:
            continue

        lat = float(dispatcher["latitude"])
        lon = float(dispatcher["longitude"])
        distances = sorted(
            (
                (_haversine_meters(lat, lon, station_lat, station_lon), station_name, station_key)
                for station_name, station_key, station_lat, station_lon in normalized_stations
            ),
            key=lambda item: item[0],
        )
        if not distances:
            continue

        nearest_distance, nearest_name, nearest_key = distances[0]
        second_distance = distances[1][0] if len(distances) > 1 else math.inf
        gap = second_distance - nearest_distance

        confirmed_radius = min(
            100.0,
            max(DISPATCHER_CONFIRMED_RADIUS_METERS, accuracy * 1.15 if accuracy else 0.0),
        )
        clear_nearest = gap >= DISPATCHER_AMBIGUITY_GAP_METERS or nearest_distance <= 35.0

        if nearest_distance <= confirmed_radius and clear_nearest:
            confidence = "confirmed"
        elif nearest_distance <= DISPATCHER_SUSPECTED_RADIUS_METERS:
            confidence = "suspected"
        else:
            continue

        candidate = {
            "station_name": nearest_name,
            "station_key": nearest_key,
            "confidence": confidence,
            "distance_m": round(nearest_distance, 1),
            "second_distance_m": None if not math.isfinite(second_distance) else round(second_distance, 1),
            "accuracy_m": round(accuracy, 1),
            "location_age_seconds": round(age, 1),
            "device_id": str(dispatcher.get("device_id") or ""),
        }

        previous = matches.get(nearest_key)
        if previous is None:
            matches[nearest_key] = candidate
        elif previous.get("confidence") != "confirmed" and confidence == "confirmed":
            matches[nearest_key] = candidate
        elif float(candidate["distance_m"]) < float(previous.get("distance_m") or math.inf):
            matches[nearest_key] = candidate

    return matches


def build_manual_intervention_event(
    *,
    station_name: str,
    bike_delta: int = 0,
    ebike_delta: int = 0,
    ai_context: dict | None = None,
    recorded_at_epoch: float | None = None,
) -> dict:
    context = dict(ai_context or {})
    timestamp = float(recorded_at_epoch or time.time())
    return {
        "event_id": uuid.uuid4().hex,
        "station_name": str(station_name or "").strip(),
        "station_key": _station_key(station_name),
        "bike_delta": int(bike_delta),
        "ebike_delta": int(ebike_delta),
        "recorded_at_epoch": timestamp,
        "operating_date": str(context.get("operating_date") or ""),
        "day_type": str(context.get("day_type") or ""),
        "shift": str(context.get("shift") or ""),
        "source_shift": str(context.get("source_shift") or ""),
        "consumed": False,
        "consumed_at_epoch": None,
    }


def _find_manual_event(
    station_key: str,
    manual_events: list[dict],
    observed_at_epoch: float,
) -> dict | None:
    candidates: list[dict] = []
    for event in manual_events:
        if not isinstance(event, dict) or bool(event.get("consumed")):
            continue
        if str(event.get("station_key") or _station_key(event.get("station_name"))) != station_key:
            continue
        recorded_at = _float_or_none(event.get("recorded_at_epoch")) or 0.0
        age = observed_at_epoch - recorded_at
        if -60 <= age <= MANUAL_EVENT_WINDOW_SECONDS:
            candidates.append(event)
    if not candidates:
        return None
    return max(candidates, key=lambda item: _float_or_none(item.get("recorded_at_epoch")) or 0.0)


def classify_live_transition(
    previous_df: pd.DataFrame,
    current_df: pd.DataFrame,
    *,
    manual_events: list[dict] | None = None,
    ai_context: dict | None = None,
    observed_at_epoch: float | None = None,
    source_event_id: str = "",
) -> dict:
    """Classify station changes using manual events + shared dispatcher GPS."""
    observed_at = float(observed_at_epoch or time.time())
    context = dict(ai_context or {})
    events = [dict(item) for item in (manual_events or []) if isinstance(item, dict)]
    previous = _records_by_station(previous_df)
    current = _records_by_station(current_df)
    records: list[dict] = []

    dispatcher_locations = sync_dispatcher_location_pool()
    station_locations = _runtime_station_locations(current_df)
    geo_matches = _location_station_matches(station_locations, dispatcher_locations, observed_at)

    try:
        st.session_state["__ai_location_guard_summary__"] = {
            "active_dispatchers": len(dispatcher_locations),
            "matched_stations": len(geo_matches),
            "confirmed": sum(1 for item in geo_matches.values() if item.get("confidence") == "confirmed"),
            "suspected": sum(1 for item in geo_matches.values() if item.get("confidence") == "suspected"),
        }
    except Exception:
        pass

    try:
        last_observed_map = st.session_state.setdefault("__ai_last_observed_by_station__", {})
        if not isinstance(last_observed_map, dict):
            last_observed_map = {}
            st.session_state["__ai_last_observed_by_station__"] = last_observed_map
    except Exception:
        last_observed_map = {}

    context_prefix = "|".join(
        (str(context.get("operating_date") or ""), str(context.get("shift") or ""))
    )

    for station_key, current_item in current.items():
        previous_item = previous.get(station_key)
        station_name = current_item["station_name"]
        current_bike = current_item["bike"]
        current_ebike = current_item["ebike"]

        previous_bike = previous_item.get("bike") if previous_item else None
        previous_ebike = previous_item.get("ebike") if previous_item else None

        timing_key = f"{context_prefix}|{station_key}"
        previous_observed_at = _float_or_none(last_observed_map.get(timing_key)) or 0.0
        elapsed_seconds: float | None = None
        if previous_observed_at > 0:
            elapsed = observed_at - previous_observed_at
            if 5.0 <= elapsed <= 4 * 60 * 60:
                elapsed_seconds = float(elapsed)
        last_observed_map[timing_key] = observed_at

        classification = "natural"
        review_status = "accepted"
        natural_weight = 1.0
        decision_weight = 0.0
        matched_manual_event_id = ""
        manual_detection_source = ""
        geo_match = None
        bike_delta: int | None = None
        ebike_delta: int | None = None

        if previous_item is None:
            classification = "baseline"
            review_status = "not_applicable"
            natural_weight = 0.0
        elif None in (previous_bike, previous_ebike, current_bike, current_ebike):
            classification = "incomplete"
            review_status = "not_applicable"
            natural_weight = 0.0
        else:
            bike_delta = int(current_bike) - int(previous_bike)
            ebike_delta = int(current_ebike) - int(previous_ebike)
            changed = bool(bike_delta or ebike_delta)

            manual_event = _find_manual_event(station_key, events, observed_at) if changed else None
            geo_match = geo_matches.get(station_key) if changed else None

            if manual_event is not None:
                classification = "manual_intervention"
                review_status = "confirmed"
                natural_weight = 0.0
                decision_weight = 1.0
                manual_detection_source = "manual_event"
                matched_manual_event_id = str(manual_event.get("event_id") or "")
                manual_event["consumed"] = True
                manual_event["consumed_at_epoch"] = observed_at
                manual_event["observed_bike_delta"] = bike_delta
                manual_event["observed_ebike_delta"] = ebike_delta
            elif isinstance(geo_match, dict) and geo_match.get("confidence") == "confirmed":
                classification = "manual_intervention"
                review_status = "confirmed"
                natural_weight = 0.0
                decision_weight = 1.0
                manual_detection_source = "geolocation"
            elif isinstance(geo_match, dict) and geo_match.get("confidence") == "suspected":
                classification = "suspected_intervention"
                review_status = "pending"
                natural_weight = 0.0
                manual_detection_source = "geolocation_nearby"
            else:
                max_single = max(abs(bike_delta), abs(ebike_delta))
                total_abs = abs(bike_delta) + abs(ebike_delta)
                if max_single >= SUSPECTED_SINGLE_TYPE_DELTA or total_abs >= SUSPECTED_TOTAL_ABS_DELTA:
                    classification = "suspected_intervention"
                    review_status = "pending"
                    natural_weight = 0.0
                    manual_detection_source = "large_delta_guard"

        records.append(
            {
                "record_id": uuid.uuid4().hex,
                "source_event_id": str(source_event_id or ""),
                "observed_at_epoch": observed_at,
                "elapsed_seconds": elapsed_seconds,
                "operating_date": str(context.get("operating_date") or ""),
                "day_type": str(context.get("day_type") or ""),
                "shift": str(context.get("shift") or ""),
                "source_shift": str(context.get("source_shift") or ""),
                "station_name": station_name,
                "station_key": station_key,
                "previous_bike": previous_bike,
                "current_bike": current_bike,
                "bike_delta": bike_delta,
                "previous_ebike": previous_ebike,
                "current_ebike": current_ebike,
                "ebike_delta": ebike_delta,
                "classification": classification,
                "natural_training_weight": natural_weight,
                "decision_training_weight": decision_weight,
                "review_status": review_status,
                "manual_event_id": matched_manual_event_id,
                "manual_detection_source": manual_detection_source,
                "dispatcher_device_id": str((geo_match or {}).get("device_id") or ""),
                "dispatcher_distance_m": (geo_match or {}).get("distance_m"),
                "dispatcher_accuracy_m": (geo_match or {}).get("accuracy_m"),
                "dispatcher_location_age_seconds": (geo_match or {}).get("location_age_seconds"),
            }
        )

    counts: dict[str, int] = {}
    for record in records:
        label = str(record.get("classification") or "unknown")
        counts[label] = counts.get(label, 0) + 1

    return {
        "records": records,
        "manual_events": events[-MAX_MANUAL_EVENTS:],
        "summary": counts,
        "observed_at_epoch": observed_at,
    }


def trim_learning_records(records: list[dict] | None) -> list[dict]:
    """Keep local history while silently pushing new samples to the shared pool."""
    global _LATEST_LEARNING_RECORDS, _PUSHED_RECORD_IDS

    local_trimmed = _sort_trim_records(records)
    new_for_shared: list[dict] = []
    for item in local_trimmed:
        identity = _record_identity(item)
        if identity in _PUSHED_RECORD_IDS:
            continue
        _PUSHED_RECORD_IDS.add(identity)
        new_for_shared.append(item)

    shared = sync_shared_learning_pool(new_for_shared)
    _LATEST_LEARNING_RECORDS = [dict(item) for item in shared]
    return local_trimmed


def _prediction_records() -> list[dict]:
    try:
        sync_dispatcher_location_pool()
    except Exception:
        pass

    shared = sync_shared_learning_pool()
    if shared:
        return shared
    try:
        session_records = st.session_state.get("__ai_prediction_records__", [])
    except Exception:
        session_records = []
    if isinstance(session_records, list) and session_records:
        return [item for item in session_records if isinstance(item, dict)]
    return [dict(item) for item in _LATEST_LEARNING_RECORDS]


def _direction(value: float) -> str:
    if value >= 0.35:
        return "↑"
    if value <= -0.35:
        return "↓"
    return "→"


def _bounded_hourly_rate(delta: float, elapsed_seconds: float) -> float:
    if elapsed_seconds <= 0:
        return 0.0
    rate = float(delta) * 3600.0 / float(elapsed_seconds)
    return max(-40.0, min(40.0, rate))


def build_early_prediction(
    station_name: str,
    *,
    records: list[dict] | None = None,
    ai_context: dict | None = None,
    now_epoch: float | None = None,
    current_bike: int | None = None,
    current_ebike: int | None = None,
) -> dict:
    """Predict the station state 60 minutes ahead from natural-demand history."""
    station_key = _station_key(station_name)
    context = dict(ai_context or {})
    now_ts = float(now_epoch or time.time())
    now_hour = time.localtime(now_ts).tm_hour
    usable: list[tuple[dict, float, float, float]] = []

    for record in (records if records is not None else _prediction_records()):
        if not isinstance(record, dict):
            continue
        if str(record.get("station_key") or _station_key(record.get("station_name"))) != station_key:
            continue
        if str(record.get("classification") or "") != "natural":
            continue
        natural_weight = _float_or_none(record.get("natural_training_weight", 1.0)) or 0.0
        if natural_weight <= 0:
            continue

        bike_delta = record.get("bike_delta")
        ebike_delta = record.get("ebike_delta")
        elapsed_seconds = _float_or_none(record.get("elapsed_seconds")) or 0.0
        if bike_delta is None or ebike_delta is None or not (5.0 <= elapsed_seconds <= 4 * 60 * 60):
            continue

        weight = natural_weight
        record_day_type = str(record.get("day_type") or "")
        record_shift = str(record.get("shift") or "")
        current_day_type = str(context.get("day_type") or "")
        current_shift = str(context.get("shift") or "")
        if current_day_type and record_day_type:
            weight *= 1.35 if current_day_type == record_day_type else 0.65
        if current_shift and record_shift:
            weight *= 1.5 if current_shift == record_shift else 0.6

        try:
            record_hour = time.localtime(float(record.get("observed_at_epoch") or 0)).tm_hour
            hour_distance = abs(record_hour - now_hour)
            hour_distance = min(hour_distance, 24 - hour_distance)
            if hour_distance <= 1:
                weight *= 1.6
            elif hour_distance <= 3:
                weight *= 1.2
            elif hour_distance >= 8:
                weight *= 0.7
        except (TypeError, ValueError, OSError):
            pass

        bike_hourly = _bounded_hourly_rate(float(bike_delta), elapsed_seconds)
        ebike_hourly = _bounded_hourly_rate(float(ebike_delta), elapsed_seconds)
        usable.append((record, weight, bike_hourly, ebike_hourly))

    if not usable:
        return {
            "ready": False,
            "samples": 0,
            "confidence": "學習中",
            "bike_direction": "→",
            "ebike_direction": "→",
            "label": "🔮 60分鐘預測：學習中",
        }

    total_weight = sum(weight for _, weight, _, _ in usable) or 1.0
    bike_60m_delta = sum(bike_rate * weight for _, weight, bike_rate, _ in usable) / total_weight
    ebike_60m_delta = sum(ebike_rate * weight for _, weight, _, ebike_rate in usable) / total_weight
    sample_count = len(usable)

    if sample_count < 5:
        confidence = "低"
    elif sample_count < 20:
        confidence = "中"
    else:
        confidence = "高"

    bike_direction = _direction(bike_60m_delta)
    ebike_direction = _direction(ebike_60m_delta)

    if current_bike is not None and current_ebike is not None:
        bike_future = max(0, int(round(float(current_bike) + bike_60m_delta)))
        ebike_future = max(0, int(round(float(current_ebike) + ebike_60m_delta)))
        label = (
            f"🔮 60分鐘預測：2.0 約{bike_future}台 {bike_direction}｜"
            f"2.0E 約{ebike_future}台 {ebike_direction}｜"
            f"信心{confidence}・{sample_count}筆"
        )
    else:
        label = (
            f"🔮 60分鐘預測：2.0 {bike_60m_delta:+.1f} {bike_direction}｜"
            f"2.0E {ebike_60m_delta:+.1f} {ebike_direction}｜"
            f"信心{confidence}・{sample_count}筆"
        )

    return {
        "ready": True,
        "samples": sample_count,
        "confidence": confidence,
        "horizon_minutes": PREDICTION_HORIZON_MINUTES,
        "bike_direction": bike_direction,
        "ebike_direction": ebike_direction,
        "bike_60m_delta": bike_60m_delta,
        "ebike_60m_delta": ebike_60m_delta,
        "label": label,
    }


def _replace_analysis_prediction_labels(body: str) -> str:
    if "analysis-ai-prediction" not in body or "AI 預測：學習中" not in body:
        return body

    try:
        context = st.session_state.get("ai_shift_context", {})
    except Exception:
        context = {}

    row_pattern = re.compile(
        r'(<tr[^>]*data-ubike-station-name="(?P<station>[^"]+)"[^>]*>.*?</tr>)',
        flags=re.DOTALL,
    )
    marker_pattern = re.compile(
        r'(<small class="analysis-ai-prediction"[^>]*>)(.*?)(</small>)',
        flags=re.DOTALL,
    )
    current_pattern = re.compile(r'<small>目前\s*([^／<]+)／標準', flags=re.DOTALL)

    def parse_current(value: str) -> int | None:
        try:
            return max(0, int(float(html_lib.unescape(value).strip())))
        except (TypeError, ValueError, OverflowError):
            return None

    def replace_row(match: re.Match) -> str:
        row_html = match.group(1)
        station_name = html_lib.unescape(match.group("station"))
        current_values = current_pattern.findall(row_html)
        current_bike = parse_current(current_values[0]) if len(current_values) >= 1 else None
        current_ebike = parse_current(current_values[1]) if len(current_values) >= 2 else None
        prediction = build_early_prediction(
            station_name,
            ai_context=context,
            current_bike=current_bike,
            current_ebike=current_ebike,
        )
        safe_label = html_lib.escape(
            str(prediction.get("label") or "🔮 60分鐘預測：學習中")
        )
        return marker_pattern.sub(
            lambda m: f"{m.group(1)}{safe_label}{m.group(3)}",
            row_html,
            count=1,
        )

    return row_pattern.sub(replace_row, body)


def _install_prediction_markdown_patch() -> None:
    """Upgrade the existing V29 prediction placeholder without rewriting legacy UI."""
    if getattr(st, "_ubike_ai_prediction_markdown_installed", False):
        return

    original_markdown = st.markdown

    def prediction_markdown(body, *args, **kwargs):
        if isinstance(body, str):
            body = _replace_analysis_prediction_labels(body)
        return original_markdown(body, *args, **kwargs)

    st.markdown = prediction_markdown
    st._ubike_ai_prediction_markdown_installed = True


# Warm shared pools at startup. Fail-open so AI extras can never block the main UI.
try:
    sync_shared_learning_pool(force_read=True)
except Exception:
    pass

try:
    sync_dispatcher_location_pool()
except Exception:
    pass

_install_prediction_markdown_patch()
