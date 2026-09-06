from __future__ import annotations

import html as html_lib
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


# Install before legacy_ui.py is exec()'d by app.py. This keeps Streamlit v1
# custom component registration inside a real imported Python module and avoids
# inspect.getmodule(caller_frame) returning None for the legacy exec context.
install_component_declare_compat()


# Conservative first-pass guardrails. Suspected intervention is excluded from
# natural-demand training until a later review/confirmation step exists.
SUSPECTED_SINGLE_TYPE_DELTA = 6
SUSPECTED_TOTAL_ABS_DELTA = 8
MANUAL_EVENT_WINDOW_SECONDS = 30 * 60
MAX_MANUAL_EVENTS = 500
MAX_TRANSITION_RECORDS = 50000
PREDICTION_HORIZON_MINUTES = 60

# All bases/devices on the same running Streamlit service contribute to this
# common pool. It is intentionally kept outside any base token so switching
# phone/tablet/computer does not restart AI learning. Streamlit local disk is
# still ephemeral across a full redeploy; a remote DB can replace this path
# later without changing the predictor contract.
SHARED_POOL_SYNC_SECONDS = 10.0
SHARED_LEARNING_POOL_PATH = (
    Path(__file__).resolve().parent / ".base_cache" / "ai_learning_pool.shared.json"
)

_LATEST_LEARNING_RECORDS: list[dict] = []
_SHARED_POOL_CACHE: list[dict] = []
_SHARED_POOL_LAST_CHECK_MONOTONIC = 0.0
_SHARED_POOL_MTIME_NS: int | None = None
_SHARED_POOL_LOCK = threading.RLock()
_PUSHED_RECORD_IDS: set[str] = set()


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
        if not isinstance(item, dict):
            continue
        deduped[_record_identity(item)] = dict(item)

    def sort_key(item: dict) -> tuple[float, str]:
        try:
            observed = float(item.get("observed_at_epoch") or 0.0)
        except (TypeError, ValueError, OverflowError):
            observed = 0.0
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
        "version": 1,
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
    """Silently merge local learning into the device-shared AI pool.

    This function is called automatically during learning writes and prediction
    rendering. There is no user-facing sync button: each normal Streamlit rerun
    performs a lightweight pool check, while new samples are pushed immediately.
    """
    global _LATEST_LEARNING_RECORDS

    with _SHARED_POOL_LOCK:
        existing = _read_shared_pool(force=force_read)
        incoming = [dict(item) for item in (incoming_records or []) if isinstance(item, dict)]
        if incoming:
            merged = _sort_trim_records([*existing, *incoming])
            existing_ids = {_record_identity(item) for item in existing}
            merged_ids = {_record_identity(item) for item in merged}
            if merged_ids != existing_ids:
                shared = _write_shared_pool(merged)
            else:
                shared = merged
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
        try:
            recorded_at = float(event.get("recorded_at_epoch") or 0)
        except (TypeError, ValueError):
            continue
        age = observed_at_epoch - recorded_at
        if -60 <= age <= MANUAL_EVENT_WINDOW_SECONDS:
            candidates.append(event)
    if not candidates:
        return None
    return max(candidates, key=lambda item: float(item.get("recorded_at_epoch") or 0))


def classify_live_transition(
    previous_df: pd.DataFrame,
    current_df: pd.DataFrame,
    *,
    manual_events: list[dict] | None = None,
    ai_context: dict | None = None,
    observed_at_epoch: float | None = None,
    source_event_id: str = "",
) -> dict:
    """Classify live count changes before they are allowed into AI training.

    natural: eligible for natural-demand learning.
    manual_intervention: excluded from natural learning, retained for later
      dispatcher-decision learning.
    suspected_intervention: conservatively excluded pending future review.
    baseline/incomplete: not a usable transition sample.

    elapsed_seconds is also recorded from the previous live observation so the
    predictor can normalize each change to a real 60-minute rate.
    """
    observed_at = float(observed_at_epoch or time.time())
    context = dict(ai_context or {})
    events = [dict(item) for item in (manual_events or []) if isinstance(item, dict)]
    previous = _records_by_station(previous_df)
    current = _records_by_station(current_df)
    records: list[dict] = []

    try:
        last_observed_map = st.session_state.setdefault("__ai_last_observed_by_station__", {})
        if not isinstance(last_observed_map, dict):
            last_observed_map = {}
            st.session_state["__ai_last_observed_by_station__"] = last_observed_map
    except Exception:
        last_observed_map = {}

    context_prefix = "|".join(
        (
            str(context.get("operating_date") or ""),
            str(context.get("shift") or ""),
        )
    )

    for station_key, current_item in current.items():
        previous_item = previous.get(station_key)
        station_name = current_item["station_name"]
        current_bike = current_item["bike"]
        current_ebike = current_item["ebike"]
        previous_bike = previous_item.get("bike") if previous_item else None
        previous_ebike = previous_item.get("ebike") if previous_item else None

        timing_key = f"{context_prefix}|{station_key}"
        try:
            previous_observed_at = float(last_observed_map.get(timing_key) or 0.0)
        except (TypeError, ValueError, OverflowError):
            previous_observed_at = 0.0
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
            if manual_event is not None:
                classification = "manual_intervention"
                review_status = "confirmed"
                natural_weight = 0.0
                decision_weight = 1.0
                matched_manual_event_id = str(manual_event.get("event_id") or "")
                manual_event["consumed"] = True
                manual_event["consumed_at_epoch"] = observed_at
                manual_event["observed_bike_delta"] = bike_delta
                manual_event["observed_ebike_delta"] = ebike_delta
            else:
                max_single = max(abs(bike_delta), abs(ebike_delta))
                total_abs = abs(bike_delta) + abs(ebike_delta)
                if max_single >= SUSPECTED_SINGLE_TYPE_DELTA or total_abs >= SUSPECTED_TOTAL_ABS_DELTA:
                    classification = "suspected_intervention"
                    review_status = "pending"
                    natural_weight = 0.0

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
    # Default background synchronization: prediction rendering checks the
    # shared file at most once per SHARED_POOL_SYNC_SECONDS, without any button.
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
    # One bad refresh must not dominate the forecast. Manual/suspected samples
    # are already excluded; this final cap protects against abnormal intervals.
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
        try:
            natural_weight = float(record.get("natural_training_weight", 1.0) or 0.0)
        except (TypeError, ValueError):
            natural_weight = 0.0
        if natural_weight <= 0:
            continue

        bike_delta = record.get("bike_delta")
        ebike_delta = record.get("ebike_delta")
        try:
            elapsed_seconds = float(record.get("elapsed_seconds") or 0.0)
        except (TypeError, ValueError, OverflowError):
            elapsed_seconds = 0.0
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
        safe_label = html_lib.escape(str(prediction.get("label") or "🔮 60分鐘預測：學習中"))
        return marker_pattern.sub(lambda m: f"{m.group(1)}{safe_label}{m.group(3)}", row_html, count=1)

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


# Load the shared pool immediately when the app starts. Subsequent reads are
# throttled and silent, so prediction learning stays synchronized by default.
try:
    sync_shared_learning_pool(force_read=True)
except Exception:
    pass

_install_prediction_markdown_patch()
