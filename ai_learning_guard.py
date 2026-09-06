from __future__ import annotations

import html as html_lib
import math
import re
import time
import unicodedata
import uuid
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

# Early prediction is deliberately conservative. It only uses records already
# accepted as natural demand; manual/suspected intervention never contributes.
# The prediction is a direction/trend signal rather than an invented exact
# future vehicle count while the sample size is still small.
_LATEST_LEARNING_RECORDS: list[dict] = []


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
    """
    observed_at = float(observed_at_epoch or time.time())
    context = dict(ai_context or {})
    events = [dict(item) for item in (manual_events or []) if isinstance(item, dict)]
    previous = _records_by_station(previous_df)
    current = _records_by_station(current_df)
    records: list[dict] = []

    for station_key, current_item in current.items():
        previous_item = previous.get(station_key)
        station_name = current_item["station_name"]
        current_bike = current_item["bike"]
        current_ebike = current_item["ebike"]
        previous_bike = previous_item.get("bike") if previous_item else None
        previous_ebike = previous_item.get("ebike") if previous_item else None

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
    """Keep the newest learning records and refresh the in-process predictor."""
    global _LATEST_LEARNING_RECORDS
    trimmed = [dict(item) for item in (records or []) if isinstance(item, dict)][-MAX_TRANSITION_RECORDS:]
    _LATEST_LEARNING_RECORDS = trimmed
    try:
        st.session_state["__ai_prediction_records__"] = trimmed
    except Exception:
        pass
    return trimmed


def _prediction_records() -> list[dict]:
    try:
        session_records = st.session_state.get("__ai_prediction_records__", [])
    except Exception:
        session_records = []
    if isinstance(session_records, list) and session_records:
        return [item for item in session_records if isinstance(item, dict)]
    return _LATEST_LEARNING_RECORDS


def _direction(value: float) -> str:
    if value >= 0.35:
        return "↑"
    if value <= -0.35:
        return "↓"
    return "→"


def build_early_prediction(
    station_name: str,
    *,
    records: list[dict] | None = None,
    ai_context: dict | None = None,
    now_epoch: float | None = None,
) -> dict:
    """Return a conservative early trend prediction for one station.

    Records are weighted by matching day type/shift and nearby hour. This is an
    early statistical predictor, not a trained ML model, so it exposes sample
    count and confidence instead of pretending to know an exact future count.
    """
    station_key = _station_key(station_name)
    context = dict(ai_context or {})
    now_ts = float(now_epoch or time.time())
    now_hour = time.localtime(now_ts).tm_hour
    usable: list[tuple[dict, float]] = []

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
        if bike_delta is None or ebike_delta is None:
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

        usable.append((record, weight))

    if not usable:
        return {
            "ready": False,
            "samples": 0,
            "confidence": "學習中",
            "bike_direction": "→",
            "ebike_direction": "→",
            "label": "🔮 AI 預測：學習中",
        }

    total_weight = sum(weight for _, weight in usable) or 1.0
    bike_mean = sum(float(record.get("bike_delta") or 0) * weight for record, weight in usable) / total_weight
    ebike_mean = sum(float(record.get("ebike_delta") or 0) * weight for record, weight in usable) / total_weight
    sample_count = len(usable)

    if sample_count < 5:
        confidence = "低"
    elif sample_count < 20:
        confidence = "中"
    else:
        confidence = "高"

    bike_direction = _direction(bike_mean)
    ebike_direction = _direction(ebike_mean)
    label = (
        f"🔮 AI 早期預測：2.0 {bike_direction}｜2.0E {ebike_direction}｜"
        f"信心{confidence}・{sample_count}筆"
    )
    return {
        "ready": True,
        "samples": sample_count,
        "confidence": confidence,
        "bike_direction": bike_direction,
        "ebike_direction": ebike_direction,
        "bike_mean_delta": bike_mean,
        "ebike_mean_delta": ebike_mean,
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

    def replace_row(match: re.Match) -> str:
        row_html = match.group(1)
        station_name = html_lib.unescape(match.group("station"))
        prediction = build_early_prediction(station_name, ai_context=context)
        safe_label = html_lib.escape(str(prediction.get("label") or "🔮 AI 預測：學習中"))
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


_install_prediction_markdown_patch()
