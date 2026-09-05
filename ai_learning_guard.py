from __future__ import annotations

import math
import re
import time
import unicodedata
import uuid
from typing import Any

import pandas as pd


# Conservative first-pass guardrails. Suspected intervention is excluded from
# natural-demand training until a later review/confirmation step exists.
SUSPECTED_SINGLE_TYPE_DELTA = 6
SUSPECTED_TOTAL_ABS_DELTA = 8
MANUAL_EVENT_WINDOW_SECONDS = 30 * 60
MAX_MANUAL_EVENTS = 500
MAX_TRANSITION_RECORDS = 5000


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
    return [dict(item) for item in (records or []) if isinstance(item, dict)][-MAX_TRANSITION_RECORDS:]
