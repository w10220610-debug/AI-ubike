"""V29 road-routing performance regressions without importing the Streamlit app."""
from __future__ import annotations

import ast
import math
import types
from pathlib import Path

import numpy as np
import pandas as pd


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
NAMES = {
    "_road_router_failure_cooldown_active",
    "road_metrics_from_origin",
    "road_metrics_to_endpoint",
    "attach_analysis_road_metrics",
}


class RoutingError(RuntimeError):
    pass


def runtime():
    source = ast.parse(APP_PATH.read_text(encoding="utf-8"))
    selected = [
        node for node in source.body
        if isinstance(node, ast.FunctionDef) and node.name in NAMES
    ]
    assert {node.name for node in selected} == NAMES
    tree = ast.fix_missing_locations(ast.Module(
        body=[ast.ImportFrom(module="__future__", names=[ast.alias(name="annotations")], level=0), *selected],
        type_ignores=[],
    ))
    state: dict = {}
    clock = {"now": 100.0}
    calls = {"count": 0, "fail": True}
    pair_cache: dict = {}

    def key(lat, lon, dest_lat, dest_lon):
        return (round(float(lat), 4), round(float(lon), 4),
                round(float(dest_lat), 5), round(float(dest_lon), 5))

    def store(cache, *, origin_latitude, origin_longitude,
              destination_latitude, destination_longitude, metric):
        cache[key(origin_latitude, origin_longitude,
                  destination_latitude, destination_longitude)] = {
            **metric, "cached_at": clock["now"]
        }

    def fetch(coordinates, sources, destinations):
        calls["count"] += 1
        if calls["fail"]:
            raise RoutingError("OSRM unavailable")
        return {
            "durations": [[60 for _ in destinations] for _ in sources],
            "distances": [[1500 for _ in destinations] for _ in sources],
            "data_version": "test",
        }

    namespace = {
        "st": types.SimpleNamespace(session_state=state),
        "time": types.SimpleNamespace(time=lambda: clock["now"]),
        "ROAD_ROUTER_FAILURE_COOLDOWN_SECONDS": 45,
        "ROAD_ROUTER_STATUS_STATE_KEY": "road_status",
        "ROAD_PAIR_CACHE_STATE_KEY": "road_pairs",
        "ROAD_ROUTER_BATCH_SIZE": 2,
        "ROAD_ROUTER_ORIGIN_PRECISION": 4,
        "ROAD_ROUTER_STATION_PRECISION": 5,
        "RoadRoutingError": RoutingError,
        "_road_pair_cache": lambda: pair_cache,
        "_road_pair_cache_key": key,
        "_rounded_coordinate": lambda lon, lat, *, precision: (
            round(float(lon), precision), round(float(lat), precision)),
        "_cache_road_pair": store,
        "haversine_distance_km": lambda *args: 1.0,
        "fetch_road_table_cached": fetch,
        "normalize_coordinate": lambda v, _low, _high: float(v) if v is not None else None,
        "location_payload_is_valid": lambda item: isinstance(item, dict) and
            "latitude" in item and "longitude" in item,
        "pd": pd,
        "np": np,
    }
    exec(compile(tree, str(APP_PATH), "exec"), namespace)
    return namespace, state, clock, calls, pair_cache


def test_cooldown_first_failure_skips_later_batches_and_reruns():
    env, state, clock, calls, cache = runtime()
    destinations = [(f"station-{i}", 22.0 + i * 0.001, 121.0) for i in range(5)]
    request = dict(origin_latitude=22.9, origin_longitude=121.1, destinations=destinations)
    first = env["road_metrics_from_origin"](**request)
    assert calls["count"] == 1
    assert len(first) == 5
    assert all(m["road_route_available"] is False for m in first.values())
    assert state["road_status"]["updated_at"] == 100
    again = env["road_metrics_from_origin"](**request)
    assert calls["count"] == 1
    assert again.keys() == first.keys()
    assert state["road_status"]["updated_at"] == 100  # 冷卻不能被無限延長

    clock["now"] = 146
    calls["fail"] = False
    recovered = env["road_metrics_from_origin"](**request)
    assert calls["count"] == 4  # 三批各查一次
    assert all(m["road_route_available"] is True for m in recovered.values())
    assert state["road_status"]["ok"] is True
    assert len(cache) == 5
    env["road_metrics_from_origin"](**request)
    assert calls["count"] == 4  # 成功的道路段沿用快取


def test_cached_good_route_remains_usable_while_service_is_down():
    env, state, clock, calls, cache = runtime()
    state["road_status"] = {"ok": False, "updated_at": 100, "message": "offline"}
    cache[(22.9, 121.1, 22.0, 121.0)] = {
        "road_route_available": True,
        "road_distance_km": 1.5,
        "drive_minutes": 2.0,
        "cached_at": 100,
    }
    result = env["road_metrics_from_origin"](
        origin_latitude=22.9, origin_longitude=121.1,
        destinations=[("cached", 22.0, 121.0), ("uncached", 22.1, 121.0)],
    )
    assert calls["count"] == 0
    assert result["cached"]["road_route_available"] is True
    assert result["uncached"]["routing_service_error"] is True
    assert state["road_status"]["updated_at"] == 100


def test_endpoint_routes_share_failure_cooldown():
    env, state, clock, calls, cache = runtime()
    candidates = [
        {"station_name": f"S{i}", "latitude": 22.0 + i * 0.01, "longitude": 121.0}
        for i in range(5)
    ]
    result = env["road_metrics_to_endpoint"](
        candidates, {"latitude": 22.9, "longitude": 121.1}
    )
    assert calls["count"] == 1
    assert len(result) == 5
    assert all(row["routing_service_error"] is True for row in result.values())
    assert state["road_status"]["updated_at"] == 100


def test_cooldown_rejects_future_corrupt_and_expired_timestamps():
    env, state, clock, calls, cache = runtime()
    active = env["_road_router_failure_cooldown_active"]
    state["road_status"] = {"ok": False, "updated_at": "bad"}
    assert active() is False
    state["road_status"] = {"ok": False, "updated_at": 200}
    assert active() is False
    state["road_status"] = {"ok": False, "updated_at": 55}
    assert active() is False  # 恰好 45 秒可重試
    state["road_status"] = {"ok": True, "updated_at": 100}
    assert active() is False
    state["road_status"] = {"ok": False, "updated_at": 99}
    assert active() is True


def test_analysis_vectorization_preserves_duplicate_station_names_and_missing_routes():
    env, state, clock, calls, cache = runtime()
    env["road_metrics_from_origin"] = lambda **kwargs: {
        "A": {"road_route_available": True, "road_distance_km": 2.0, "drive_minutes": 5.0},
        "B": {"road_route_available": False},
        "C": {"road_route_available": True, "road_distance_km": 3.0, "drive_minutes": 6.0},
    }
    source = pd.DataFrame({"場站名稱": ["A", "B", "A", "C"]}, index=[3, 5, 8, 11])
    places = {
        name: {"latitude": 22.0, "longitude": 121.0} for name in ["A", "B", "C"]
    }
    result, available, message = env["attach_analysis_road_metrics"](
        source, station_locations=places,
        current_location={"latitude": 22.9, "longitude": 121.1}
    )
    assert available is True
    assert "3 個場站" in message
    assert result.index.tolist() == [3, 5, 8, 11]
    assert result["距離目前位置 (km)"].tolist()[0] == 2.0
    assert math.isnan(result["距離目前位置 (km)"].tolist()[1])
    assert result["距離目前位置 (km)"].tolist()[2] == 2.0
    assert result["預估行車時間 (分)"].tolist()[3] == 6.0
    assert "距離目前位置 (km)" not in source.columns


def test_analysis_without_gps_does_not_request_routes():
    env, state, clock, calls, cache = runtime()
    source = pd.DataFrame({"場站名稱": ["A"]})
    result, available, _ = env["attach_analysis_road_metrics"](
        source, station_locations={}, current_location=None
    )
    assert not available
    assert result["距離目前位置 (km)"].isna().all()
    assert calls["count"] == 0
