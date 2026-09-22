from __future__ import annotations

import ast
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import Mock, patch

import pytest

from performance_cache import CatalogMapCache, compile_legacy_source


def route(name="A", district="X"):
    return {"D1": [{"name": name, "district": district}]}


def test_compile_reuses_only_identical_code_and_filename():
    compile_legacy_source.cache_clear()
    code = compile_legacy_source("value = []", "legacy.py")
    assert code is compile_legacy_source("value = []", "legacy.py")
    assert code is not compile_legacy_source("value = [1]", "legacy.py")
    assert code is not compile_legacy_source("value = []", "other.py")
    assert compile_legacy_source.cache_info().currsize <= 2


def test_compiled_code_does_not_share_user_state():
    code = compile_legacy_source("value = []", "legacy.py")
    a, b = {}, {}
    exec(code, a)
    exec(code, b)
    a["value"].append("user A")
    assert b["value"] == []


def test_compile_errors_are_not_cached():
    for _ in range(2):
        with pytest.raises(SyntaxError):
            compile_legacy_source("if :", "bad.py")


def test_map_cache_reuses_and_copies_result():
    cache = CatalogMapCache()
    catalog = ({"id": 1},)
    build = Mock(side_effect=lambda clean, cat: clean)
    first = cache.resolve(route(), catalog, build)
    first["D1"][0]["name"] = "changed"
    assert cache.resolve(route(), catalog, build) == route()
    assert build.call_count == 1


def test_map_cache_invalidates_catalog_and_configuration():
    cache = CatalogMapCache()
    catalog = ({"id": 1},)
    build = Mock(side_effect=lambda clean, cat: clean)
    cache.resolve(route(), catalog, build)
    cache.resolve(route("B"), catalog, build)
    cache.resolve(route(district="Y"), catalog, build)
    cache.resolve(route(), ({"id": 2},), build)
    assert build.call_count == 4


def test_map_cache_bound_and_failed_build_retry():
    cache = CatalogMapCache(max_entries=2)
    catalog = ()
    build = Mock(side_effect=lambda clean, cat: clean)
    for name in ("A", "B", "C", "A"):
        cache.resolve(route(name), catalog, build)
    assert len(cache._entries) == 2
    assert build.call_count == 4
    failing = Mock(side_effect=ValueError("retry"))
    with pytest.raises(ValueError):
        cache.resolve(route("D"), catalog, failing)
    assert cache.resolve(route("D"), catalog, build) == route("D")


def test_map_cache_same_request_single_build():
    cache = CatalogMapCache()
    catalog = ({"id": 1},)
    build = Mock(side_effect=lambda clean, cat: clean)
    with ThreadPoolExecutor(max_workers=8) as pool:
        results = list(pool.map(lambda _: cache.resolve(route(), catalog, build), range(30)))
    assert all(item == route() for item in results)
    assert build.call_count == 1


def prepared_app():
    """Apply every real compatibility patch, without executing the UI/network."""
    path = Path(__file__).resolve().parents[1] / "app.py"
    tree = ast.parse(path.read_text(encoding="utf-8"))
    assert isinstance(tree.body[-1], ast.Expr)
    assert tree.body[-1].value.func.id == "exec"
    tree.body.pop()
    namespace = {"__file__": str(path), "__name__": "performance_test_app"}
    exec(compile(tree, str(path), "exec"), namespace)
    return namespace


def test_all_compatibility_patches_and_final_compile():
    ns = prepared_app()
    code = compile_legacy_source(ns["source"], str(ns["LEGACY_APP"]))
    assert code is compile_legacy_source(ns["source"], str(ns["LEGACY_APP"]))
    assert "效能優化（保留現有功能）" in ns["source"]
    assert 'with st.expander("🆕 更新內容"' in ns["source"]
    assert "v29_server_live_force_refresh" in ns["source"]
    assert "render_priority_station_manager(" in ns["source"]


def test_battery_resolver_output_cache_and_catalog_error():
    import battery_upgrade as battery
    import station_service as station
    catalog = ({"station_no": "1", "station_name": "A", "station_key": "a",
                "district": "X", "latitude": 23, "longitude": 121},)
    with patch.object(battery, "_resolved_station_maps", CatalogMapCache()), \
         patch.object(battery, "get_station_catalog", return_value=catalog), \
         patch.object(battery, "match_station", wraps=station.match_station) as match:
        expected = battery._build_resolved_station_map(route(), catalog)
        match.reset_mock()
        assert battery._resolve_station_numbers(route()) == (expected, "")
        assert battery._resolve_station_numbers(route()) == (expected, "")
        assert match.call_count == 1
        with patch.object(battery, "get_station_catalog", side_effect=station.StationServiceError("offline")):
            result, error = battery._resolve_station_numbers(route())
            assert error == "offline"
            assert result["D1"][0]["station_no"] == ""
        assert battery._resolve_station_numbers(route()) == (expected, "")


def test_unconfigured_app_starts_without_network():
    from streamlit.testing.v1 import AppTest
    path = Path(__file__).resolve().parents[1] / "app.py"
    with patch("battery_upgrade.get_station_catalog", return_value=()), \
         patch("urllib.request.urlopen", side_effect=AssertionError("unexpected network")):
        app = AppTest.from_file(str(path)).run(timeout=20)
        assert not app.exception
        assert any("請上傳配置表" in item.value for item in app.info)


def test_configured_app_rerun_keeps_data_and_mode(tmp_path, monkeypatch):
    import hashlib
    from io import BytesIO
    import pandas as pd
    from streamlit.testing.v1 import AppTest
    import battery_upgrade as battery
    import ai_learning_guard
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(ai_learning_guard, "SHARED_LEARNING_POOL_PATH", tmp_path / "learning.json")
    monkeypatch.setattr(ai_learning_guard, "SHARED_DISPATCHER_LOCATION_PATH", tmp_path / "locations.json")
    path = Path(__file__).resolve().parents[1] / "app.py"
    rows = []
    for zone in ("D1", "D2", "D3"):
        rows.extend([[zone, "行政區", "場站名稱", "", "夜班配置", "2.0E", "2.0綁車", "早班配置", "2.0E", "2.0綁車", "晚班配置", "2.0E"],
                     [1, "台東市", zone + "測試站", "", 4, 2, 1, 4, 2, 1, 4, 2]])
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name="平日", header=False, index=False)
    data = buffer.getvalue()
    token = "a" * 32
    payload = {"ok": True, "event_id": "fixture-event", "records": [], "fetched_at": "fixture"}
    with patch.object(battery, "get_station_catalog", return_value=()), \
         patch("live_status_service.get_live_status_for_stations", return_value=payload), \
         patch("urllib.request.urlopen", side_effect=AssertionError("unexpected network")):
        app = AppTest.from_file(str(path))
        app.query_params["base"] = token
        app.session_state[f"disk_base_cache::{token}"] = {
            "token": token, "name": "fixture.xlsx", "bytes": data,
            "sha256": hashlib.sha256(data).hexdigest(), "expires_at": None,
        }
        app.run(timeout=20)
        assert not app.exception
        work_mode = next(item for item in app.radio if item.label == "工作模式")
        work_mode.set_value("智慧調度").run(timeout=20)
        assert not app.exception
        assert next(item.value for item in app.radio if item.label == "工作模式") == "智慧調度"
        app.run(timeout=20)
        assert not app.exception
        assert app.session_state[f"disk_base_cache::{token}"]["bytes"] == data
