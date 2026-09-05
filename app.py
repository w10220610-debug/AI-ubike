from __future__ import annotations

"""V29 old-UI compatibility entrypoint.

The original legacy UI is kept byte-for-byte in ``legacy_ui.py``. Before
executing it, this entrypoint applies focused V29 compatibility fixes:

1. battery ranges accept any non-empty Excel zone instead of D1/D2/D3 only;
2. an uploaded workbook never gets merged with the built-in Taitung fallback;
3. Taitung fallback remains available only in the existing no-workbook path;
4. live station sync uses the V29 Python Server service instead of requiring a
   hidden browser Streamlit component on mobile;
5. the floating refresh button requests a fresh server sync by reloading with a
   one-time refresh token when no browser component exists.
"""

from pathlib import Path


LEGACY_APP = Path(__file__).with_name("legacy_ui.py")
source = LEGACY_APP.read_text(encoding="utf-8")


def replace_exact(old: str, new: str, *, label: str) -> None:
    global source
    count = source.count(old)
    if count != 1:
        raise RuntimeError(
            f"V29 compatibility patch failed for {label}: expected 1 match, got {count}."
        )
    source = source.replace(old, new, 1)


replace_exact(
    '"""由目前選定配置整理 D1／D2／D3 場站，供 2.0E 電量查詢使用。"""',
    '"""由目前選定配置整理任意區域場站，供 2.0E 電量查詢使用。"""',
    label="battery map docstring",
)
replace_exact(
    'route = "" if pd.isna(route_raw) else str(route_raw).strip().upper()',
    'route = "" if pd.isna(route_raw) else str(route_raw).strip()',
    label="dynamic zone label",
)
replace_exact(
    'if route not in ("D1", "D2", "D3") or not station_key:',
    'if not route or not station_key:',
    label="remove D1-D3 restriction",
)
replace_exact(
    'st.warning("D1、D2、D3 沒有可用場站。")',
    'st.warning("配置表沒有可用場站。")',
    label="empty-zone message",
)
replace_exact(
    '''battery_route_map = merge_battery_route_station_maps(\n    build_battery_route_station_map(base_df),\n    DEFAULT_BATTERY_ROUTE_STATION_MAP,\n)''',
    '''battery_route_map = build_battery_route_station_map(base_df)''',
    label="uploaded-workbook battery range",
)

# Replace the browser-only component factory with a Python Server backed
# callable. Downstream legacy UI code still receives the same payload shape, so
# the existing matching, cache persistence and UI rendering do not need to be
# rewritten.
replace_exact(
    'def normalize_browser_live_payload(payload) -> dict:',
    '''def get_youbike_browser_sync_component():\n    """V29 compatibility: obtain live station data from the Python Server."""\n    from live_status_service import LiveStatusServiceError, get_live_status_for_stations\n\n    def _server_sync_component(**_kwargs):\n        stations = globals().get("_V29_SERVER_LIVE_STATIONS", [])\n        if not stations:\n            return {\n                "ok": False,\n                "event_id": uuid.uuid4().hex,\n                "error": "目前配置沒有可供同步的場站。",\n            }\n\n        refresh_token = ""\n        try:\n            refresh_token = str(st.query_params.get("live_refresh", "") or "").strip()\n        except Exception:\n            refresh_token = ""\n        refresh_state_key = "v29_server_live_refresh_token"\n        force_refresh = bool(\n            refresh_token\n            and st.session_state.get(refresh_state_key) != refresh_token\n        )\n        if force_refresh:\n            st.session_state[refresh_state_key] = refresh_token\n\n        try:\n            return get_live_status_for_stations(stations, force=force_refresh)\n        except LiveStatusServiceError as exc:\n            return {\n                "ok": False,\n                "event_id": uuid.uuid4().hex,\n                "error": str(exc),\n            }\n        except Exception as exc:\n            return {\n                "ok": False,\n                "event_id": uuid.uuid4().hex,\n                "error": f"Server 即時車數同步失敗：{exc}",\n            }\n\n    return _server_sync_component\n\n\ndef normalize_browser_live_payload(payload) -> dict:''',
    label="server live component adapter",
)

# Build the exact station scope from the currently selected/uploaded workbook.
# This also makes the server sync automatically follow cross-county Excel data.
replace_exact(
    '    browser_payload = None',
    '''    _V29_SERVER_LIVE_STATIONS = [\n        {\n            "name": str(row.get("場站名稱") or "").strip(),\n            "district": str(row.get("行政區") or "").strip(),\n        }\n        for _, row in base_df.iterrows()\n        if str(row.get("場站名稱") or "").strip()\n    ]\n    browser_payload = None''',
    label="server live station scope",
)

# The legacy floating refresh button used to require discovery of a hidden
# Streamlit iframe. With the V29 server adapter there is intentionally no iframe.
# A refresh token causes one forced server fetch on the next Streamlit run.
replace_exact(
    '''                if (!postedCount) {{\n                    showToast("同步元件尚未準備完成，請稍後再按一次");\n                    return;\n                }}\n                setRefreshButtonState(true);''',
    '''                if (!postedCount) {{\n                    setRefreshButtonState(true);\n                    showToast("正在重新同步 YouBike 即時資料…");\n                    try {{\n                        const refreshUrl = new URL(win.location.href);\n                        refreshUrl.searchParams.set("live_refresh", String(Date.now()));\n                        win.location.href = refreshUrl.toString();\n                    }} catch (_refreshError) {{\n                        win.location.reload();\n                    }}\n                    return;\n                }}\n                setRefreshButtonState(true);''',
    label="server live floating refresh",
)

exec(compile(source, str(LEGACY_APP), "exec"), globals(), globals())
