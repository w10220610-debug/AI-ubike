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
   one-time refresh token when no browser component exists;
6. the floating battery query uses the V29 Server battery engine and a mobile-
   safe one-way HTML UI, avoiding custom-component readiness failures;
7. the V29 battery entry occupies the exact legacy battery-button slot so the
   new engine replaces the old entry instead of appearing as a second control.
"""

from pathlib import Path

from battery_upgrade import render_floating_server_battery as _render_floating_server_battery


LEGACY_APP = Path(__file__).with_name("legacy_ui.py")
source = LEGACY_APP.read_text(encoding="utf-8")


def render_floating_battery_query(
    route_station_map: dict[str, list[dict]],
    mobile_mode: bool,
) -> None:
    """Render V29 Server battery UI in the legacy floating-button position."""
    _render_floating_server_battery(route_station_map, mobile_mode)

    # battery_upgrade owns the data/UI engine. This tiny one-way HTML shim only
    # preserves the legacy physical slot and removes any hot-reload leftovers.
    # ``components`` is imported by legacy_ui.py before any call reaches here.
    components.html(
        r'''
        <script>
        (() => {
          const doc = window.parent.document;
          ['ubike-battery-fab', 'ubike-battery-page', 'ubike-battery-style'].forEach(id => {
            try { doc.getElementById(id)?.remove(); } catch (_) {}
          });

          let style = doc.getElementById('ubike-v29-legacy-slot-style');
          if (!style) {
            style = doc.createElement('style');
            style.id = 'ubike-v29-legacy-slot-style';
            doc.head.appendChild(style);
          }
          style.textContent = `
            #ub-v29-fab {
              right: 18px !important;
              bottom: 278px !important;
              width: 56px !important;
              height: 56px !important;
              border-radius: 16px !important;
              font: 850 13px/1.1 -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft JhengHei", sans-serif !important;
              color: #061116 !important;
              background: linear-gradient(135deg, #55f6ff, #f4ff57) !important;
              box-shadow: 0 0 24px rgba(85,246,255,.38), 0 8px 28px rgba(0,0,0,.34) !important;
            }
            @media (max-width: 700px) {
              #ub-v29-fab {
                right: 10px !important;
                bottom: calc(312px + env(safe-area-inset-bottom, 0px)) !important;
                width: 52px !important;
                height: 52px !important;
              }
            }
          `;

          const applyLabel = () => {
            const fab = doc.getElementById('ub-v29-fab');
            if (!fab) return;
            fab.textContent = '⚡電量';
            fab.title = '查詢 YouBike 2.0E 電量';
            fab.setAttribute('aria-label', '電量查詢');
          };
          applyLabel();
          window.setTimeout(applyLabel, 80);
          window.setTimeout(applyLabel, 300);
        })();
        </script>
        ''',
        height=0,
        scrolling=False,
    )


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

# Keep the legacy browser battery implementation in the source for rollback,
# but rename it so every existing call site resolves to the V29 wrapper above.
replace_exact(
    'def render_floating_battery_query(\n',
    'def render_floating_battery_query_legacy(\n',
    label="replace floating battery implementation",
)

# Replace the browser-only live-status component factory with a Python Server
# backed callable. Downstream legacy UI code still receives the same payload
# shape, so the existing matching, cache persistence and rendering remain intact.
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
