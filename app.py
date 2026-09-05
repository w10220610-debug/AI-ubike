from __future__ import annotations

"""V29 old-UI compatibility entrypoint.

The original legacy UI is kept byte-for-byte in ``legacy_ui.py``. Before
executing it, this entrypoint applies focused V29 compatibility fixes:

1. battery ranges accept any non-empty Excel zone instead of D1/D2/D3 only;
2. an uploaded workbook never gets merged with the built-in Taitung fallback;
3. Taitung fallback remains available only in the existing no-workbook path;
4. mobile manual refresh can reach the hidden YouBike sync component reliably.
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

# Mobile browsers can fail to expose the hidden Streamlit component iframe with
# the title/src used by the floating refresh button. Publish a direct bridge on
# the parent window as soon as the component is ready; postMessage remains the
# fallback for browsers where direct parent access is unavailable.
replace_exact(
    '''  button.addEventListener("click", () => runSync({ forceDelivery: true }));\n  window.addEventListener("message", event => {''',
    '''  button.addEventListener("click", () => runSync({ forceDelivery: true }));\n  try {\n    window.parent.__ubikeManualSync = () => runSync({ forceDelivery: true });\n    window.parent.__ubikeSyncReady = true;\n  } catch (_bridgeError) {\n    // Direct parent bridge unavailable; postMessage fallback remains active.\n  }\n  window.addEventListener("message", event => {''',
    label="mobile sync direct bridge",
)

# Prefer the direct bridge. If the component has not mounted yet, retry briefly
# instead of immediately showing the old "sync component not ready" dead end.
replace_exact(
    '''            function requestManualSync() {{\n                let postedCount = 0;''',
    '''            function requestManualSync() {{\n                if (typeof win.__ubikeManualSync === "function") {{\n                    win.__ubikeManualSyncRetryCount = 0;\n                    setRefreshButtonState(true);\n                    showToast("正在手動更新 YouBike 即時車數");\n                    try {{\n                        win.__ubikeManualSync();\n                    }} catch (_bridgeError) {{\n                        setRefreshButtonState(false);\n                        showToast("同步連線暫時中斷，正在重新連線");\n                    }}\n                    return;\n                }}\n                let postedCount = 0;''',
    label="mobile sync prefer direct bridge",
)
replace_exact(
    '''                if (!postedCount) {{\n                    showToast("同步元件尚未準備完成，請稍後再按一次");\n                    return;\n                }}\n                setRefreshButtonState(true);''',
    '''                if (!postedCount) {{\n                    const retryCount = Number(win.__ubikeManualSyncRetryCount || 0);\n                    if (retryCount < 6) {{\n                        win.__ubikeManualSyncRetryCount = retryCount + 1;\n                        setRefreshButtonState(true);\n                        showToast("正在連接 YouBike 同步元件…");\n                        win.setTimeout(() => {{\n                            setRefreshButtonState(false);\n                            requestManualSync();\n                        }}, 500);\n                        return;\n                    }}\n                    win.__ubikeManualSyncRetryCount = 0;\n                    setRefreshButtonState(false);\n                    showToast("同步元件無法啟動，請重新整理頁面後再試");\n                    return;\n                }}\n                win.__ubikeManualSyncRetryCount = 0;\n                setRefreshButtonState(true);''',
    label="mobile sync readiness retry",
)

exec(compile(source, str(LEGACY_APP), "exec"), globals(), globals())
