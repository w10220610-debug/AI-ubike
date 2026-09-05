from __future__ import annotations

"""V29 old-UI compatibility entrypoint.

The original legacy UI is kept byte-for-byte in ``legacy_ui.py``.  Before
executing it, this entrypoint applies only the cross-county compatibility
changes required by V29:

1. battery ranges accept any non-empty Excel zone instead of D1/D2/D3 only;
2. an uploaded workbook never gets merged with the built-in Taitung fallback;
3. Taitung fallback remains available only in the existing no-workbook path.
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

exec(compile(source, str(LEGACY_APP), "exec"), globals(), globals())
