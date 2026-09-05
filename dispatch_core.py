from __future__ import annotations

"""調度核心（相容版）。

Excel/區域解析已搬到 excel_service；這裡只保留舊 app.py 目前會 import 的名稱，
以及純計算函式，降低一次重構造成既有 UI 壞掉的風險。
"""

from typing import Tuple

import pandas as pd

from excel_service import (
    SHIFT_COLUMNS,
    available_sources,
    build_workbook_profile,
    build_zone_station_map,
    discover_zones,
    load_workbook,
    parse_route,
    route_header_rows,
)


def safe_int(value: object) -> int:
    number = pd.to_numeric(value, errors="coerce")
    return 0 if pd.isna(number) else int(number)


def diff_label(current: object, standard: object) -> str:
    diff = safe_int(current) - safe_int(standard)
    if diff > 0:
        return f"多 {diff} 台"
    if diff < 0:
        return f"缺 {abs(diff)} 台"
    return "符合"


def build_result(edited: pd.DataFrame) -> pd.DataFrame:
    result = edited[["行政區", "場站名稱"]].copy()
    result["2.0 缺／多幾台"] = [
        diff_label(current, standard)
        for current, standard in zip(edited["2.0 現況"], edited["2.0 標準"])
    ]
    result["2.0E 缺／多幾台"] = [
        diff_label(current, standard)
        for current, standard in zip(edited["2.0E 現況"], edited["2.0E 標準"])
    ]
    return result


def calculate_totals(edited: pd.DataFrame, current_col: str, standard_col: str) -> Tuple[int, int]:
    diff = (
        pd.to_numeric(edited[current_col], errors="coerce").fillna(0)
        - pd.to_numeric(edited[standard_col], errors="coerce").fillna(0)
    )
    shortage = int((-diff[diff < 0]).sum())
    surplus = int(diff[diff > 0].sum())
    return shortage, surplus


__all__ = [
    "SHIFT_COLUMNS",
    "available_sources",
    "build_workbook_profile",
    "build_zone_station_map",
    "discover_zones",
    "load_workbook",
    "parse_route",
    "route_header_rows",
    "safe_int",
    "diff_label",
    "build_result",
    "calculate_totals",
]
