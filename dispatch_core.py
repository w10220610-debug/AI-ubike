from __future__ import annotations

from io import BytesIO
from typing import Dict, List, Tuple

import pandas as pd


SHIFT_COLUMNS: Dict[str, Tuple[int, int]] = {
    "夜班配置": (4, 5),
    "早班配置": (7, 8),
    "晚班配置": (10, 11),
}
ROUTES = {"D1", "D2", "D3"}


def load_workbook(source: bytes | str) -> Dict[str, pd.DataFrame]:
    """讀取 Excel 全部工作表，保留原始列欄位置。"""
    excel_source = BytesIO(source) if isinstance(source, bytes) else source
    book = pd.ExcelFile(excel_source, engine="openpyxl")
    return {
        sheet_name: pd.read_excel(
            book,
            sheet_name=sheet_name,
            header=None,
            dtype=object,
            engine="openpyxl",
        )
        for sheet_name in book.sheet_names
    }


def text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def safe_int(value: object) -> int:
    number = pd.to_numeric(value, errors="coerce")
    return 0 if pd.isna(number) else int(number)


def route_header_rows(raw: pd.DataFrame) -> List[Tuple[int, str]]:
    """找出一張工作表裡 D1／D2／D3 區塊的標題列。"""
    found: List[Tuple[int, str]] = []
    for row_index in range(len(raw)):
        route = text(raw.iat[row_index, 0]).upper() if raw.shape[1] > 0 else ""
        station_header = text(raw.iat[row_index, 2]) if raw.shape[1] > 2 else ""
        if route in ROUTES and "場站名稱" in station_header:
            found.append((row_index, route))
    return found


def available_sources(sheets: Dict[str, pd.DataFrame]) -> List[Tuple[str, str]]:
    """回傳可選的（工作表、分區）組合，順序沿用 Excel。"""
    options: List[Tuple[str, str]] = []
    for sheet_name, raw in sheets.items():
        for _, route in route_header_rows(raw):
            options.append((sheet_name, route))
    return options


def parse_route(raw: pd.DataFrame, route: str, shift: str) -> pd.DataFrame:
    """擷取指定 D 區與班別的 2.0／2.0E 標準配置。"""
    headers = route_header_rows(raw)
    target_start = None
    target_end = len(raw)

    for index, (row_index, current_route) in enumerate(headers):
        if current_route == route:
            target_start = row_index + 1
            if index + 1 < len(headers):
                target_end = headers[index + 1][0]
            break

    if target_start is None:
        return pd.DataFrame()

    bike_col, ebike_col = SHIFT_COLUMNS[shift]
    rows = []

    for row_index in range(target_start, target_end):
        if raw.shape[1] <= max(2, bike_col, ebike_col):
            continue

        station_code = raw.iat[row_index, 0]
        region = text(raw.iat[row_index, 1])
        station = text(raw.iat[row_index, 2])
        bike = pd.to_numeric(raw.iat[row_index, bike_col], errors="coerce")
        ebike = pd.to_numeric(raw.iat[row_index, ebike_col], errors="coerce")

        if not station or pd.isna(station_code):
            continue
        if pd.isna(bike) and pd.isna(ebike):
            continue

        rows.append(
            {
                "行政區": region or "未分類",
                "場站名稱": station,
                "2.0 現況": safe_int(bike),
                "2.0E 現況": safe_int(ebike),
                "2.0 標準": safe_int(bike),
                "2.0E 標準": safe_int(ebike),
            }
        )

    return pd.DataFrame(rows)


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
