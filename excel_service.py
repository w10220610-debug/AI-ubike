from __future__ import annotations

import re
import unicodedata
from collections import OrderedDict
from dataclasses import dataclass
from io import BytesIO
from typing import Any, Dict, Iterable, List, Tuple

import pandas as pd


SHIFT_COLUMNS: Dict[str, Tuple[int, int]] = {
    "夜班配置": (4, 5),
    "早班配置": (7, 8),
    "晚班配置": (10, 11),
}


@dataclass(frozen=True)
class ZoneHeader:
    row_index: int
    zone: str
    station_col: int
    region_col: int
    code_col: int


def text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def safe_int(value: object) -> int:
    number = pd.to_numeric(value, errors="coerce")
    return 0 if pd.isna(number) else int(number)


def normalize_zone(value: Any) -> str:
    value = unicodedata.normalize("NFKC", str(value or "")).strip()
    return re.sub(r"\s+", "", value).upper()


def _source(source: bytes | str):
    return BytesIO(source) if isinstance(source, bytes) else source


def load_workbook(source: bytes | str, *, visible_only: bool = True) -> Dict[str, pd.DataFrame]:
    """讀取 Excel；預設只讀「顯示中」工作表。"""
    with pd.ExcelFile(_source(source), engine="openpyxl") as book:
        workbook = book.book
        names = []
        for sheet_name in book.sheet_names:
            if not visible_only:
                names.append(sheet_name)
                continue
            worksheet = workbook[sheet_name]
            if getattr(worksheet, "sheet_state", "visible") == "visible":
                names.append(sheet_name)
        return {
            sheet_name: pd.read_excel(
                book,
                sheet_name=sheet_name,
                header=None,
                dtype=object,
                engine="openpyxl",
            )
            for sheet_name in names
        }


def _find_zone_headers(raw: pd.DataFrame) -> List[ZoneHeader]:
    """不再寫死 D1/D2/D3；看到「場站名稱」標題，就把同列的區域代碼當成 zone。"""
    headers: List[ZoneHeader] = []
    if raw.empty:
        return headers

    for row_index in range(len(raw)):
        row = raw.iloc[row_index]
        station_col = None
        for col_index, value in enumerate(row.tolist()):
            if "場站名稱" in text(value):
                station_col = col_index
                break
        if station_col is None:
            continue

        zone = ""
        for col_index in range(station_col):
            candidate = text(raw.iat[row_index, col_index])
            normalized = normalize_zone(candidate)
            if not candidate:
                continue
            if any(token in normalized for token in ("行政區", "場站編號", "站點編號", "編號")):
                continue
            zone = candidate
            break
        if not zone:
            continue

        headers.append(
            ZoneHeader(
                row_index=row_index,
                zone=zone,
                station_col=station_col,
                region_col=max(0, station_col - 1),
                code_col=0,
            )
        )
    return headers


def route_header_rows(raw: pd.DataFrame) -> List[Tuple[int, str]]:
    """向下相容舊 dispatch_core API，但區域名稱改為動態。"""
    return [(header.row_index, header.zone) for header in _find_zone_headers(raw)]


def available_sources(sheets: Dict[str, pd.DataFrame]) -> List[Tuple[str, str]]:
    options: List[Tuple[str, str]] = []
    for sheet_name, raw in sheets.items():
        for header in _find_zone_headers(raw):
            options.append((sheet_name, header.zone))
    return options


def discover_zones(sheets: Dict[str, pd.DataFrame]) -> List[str]:
    zones: OrderedDict[str, str] = OrderedDict()
    for _, zone in available_sources(sheets):
        key = normalize_zone(zone)
        if key and key not in zones:
            zones[key] = zone
    return list(zones.values())


def _target_header(raw: pd.DataFrame, route: str) -> tuple[ZoneHeader, int] | None:
    headers = _find_zone_headers(raw)
    target_key = normalize_zone(route)
    for index, header in enumerate(headers):
        if normalize_zone(header.zone) != target_key:
            continue
        end_row = headers[index + 1].row_index if index + 1 < len(headers) else len(raw)
        return header, end_row
    return None


def _station_rows(raw: pd.DataFrame, route: str) -> Iterable[tuple[int, ZoneHeader]]:
    resolved = _target_header(raw, route)
    if resolved is None:
        return
    header, end_row = resolved
    for row_index in range(header.row_index + 1, end_row):
        yield row_index, header


def parse_route(raw: pd.DataFrame, route: str, shift: str) -> pd.DataFrame:
    """擷取任意區域代碼與班別；不再只接受 D1/D2/D3。"""
    if shift not in SHIFT_COLUMNS:
        raise KeyError(f"未知班別：{shift}")
    bike_col, ebike_col = SHIFT_COLUMNS[shift]
    rows: list[dict] = []

    for row_index, header in _station_rows(raw, route) or []:
        required_col = max(header.station_col, header.region_col, bike_col, ebike_col)
        if raw.shape[1] <= required_col:
            continue

        station = text(raw.iat[row_index, header.station_col])
        region = text(raw.iat[row_index, header.region_col])
        station_code = raw.iat[row_index, header.code_col] if raw.shape[1] > header.code_col else None
        bike = pd.to_numeric(raw.iat[row_index, bike_col], errors="coerce")
        ebike = pd.to_numeric(raw.iat[row_index, ebike_col], errors="coerce")

        if not station:
            continue
        if station_code is not None and pd.isna(station_code):
            continue
        if pd.isna(bike) and pd.isna(ebike):
            continue

        rows.append(
            {
                "行政區": region or "未分類",
                "場站名稱": station,
                "路線區域": header.zone,
                "2.0 現況": safe_int(bike),
                "2.0E 現況": safe_int(ebike),
                "2.0 標準": safe_int(bike),
                "2.0E 標準": safe_int(ebike),
            }
        )
    return pd.DataFrame(rows)


def build_zone_station_map(sheets: Dict[str, pd.DataFrame]) -> dict[str, list[dict]]:
    """把『同區但不同平/假日 Sheet』合併成一個電池查詢範圍。

    例如 D1 平日 + D1 假日 → UI 只顯示一個 D1；
    D2、D3 同一工作表 → 仍各自形成 D2 / D3。
    """
    output: OrderedDict[str, OrderedDict[str, dict]] = OrderedDict()
    display_names: dict[str, str] = {}

    for _, raw in sheets.items():
        for header in _find_zone_headers(raw):
            zone_key = normalize_zone(header.zone)
            display_names.setdefault(zone_key, header.zone)
            output.setdefault(zone_key, OrderedDict())
            resolved = _target_header(raw, header.zone)
            if resolved is None:
                continue
            _, end_row = resolved
            for row_index in range(header.row_index + 1, end_row):
                if raw.shape[1] <= header.station_col:
                    continue
                station = text(raw.iat[row_index, header.station_col])
                if not station:
                    continue
                district = text(raw.iat[row_index, header.region_col]) if raw.shape[1] > header.region_col else ""
                station_key = re.sub(r"\s+", "", unicodedata.normalize("NFKC", station)).lower()
                output[zone_key].setdefault(
                    station_key,
                    {"name": station, "district": district or "未分類"},
                )

    return {
        display_names[zone_key]: list(stations.values())
        for zone_key, stations in output.items()
        if stations
    }


def build_workbook_profile(source: bytes | str) -> dict:
    sheets = load_workbook(source, visible_only=True)
    sources = available_sources(sheets)
    zones = discover_zones(sheets)
    sheet_zones: dict[str, list[str]] = {}
    for sheet_name, zone in sources:
        sheet_zones.setdefault(sheet_name, [])
        if zone not in sheet_zones[sheet_name]:
            sheet_zones[sheet_name].append(zone)
    return {
        "visible_sheets": list(sheets.keys()),
        "sources": sources,
        "zones": zones,
        "sheet_zones": sheet_zones,
        "zone_station_map": build_zone_station_map(sheets),
        "shift_options": list(SHIFT_COLUMNS.keys()),
    }
