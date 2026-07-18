from __future__ import annotations

import hashlib
import json
import math
import re
import time
import unicodedata
import uuid
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image, ImageEnhance, ImageOps
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from dispatch_core import (
    SHIFT_COLUMNS,
    available_sources,
    build_result,
    calculate_totals,
    load_workbook,
    parse_route,
)


def is_mobile_browser() -> bool:
    """依瀏覽器 User-Agent 判斷是否為手機或平板。"""
    try:
        user_agent = st.context.headers.get("User-Agent", "")
    except Exception:
        # 舊版 Streamlit 沒有 st.context 時，改由側邊欄手動切換。
        return False

    mobile_tokens = ("android", "iphone", "ipad", "ipod", "mobile", "windows phone")
    normalized_user_agent = user_agent.lower()
    return any(token in normalized_user_agent for token in mobile_tokens)


def safe_nonnegative_int(value) -> int:
    """把 Excel 儲存格值安全轉成不小於 0 的整數。"""
    try:
        return max(0, int(value))
    except (TypeError, ValueError, OverflowError):
        return 0


def add_dispatch_indicator(value) -> str:
    """在分析結果後方加上容易辨識的缺車／多車圖案。"""
    text = str(value)
    if "多" in text:
        return f"{text} 🔴"
    if "缺" in text or "少" in text:
        return f"{text} 🟠"
    return text


def extract_dispatch_count(value) -> int | None:
    """只擷取缺車／多車的台數；「符合」不納入排序條件。"""
    text = str(value)
    if not any(status in text for status in ("多", "缺", "少")):
        return None

    match = re.search(r"\d+", text)
    return int(match.group()) if match else 0


SORT_FIELD_OPTIONS = {
    "最大缺／多台數": "max",
    "缺／多總台數": "total",
    "2.0 缺／多台數": "bike",
    "2.0E 缺／多台數": "ebike",
    "場站名稱": "station",
    "Excel 原始順序": "original",
}


def sort_dispatch_results(result_df, sort_field: str, descending: bool):
    """依使用者選擇排序；「符合」以 -1 表示，不會被當成缺／多台數。"""
    if result_df.empty:
        return result_df

    sorted_df = result_df.copy()
    bike_counts = sorted_df["2.0 缺／多幾台"].map(extract_dispatch_count)
    ebike_counts = sorted_df["2.0E 缺／多幾台"].map(extract_dispatch_count)

    valid_counts_per_row = [
        [int(count) for count in (bike, ebike) if pd.notna(count)]
        for bike, ebike in zip(bike_counts, ebike_counts)
    ]
    sorted_df["_排序最大台數"] = [
        max(valid_counts, default=-1) for valid_counts in valid_counts_per_row
    ]
    sorted_df["_排序總台數"] = [
        sum(valid_counts) for valid_counts in valid_counts_per_row
    ]
    sorted_df["_排序2.0台數"] = bike_counts.fillna(-1).astype(int)
    sorted_df["_排序2.0E台數"] = ebike_counts.fillna(-1).astype(int)
    sorted_df["_原始順序"] = range(len(sorted_df))

    sort_key = SORT_FIELD_OPTIONS.get(sort_field, "max")
    ascending = not descending

    if sort_key == "max":
        by = ["_排序最大台數", "_排序總台數", "_排序2.0台數", "_排序2.0E台數", "_原始順序"]
        ascending_values = [ascending, ascending, ascending, ascending, True]
    elif sort_key == "total":
        by = ["_排序總台數", "_排序最大台數", "_排序2.0台數", "_排序2.0E台數", "_原始順序"]
        ascending_values = [ascending, ascending, ascending, ascending, True]
    elif sort_key == "bike":
        by = ["_排序2.0台數", "_排序2.0E台數", "_排序總台數", "_原始順序"]
        ascending_values = [ascending, ascending, ascending, True]
    elif sort_key == "ebike":
        by = ["_排序2.0E台數", "_排序2.0台數", "_排序總台數", "_原始順序"]
        ascending_values = [ascending, ascending, ascending, True]
    elif sort_key == "station":
        by = ["場站名稱", "_原始順序"]
        ascending_values = [ascending, True]
    else:
        by = ["_原始順序"]
        ascending_values = [ascending]

    sorted_df = sorted_df.sort_values(
        by=by,
        ascending=ascending_values,
        kind="mergesort",
    )

    return sorted_df.drop(
        columns=[
            "_排序最大台數",
            "_排序總台數",
            "_排序2.0台數",
            "_排序2.0E台數",
            "_原始順序",
        ]
    ).reset_index(drop=True)


def make_colored_export_df(result_df: pd.DataFrame) -> pd.DataFrame:
    """建立含紅／橘圖案的匯出資料，CSV 可直接辨識缺車與多車。"""
    export_df = result_df[
        ["行政區", "場站名稱", "2.0 缺／多幾台", "2.0E 缺／多幾台"]
    ].copy()
    for status_column in ("2.0 缺／多幾台", "2.0E 缺／多幾台"):
        export_df[status_column] = export_df[status_column].map(add_dispatch_indicator)
    return export_df


def build_colored_excel(export_df: pd.DataFrame) -> bytes:
    """輸出真正帶有儲存格底色的 Excel 分析表。"""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "調度分析"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    red_fill = PatternFill("solid", fgColor="F4CCCC")
    orange_fill = PatternFill("solid", fgColor="FCE5CD")
    green_fill = PatternFill("solid", fgColor="D9EAD3")

    headers = list(export_df.columns)
    worksheet.append(headers)
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for record in export_df.itertuples(index=False, name=None):
        worksheet.append(list(record))

    status_columns = {3, 4}
    for row in worksheet.iter_rows(min_row=2):
        for column_index, cell in enumerate(row, start=1):
            cell.alignment = Alignment(
                horizontal="center" if column_index in status_columns else "left",
                vertical="center",
            )
            if column_index not in status_columns:
                continue
            text = str(cell.value or "")
            if "多" in text:
                cell.fill = red_fill
            elif "缺" in text or "少" in text:
                cell.fill = orange_fill
            elif "符合" in text:
                cell.fill = green_fill

    widths = [14, 34, 20, 20]
    for column_index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column_index)].width = width

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.sheet_view.showGridLines = False

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()



OCR_MATCH_THRESHOLD = 0.62
OCR_MIN_TEXT_CONFIDENCE = 0.35
OCR_MAX_COUNT = 200


def normalize_ocr_text(value) -> str:
    """統一全形字、空白與常見 OCR 符號，便於後續比對。"""
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = text.replace("｜", "|").replace("—", "-").replace("–", "-")
    return re.sub(r"\s+", " ", text).strip()


def normalize_station_key(value) -> str:
    """建立只保留中英數字的場站比對鍵。"""
    text = normalize_ocr_text(value).lower()
    return re.sub(r"[^0-9a-z\u3400-\u9fff]", "", text)


def partial_similarity(needle: str, haystack: str) -> float:
    """計算短場站名稱在較長 OCR 列文字中的局部相似度。"""
    needle = normalize_station_key(needle)
    haystack = normalize_station_key(haystack)
    if not needle or not haystack:
        return 0.0
    if needle in haystack:
        return 1.0

    direct = SequenceMatcher(None, needle, haystack).ratio()
    if len(haystack) <= len(needle):
        return direct

    best = direct
    minimum = max(1, len(needle) - 2)
    maximum = min(len(haystack), len(needle) + 3)
    for window_size in range(minimum, maximum + 1):
        for start in range(0, len(haystack) - window_size + 1):
            window = haystack[start : start + window_size]
            best = max(best, SequenceMatcher(None, needle, window).ratio())
    return best


def safe_checkbox_value(value, default: bool = False) -> bool:
    """安全讀取 data_editor 核取方塊，避免空白值造成布林轉換錯誤。"""
    if value is None or pd.isna(value):
        return default
    return bool(value)


def safe_optional_count(value) -> int | None:
    """將可編輯表格中的數量安全轉成整數；空白維持 None。"""
    if value is None or (isinstance(value, float) and math.isnan(value)) or pd.isna(value):
        return None
    try:
        number = int(float(value))
    except (TypeError, ValueError, OverflowError):
        return None
    return number if 0 <= number <= OCR_MAX_COUNT else None


def parse_count_token(value) -> int | None:
    """辨識單一 OCR 文字框是否為合理車數。"""
    text = normalize_ocr_text(value)
    if not text or any(symbol in text for symbol in (":", "/", "%")):
        return None
    # 2.0、2.0E 是欄位名稱，不是車數。
    if re.search(r"2\s*[.．]\s*0\s*[Ee]?", text):
        return None

    compact = re.sub(r"\s+", "", text)
    compact = compact.translate(str.maketrans({"O": "0", "o": "0", "〇": "0", "○": "0"}))
    match = re.fullmatch(r"[^0-9]*([0-9]{1,3})(?:台)?[^0-9]*", compact)
    if not match:
        return None
    number = int(match.group(1))
    return number if 0 <= number <= OCR_MAX_COUNT else None


def extract_numbers_from_row(row_items: list[dict]) -> list[tuple[int, float]]:
    """依畫面由左到右，取出一列中看起來像車數的數字。"""
    numbers: list[tuple[int, float]] = []
    for item in sorted(row_items, key=lambda current: current["x"]):
        number = parse_count_token(item["text"])
        if number is not None:
            numbers.append((number, float(item["x"])))

    if numbers:
        return numbers

    # 某些畫面會把整列辨識成單一文字框，改用整列文字做後備解析。
    row_text = " ".join(str(item["text"]) for item in sorted(row_items, key=lambda current: current["x"]))
    row_text = normalize_ocr_text(row_text)
    row_text = re.sub(r"2\s*[.．]\s*0\s*[Ee]?", " ", row_text, flags=re.IGNORECASE)
    row_text = row_text.translate(str.maketrans({"O": "0", "o": "0", "〇": "0", "○": "0"}))
    for match in re.finditer(r"(?<![0-9.])([0-9]{1,3})(?![0-9.])", row_text):
        number = int(match.group(1))
        if 0 <= number <= OCR_MAX_COUNT:
            numbers.append((number, float(match.start())))
    return numbers


def group_ocr_items_into_rows(items: list[dict]) -> list[list[dict]]:
    """依文字框垂直位置合併成表格列。"""
    if not items:
        return []

    heights = [max(1.0, float(item["height"])) for item in items]
    median_height = float(np.median(heights)) if heights else 20.0
    tolerance = max(10.0, median_height * 0.72)
    rows: list[dict] = []

    for item in sorted(items, key=lambda current: (current["y"], current["x"])):
        target = None
        target_distance = None
        for row in rows:
            distance = abs(float(item["y"]) - float(row["center_y"]))
            if distance <= tolerance and (target_distance is None or distance < target_distance):
                target = row
                target_distance = distance
        if target is None:
            rows.append({"center_y": float(item["y"]), "items": [item]})
        else:
            target["items"].append(item)
            target["center_y"] = sum(float(current["y"]) for current in target["items"]) / len(target["items"])

    return [sorted(row["items"], key=lambda current: current["x"]) for row in sorted(rows, key=lambda current: current["center_y"])]


@st.cache_resource(show_spinner=False)
def get_local_ocr_engine():
    """載入免費本機 RapidOCR；模型只在第一次載入。"""
    from rapidocr import RapidOCR

    return RapidOCR()


def decode_rapidocr_output(result) -> list[dict]:
    """同時相容 RapidOCR 新版物件輸出與較舊版列表輸出。"""
    decoded: list[dict] = []

    boxes = getattr(result, "boxes", None)
    texts = getattr(result, "txts", None)
    scores = getattr(result, "scores", None)
    if boxes is not None and texts is not None:
        scores = scores if scores is not None else [1.0] * len(texts)
        for box, text, score in zip(boxes, texts, scores):
            points = np.asarray(box, dtype=float)
            if points.size < 8:
                continue
            x_min, x_max = float(points[:, 0].min()), float(points[:, 0].max())
            y_min, y_max = float(points[:, 1].min()), float(points[:, 1].max())
            decoded.append(
                {
                    "text": normalize_ocr_text(text),
                    "score": float(score or 0.0),
                    "x": x_min,
                    "y": (y_min + y_max) / 2,
                    "width": max(1.0, x_max - x_min),
                    "height": max(1.0, y_max - y_min),
                }
            )
        return decoded

    legacy = result[0] if isinstance(result, tuple) and result else result
    if isinstance(legacy, list):
        for entry in legacy:
            if not isinstance(entry, (list, tuple)) or len(entry) < 3:
                continue
            box, text, score = entry[0], entry[1], entry[2]
            points = np.asarray(box, dtype=float)
            if points.size < 8:
                continue
            x_min, x_max = float(points[:, 0].min()), float(points[:, 0].max())
            y_min, y_max = float(points[:, 1].min()), float(points[:, 1].max())
            decoded.append(
                {
                    "text": normalize_ocr_text(text),
                    "score": float(score or 0.0),
                    "x": x_min,
                    "y": (y_min + y_max) / 2,
                    "width": max(1.0, x_max - x_min),
                    "height": max(1.0, y_max - y_min),
                }
            )
    return decoded


def prepare_image_for_ocr(file_bytes: bytes) -> np.ndarray:
    """校正方向、對比與清晰度，並在照片過小時放大。"""
    image = Image.open(BytesIO(file_bytes))
    image = ImageOps.exif_transpose(image).convert("RGB")
    if image.width < 1600:
        scale = min(2.0, 1600 / max(1, image.width))
        image = image.resize((int(image.width * scale), int(image.height * scale)), Image.Resampling.LANCZOS)
    image = ImageOps.autocontrast(image, cutoff=1)
    image = ImageEnhance.Sharpness(image).enhance(1.35)
    return np.asarray(image)


def run_ocr_on_photo(file_name: str, file_bytes: bytes) -> tuple[list[list[dict]], list[dict]]:
    """辨識單張照片，回傳表格列與原始文字框。"""
    engine = get_local_ocr_engine()
    image_array = prepare_image_for_ocr(file_bytes)
    result = engine(image_array)
    items = [item for item in decode_rapidocr_output(result) if item["text"] and item["score"] >= OCR_MIN_TEXT_CONFIDENCE]
    for item in items:
        item["file_name"] = file_name
    return group_ocr_items_into_rows(items), items


def pick_station_for_row(row_text: str, station_names: list[str]) -> tuple[str | None, float]:
    """從 Excel 場站清單中找出最接近 OCR 列文字的場站。"""
    best_station = None
    best_score = 0.0
    for station_name in station_names:
        score = partial_similarity(station_name, row_text)
        if score > best_score:
            best_station = station_name
            best_score = score
    if best_score < OCR_MATCH_THRESHOLD:
        return None, best_score
    return best_station, best_score


def choose_counts(numbers: list[tuple[int, float]], bike_from_right: int, ebike_from_right: int) -> tuple[int | None, int | None]:
    """依使用者設定，從每列最右方數字選出 2.0 與 2.0E。"""
    values = [number for number, _ in numbers]

    def from_right(position: int) -> int | None:
        if position <= 0 or len(values) < position:
            return None
        return values[-position]

    return from_right(bike_from_right), from_right(ebike_from_right)


def analyze_station_photos(
    uploaded_photos,
    station_df: pd.DataFrame,
    bike_from_right: int,
    ebike_from_right: int,
) -> dict:
    """辨識多張滾動頁面照片，合併重複場站並整理異常。"""
    station_names = [str(value).strip() for value in station_df["場站名稱"].tolist() if str(value).strip()]
    current_by_station = {
        str(row["場站名稱"]): (
            safe_nonnegative_int(row["2.0 現況"]),
            safe_nonnegative_int(row["2.0E 現況"]),
        )
        for _, row in station_df.iterrows()
    }

    valid_observations: dict[str, list[dict]] = {name: [] for name in station_names}
    matched_incomplete: dict[str, list[dict]] = {name: [] for name in station_names}
    unmatched_rows: list[dict] = []
    raw_rows: list[dict] = []

    for uploaded_photo in uploaded_photos:
        file_name = str(uploaded_photo.name)
        rows, _ = run_ocr_on_photo(file_name, uploaded_photo.getvalue())
        best_per_station_in_photo: dict[str, dict] = {}

        for row_items in rows:
            row_text = " ".join(item["text"] for item in row_items).strip()
            if not row_text:
                continue
            row_score = float(np.mean([item["score"] for item in row_items]))
            numbers = extract_numbers_from_row(row_items)
            bike_count, ebike_count = choose_counts(numbers, bike_from_right, ebike_from_right)
            station_name, match_score = pick_station_for_row(row_text, station_names)
            record = {
                "檔案": file_name,
                "OCR文字": row_text,
                "場站名稱": station_name or "",
                "2.0 現況": bike_count,
                "2.0E 現況": ebike_count,
                "文字信心": round(row_score, 3),
                "場站比對": round(match_score, 3),
            }
            raw_rows.append(record)

            if station_name:
                evidence_score = (row_score * 0.45) + (match_score * 0.55)
                record["綜合信心"] = round(evidence_score, 3)
                if bike_count is not None and ebike_count is not None:
                    previous = best_per_station_in_photo.get(station_name)
                    if previous is None or record["綜合信心"] > previous["綜合信心"]:
                        best_per_station_in_photo[station_name] = record
                else:
                    matched_incomplete[station_name].append(record)
            else:
                contains_cjk = bool(re.search(r"[\u3400-\u9fff]", row_text))
                if contains_cjk and numbers:
                    unmatched_rows.append(record)

        for station_name, record in best_per_station_in_photo.items():
            valid_observations[station_name].append(record)

    normal_rows: list[dict] = []
    anomaly_rows: list[dict] = []
    observed_station_names: set[str] = set()

    for station_name in station_names:
        observations = valid_observations.get(station_name, [])
        unique_values = sorted({(record["2.0 現況"], record["2.0E 現況"]) for record in observations})
        if len(unique_values) == 1:
            bike_count, ebike_count = unique_values[0]
            observed_station_names.add(station_name)
            normal_rows.append(
                {
                    "套用": True,
                    "場站名稱": station_name,
                    "2.0 現況": bike_count,
                    "2.0E 現況": ebike_count,
                    "出現次數": len(observations),
                    "來源": "、".join(dict.fromkeys(record["檔案"] for record in observations)),
                    "信心": round(float(np.mean([record.get("綜合信心", 0.0) for record in observations])), 3),
                }
            )
            continue

        if len(unique_values) > 1:
            observed_station_names.add(station_name)
            candidates = "；".join(f"2.0={bike}、2.0E={ebike}" for bike, ebike in unique_values)
            anomaly_rows.append(
                {
                    "套用": False,
                    "狀態": "數據衝突",
                    "場站名稱": station_name,
                    "2.0 現況": None,
                    "2.0E 現況": None,
                    "說明": candidates,
                }
            )
            continue

        incomplete = matched_incomplete.get(station_name, [])
        if incomplete:
            observed_station_names.add(station_name)
            current_bike, current_ebike = current_by_station[station_name]
            anomaly_rows.append(
                {
                    "套用": False,
                    "狀態": "數字不完整",
                    "場站名稱": station_name,
                    "2.0 現況": current_bike,
                    "2.0E 現況": current_ebike,
                    "說明": "｜".join(dict.fromkeys(record["OCR文字"] for record in incomplete))[:240],
                }
            )

    # 無法對應場站的可疑表格列，去除重複後提供人工指定。
    unmatched_seen: set[tuple] = set()
    for record in unmatched_rows:
        key = (normalize_station_key(record["OCR文字"]), record["2.0 現況"], record["2.0E 現況"])
        if key in unmatched_seen:
            continue
        unmatched_seen.add(key)
        anomaly_rows.append(
            {
                "套用": False,
                "狀態": "無法對應場站",
                "場站名稱": "",
                "2.0 現況": record["2.0 現況"],
                "2.0E 現況": record["2.0E 現況"],
                "說明": f"{record['檔案']}｜{record['OCR文字']}"[:240],
            }
        )

    # 完全沒出現在任何照片中的場站列出，但預設不套用，避免覆蓋原資料。
    for station_name in station_names:
        if station_name in observed_station_names:
            continue
        current_bike, current_ebike = current_by_station[station_name]
        anomaly_rows.append(
            {
                "套用": False,
                "狀態": "照片未辨識到",
                "場站名稱": station_name,
                "2.0 現況": current_bike,
                "2.0E 現況": current_ebike,
                "說明": "可直接修改數量並勾選套用，或維持不勾選保留原資料。",
            }
        )

    return {
        "normal_df": pd.DataFrame(normal_rows),
        "anomaly_df": pd.DataFrame(anomaly_rows),
        "raw_df": pd.DataFrame(raw_rows),
        "photo_count": len(uploaded_photos),
    }


def apply_ocr_updates_to_dataframe(base_df: pd.DataFrame, updates: dict[str, tuple[int, int]]) -> pd.DataFrame:
    """把確認後的 OCR 數量套入完整場站現況。"""
    updated_df = base_df.copy()
    for row_index, row in updated_df.iterrows():
        station_name = str(row.get("場站名稱", ""))
        if station_name not in updates:
            continue
        bike_count, ebike_count = updates[station_name]
        updated_df.at[row_index, "2.0 現況"] = bike_count
        updated_df.at[row_index, "2.0E 現況"] = ebike_count
    return updated_df

BASE_CACHE_HOURS = 5
BASE_CACHE_SECONDS = BASE_CACHE_HOURS * 60 * 60
BASE_CACHE_DIR = Path(__file__).resolve().parent / ".base_cache"


def get_base_token() -> str | None:
    """從網址參數取得本次瀏覽器使用的基底識別碼。"""
    try:
        token = st.query_params.get("base")
    except Exception:
        params = st.experimental_get_query_params()
        values = params.get("base", [])
        token = values[0] if values else None

    if isinstance(token, list):
        token = token[0] if token else None

    token = str(token or "").strip()
    if len(token) != 32 or any(char not in "0123456789abcdef" for char in token.lower()):
        return None
    return token.lower()


def set_base_token(token: str) -> None:
    """把基底識別碼寫入網址，讓重新整理後仍能找到同一份基底。"""
    try:
        st.query_params["base"] = token
    except Exception:
        params = st.experimental_get_query_params()
        params["base"] = token
        st.experimental_set_query_params(**params)


def clear_base_token() -> None:
    """移除已失效的基底識別碼。"""
    try:
        if "base" in st.query_params:
            del st.query_params["base"]
    except Exception:
        params = st.experimental_get_query_params()
        params.pop("base", None)
        st.experimental_set_query_params(**params)


def base_cache_paths(token: str) -> tuple[Path, Path]:
    return BASE_CACHE_DIR / f"{token}.xlsx", BASE_CACHE_DIR / f"{token}.json"


def current_status_cache_path(token: str) -> Path:
    """取得指定基底所對應的現況暫存檔路徑。"""
    return BASE_CACHE_DIR / f"{token}.status.json"


def delete_cached_base(token: str) -> None:
    """刪除指定的暫存基底，以及與它綁定的現況資料。"""
    excel_path, metadata_path = base_cache_paths(token)
    status_path = current_status_cache_path(token)
    for path in (excel_path, metadata_path, status_path):
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass


def load_cached_base(token: str | None) -> tuple[dict | None, bool]:
    """讀取尚未超過 5 小時的基底；第二個回傳值表示是否剛過期。"""
    if not token:
        return None, False

    excel_path, metadata_path = base_cache_paths(token)
    if not excel_path.exists() or not metadata_path.exists():
        return None, False

    try:
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        uploaded_at = float(metadata["uploaded_at"])
        expires_at = uploaded_at + BASE_CACHE_SECONDS

        if time.time() >= expires_at:
            delete_cached_base(token)
            return None, True

        file_bytes = excel_path.read_bytes()
        return {
            "token": token,
            "name": str(metadata.get("name") or "配置基底.xlsx"),
            "bytes": file_bytes,
            "sha256": str(metadata.get("sha256") or hashlib.sha256(file_bytes).hexdigest()),
            "uploaded_at": uploaded_at,
            "expires_at": expires_at,
        }, False
    except (OSError, ValueError, TypeError, KeyError, json.JSONDecodeError):
        delete_cached_base(token)
        return None, False


def load_cached_status(token: str, expires_at: float) -> dict:
    """讀取與基底綁定的現況資料；基底到期後不再保留。"""
    if time.time() >= expires_at:
        try:
            current_status_cache_path(token).unlink(missing_ok=True)
        except OSError:
            pass
        return {"contexts": {}}

    status_path = current_status_cache_path(token)
    if not status_path.exists():
        return {"contexts": {}}

    try:
        payload = json.loads(status_path.read_text(encoding="utf-8"))
        contexts = payload.get("contexts", {})
        if not isinstance(contexts, dict):
            return {"contexts": {}}
        return {"contexts": contexts}
    except (OSError, TypeError, ValueError, json.JSONDecodeError):
        try:
            status_path.unlink(missing_ok=True)
        except OSError:
            pass
        return {"contexts": {}}


def save_cached_status(token: str, expires_at: float, payload: dict) -> None:
    """保存現況資料，但不延長基底原本的 5 小時期限。"""
    if time.time() >= expires_at:
        try:
            current_status_cache_path(token).unlink(missing_ok=True)
        except OSError:
            pass
        return

    BASE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    status_path = current_status_cache_path(token)
    encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

    try:
        if status_path.exists() and status_path.read_text(encoding="utf-8") == encoded:
            return
    except OSError:
        pass

    temporary_path = status_path.with_suffix(".status.tmp")
    try:
        temporary_path.write_text(encoded, encoding="utf-8")
        temporary_path.replace(status_path)
    except OSError:
        try:
            temporary_path.unlink(missing_ok=True)
        except OSError:
            pass


def status_context_key(sheet_name: str, route: str, shift: str) -> str:
    """為每個工作表、分區與班別建立獨立的現況保存區。"""
    return "｜".join((str(sheet_name), str(route), str(shift)))


def restore_current_status(base_df: pd.DataFrame, saved_records) -> pd.DataFrame:
    """把先前保存的現況套回剛解析完成的基底資料。"""
    restored_df = base_df.copy()
    if not isinstance(saved_records, list):
        return restored_df

    saved_values = {}
    for record in saved_records:
        if not isinstance(record, dict):
            continue
        key = (str(record.get("行政區", "")), str(record.get("場站名稱", "")))
        saved_values[key] = (
            safe_nonnegative_int(record.get("2.0 現況")),
            safe_nonnegative_int(record.get("2.0E 現況")),
        )

    for row_index, row in restored_df.iterrows():
        key = (str(row.get("行政區", "")), str(row.get("場站名稱", "")))
        if key not in saved_values:
            continue
        bike_now, ebike_now = saved_values[key]
        restored_df.at[row_index, "2.0 現況"] = bike_now
        restored_df.at[row_index, "2.0E 現況"] = ebike_now

    return restored_df


def merge_current_status(full_df: pd.DataFrame, edited_df: pd.DataFrame) -> pd.DataFrame:
    """把目前畫面修改的現況合併回完整分區資料，避免切換行政區時遺失。"""
    merged_df = full_df.copy()
    edited_values = {
        (str(row.get("行政區", "")), str(row.get("場站名稱", ""))): (
            safe_nonnegative_int(row.get("2.0 現況")),
            safe_nonnegative_int(row.get("2.0E 現況")),
        )
        for _, row in edited_df.iterrows()
    }

    for row_index, row in merged_df.iterrows():
        key = (str(row.get("行政區", "")), str(row.get("場站名稱", "")))
        if key not in edited_values:
            continue
        bike_now, ebike_now = edited_values[key]
        merged_df.at[row_index, "2.0 現況"] = bike_now
        merged_df.at[row_index, "2.0E 現況"] = ebike_now

    return merged_df


def dataframe_to_status_records(status_df: pd.DataFrame) -> list[dict]:
    """將完整現況轉成可寫入 JSON 的簡潔資料格式。"""
    records = []
    for _, row in status_df.iterrows():
        records.append(
            {
                "行政區": str(row.get("行政區", "")),
                "場站名稱": str(row.get("場站名稱", "")),
                "2.0 現況": safe_nonnegative_int(row.get("2.0 現況")),
                "2.0E 現況": safe_nonnegative_int(row.get("2.0E 現況")),
            }
        )
    return records


def save_cached_base(file_name: str, file_bytes: bytes) -> dict:
    """將新上傳的 Excel 暫存在伺服器 5 小時。"""
    BASE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    token = uuid.uuid4().hex
    uploaded_at = time.time()
    digest = hashlib.sha256(file_bytes).hexdigest()
    excel_path, metadata_path = base_cache_paths(token)

    excel_path.write_bytes(file_bytes)
    metadata_path.write_text(
        json.dumps(
            {
                "name": file_name,
                "uploaded_at": uploaded_at,
                "sha256": digest,
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    set_base_token(token)

    return {
        "token": token,
        "name": file_name,
        "bytes": file_bytes,
        "sha256": digest,
        "uploaded_at": uploaded_at,
        "expires_at": uploaded_at + BASE_CACHE_SECONDS,
    }


def format_remaining_time(expires_at: float) -> str:
    """把剩餘秒數轉成易讀的時、分格式。"""
    remaining_seconds = max(0, int(expires_at - time.time()))
    hours, remainder = divmod(remaining_seconds, 3600)
    minutes = remainder // 60
    return f"{hours} 小時 {minutes} 分鐘"


def rerun_app() -> None:
    """相容新舊版 Streamlit 的重新執行方法。"""
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()


def clear_editor_session_state(base_token: str | None = None) -> None:
    """清除資料編輯器與手機數字欄位的暫存，避免舊值蓋回歸零結果。"""
    token_prefix = f"editor::{base_token}::" if base_token else "editor::"
    keys_to_remove = [
        key
        for key in list(st.session_state.keys())
        if isinstance(key, str) and key.startswith(token_prefix)
    ]
    for key in keys_to_remove:
        del st.session_state[key]


def build_standard_status_cache(workbook_data: dict, options: list[tuple[str, str]]) -> dict:
    """把所有工作表、分區與班別的現況恢復成各自的標準配置數字。"""
    standard_contexts: dict[str, list[dict]] = {}

    for sheet_name, route in options:
        for shift in SHIFT_COLUMNS.keys():
            try:
                route_df = parse_route(workbook_data[sheet_name], route, shift)
            except Exception:
                continue

            if route_df.empty:
                continue

            standard_df = route_df.copy()
            standard_df["2.0 現況"] = standard_df["2.0 標準"].map(safe_nonnegative_int)
            standard_df["2.0E 現況"] = standard_df["2.0E 標準"].map(safe_nonnegative_int)
            context_key = status_context_key(sheet_name, route, shift)
            standard_contexts[context_key] = dataframe_to_status_records(standard_df)

    return {"contexts": standard_contexts}


def render_floating_station_search(station_df: pd.DataFrame, mobile_mode: bool) -> None:
    """在右下角建立懸浮場站搜尋，點選後跳到對應的現況調整列。"""
    stations = []
    for row_index, row in station_df.reset_index(drop=True).iterrows():
        station_name = str(row.get("場站名稱", "")).strip()
        if not station_name:
            continue
        stations.append(
            {
                "name": station_name,
                "region": str(row.get("行政區", "")).strip(),
                "index": int(row_index),
                "anchor": f"station-anchor-{row_index}",
            }
        )

    # 場站名稱來自 Excel，先用 JSON 安全編碼再嵌入 JavaScript。
    station_payload = json.dumps(stations, ensure_ascii=False).replace("</", "<\\/")
    display_mode = "mobile" if mobile_mode else "desktop"

    components.html(
        f"""
        <script>
        (() => {{
            const stations = {station_payload};
            const displayMode = {json.dumps(display_mode)};
            const doc = window.parent.document;
            const win = window.parent;

            const oldRoot = doc.getElementById("ubike-float-search");
            if (oldRoot) oldRoot.remove();

            const oldStyle = doc.getElementById("ubike-float-search-style");
            if (oldStyle) oldStyle.remove();

            const style = doc.createElement("style");
            style.id = "ubike-float-search-style";
            style.textContent = `
                #ubike-float-search {{
                    position: fixed;
                    right: 22px;
                    bottom: 22px;
                    z-index: 2147483000;
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI",
                                 "Microsoft JhengHei", sans-serif;
                }}
                #ubike-float-search * {{
                    box-sizing: border-box;
                }}
                #ubike-float-search .ufs-trigger {{
                    width: 58px;
                    height: 58px;
                    border: 0;
                    border-radius: 50%;
                    background: #ffbf00;
                    color: #141414;
                    font-size: 25px;
                    cursor: pointer;
                    box-shadow: 0 8px 28px rgba(0, 0, 0, 0.28);
                    transition: transform 0.16s ease, box-shadow 0.16s ease;
                }}
                #ubike-float-search .ufs-trigger:hover {{
                    transform: translateY(-2px);
                    box-shadow: 0 11px 32px rgba(0, 0, 0, 0.34);
                }}
                #ubike-float-search .ufs-panel {{
                    position: absolute;
                    right: 0;
                    bottom: 70px;
                    width: 340px;
                    max-width: calc(100vw - 28px);
                    padding: 13px;
                    border: 1px solid rgba(0, 0, 0, 0.13);
                    border-radius: 16px;
                    background: rgba(255, 255, 255, 0.98);
                    color: #171717;
                    box-shadow: 0 16px 46px rgba(0, 0, 0, 0.28);
                    backdrop-filter: blur(8px);
                }}
                #ubike-float-search .ufs-panel[hidden] {{
                    display: none !important;
                }}
                #ubike-float-search .ufs-title-row {{
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    margin-bottom: 9px;
                }}
                #ubike-float-search .ufs-title {{
                    font-size: 15px;
                    font-weight: 800;
                }}
                #ubike-float-search .ufs-close {{
                    border: 0;
                    background: transparent;
                    font-size: 20px;
                    line-height: 1;
                    cursor: pointer;
                    color: #555;
                }}
                #ubike-float-search .ufs-input {{
                    width: 100%;
                    min-height: 44px;
                    padding: 10px 12px;
                    border: 2px solid #e0e0e0;
                    border-radius: 11px;
                    outline: none;
                    font-size: 16px;
                    color: #111;
                    background: #fff;
                }}
                #ubike-float-search .ufs-input:focus {{
                    border-color: #ffbf00;
                    box-shadow: 0 0 0 3px rgba(255, 191, 0, 0.18);
                }}
                #ubike-float-search .ufs-hint {{
                    margin: 7px 2px 8px;
                    color: #666;
                    font-size: 12px;
                }}
                #ubike-float-search .ufs-results {{
                    display: flex;
                    flex-direction: column;
                    gap: 5px;
                    max-height: 310px;
                    overflow-y: auto;
                }}
                #ubike-float-search .ufs-result {{
                    width: 100%;
                    padding: 9px 10px;
                    border: 1px solid #e4e4e4;
                    border-radius: 10px;
                    background: #fff;
                    text-align: left;
                    cursor: pointer;
                }}
                #ubike-float-search .ufs-result:hover,
                #ubike-float-search .ufs-result:focus {{
                    border-color: #ffbf00;
                    background: #fff9df;
                    outline: none;
                }}
                #ubike-float-search .ufs-result-name {{
                    display: block;
                    color: #111;
                    font-size: 14px;
                    font-weight: 750;
                }}
                #ubike-float-search .ufs-result-region {{
                    display: block;
                    margin-top: 2px;
                    color: #777;
                    font-size: 12px;
                }}
                #ubike-float-search .ufs-empty {{
                    padding: 12px 4px 5px;
                    color: #777;
                    font-size: 13px;
                    text-align: center;
                }}
                #ubike-search-toast {{
                    position: fixed;
                    left: 50%;
                    bottom: 26px;
                    z-index: 2147483001;
                    transform: translateX(-50%);
                    padding: 10px 15px;
                    border-radius: 999px;
                    background: rgba(20, 20, 20, 0.92);
                    color: #fff;
                    font: 700 14px/1.25 -apple-system, BlinkMacSystemFont, "Segoe UI",
                          "Microsoft JhengHei", sans-serif;
                    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.28);
                    pointer-events: none;
                }}
                .ubike-search-row-flash {{
                    position: fixed;
                    z-index: 2147482500;
                    border: 3px solid #ff9f00;
                    border-radius: 8px;
                    background: rgba(255, 191, 0, 0.18);
                    box-shadow: 0 0 0 4px rgba(255, 191, 0, 0.13);
                    pointer-events: none;
                    animation: ubikeSearchPulse 0.75s ease-in-out 2;
                }}
                @keyframes ubikeSearchPulse {{
                    0%, 100% {{ opacity: 0.45; }}
                    50% {{ opacity: 1; }}
                }}
                @media (max-width: 700px) {{
                    #ubike-float-search {{
                        right: 12px;
                        bottom: 12px;
                    }}
                    #ubike-float-search .ufs-trigger {{
                        width: 54px;
                        height: 54px;
                    }}
                    #ubike-float-search .ufs-panel {{
                        position: fixed;
                        left: 12px;
                        right: 12px;
                        bottom: 76px;
                        width: auto;
                        max-width: none;
                    }}
                }}
            `;
            doc.head.appendChild(style);

            const root = doc.createElement("div");
            root.id = "ubike-float-search";
            root.innerHTML = `
                <div class="ufs-panel" hidden>
                    <div class="ufs-title-row">
                        <div class="ufs-title">🔎 快速搜尋場站</div>
                        <button class="ufs-close" type="button" aria-label="關閉搜尋">×</button>
                    </div>
                    <input class="ufs-input" type="search"
                           placeholder="輸入場站名稱或行政區"
                           autocomplete="off" />
                    <div class="ufs-hint">搜尋目前畫面中的場站｜快捷鍵 Ctrl + K</div>
                    <div class="ufs-results"></div>
                </div>
                <button class="ufs-trigger" type="button"
                        title="快速搜尋場站（Ctrl + K）"
                        aria-label="快速搜尋場站">🔎</button>
            `;
            doc.body.appendChild(root);

            const trigger = root.querySelector(".ufs-trigger");
            const panel = root.querySelector(".ufs-panel");
            const closeButton = root.querySelector(".ufs-close");
            const input = root.querySelector(".ufs-input");
            const results = root.querySelector(".ufs-results");

            function showToast(message) {{
                const previous = doc.getElementById("ubike-search-toast");
                if (previous) previous.remove();
                const toast = doc.createElement("div");
                toast.id = "ubike-search-toast";
                toast.textContent = message;
                doc.body.appendChild(toast);
                win.setTimeout(() => toast.remove(), 2100);
            }}

            function setOpen(open) {{
                panel.hidden = !open;
                if (open) {{
                    input.value = "";
                    renderResults("");
                    win.setTimeout(() => input.focus(), 30);
                }}
            }}

            function flashDesktopRow(editor, rowIndex) {{
                win.setTimeout(() => {{
                    const rect = editor.getBoundingClientRect();
                    const rowHeight = 38;
                    const headerHeight = 40;
                    const overlay = doc.createElement("div");
                    overlay.className = "ubike-search-row-flash";
                    overlay.style.left = `${{Math.max(0, rect.left)}}px`;
                    overlay.style.top = `${{rect.top + headerHeight + (rowIndex * rowHeight)}}px`;
                    overlay.style.width = `${{rect.width}}px`;
                    overlay.style.height = `${{rowHeight}}px`;
                    doc.body.appendChild(overlay);
                    win.setTimeout(() => overlay.remove(), 1700);
                }}, 560);
            }}

            function jumpToStation(station) {{
                setOpen(false);

                const anchor = doc.getElementById(station.anchor);
                if (anchor) {{
                    anchor.scrollIntoView({{ behavior: "smooth", block: "center" }});
                    showToast(`已跳到：${{station.name}}`);
                    return;
                }}

                const editor = doc.querySelector('div[data-testid="stDataEditor"]');
                if (!editor) {{
                    showToast("找不到現況調整表，請重新整理後再試");
                    return;
                }}

                const rowHeight = 38;
                const headerHeight = 40;
                const editorTop = win.scrollY + editor.getBoundingClientRect().top;
                const targetTop = Math.max(
                    0,
                    editorTop + headerHeight + (station.index * rowHeight) -
                    Math.min(170, win.innerHeight * 0.28)
                );

                win.scrollTo({{ top: targetTop, behavior: "smooth" }});
                flashDesktopRow(editor, station.index);
                showToast(`已跳到：${{station.name}}`);
            }}

            function createResultButton(station) {{
                const button = doc.createElement("button");
                button.type = "button";
                button.className = "ufs-result";

                const name = doc.createElement("span");
                name.className = "ufs-result-name";
                name.textContent = station.name;
                button.appendChild(name);

                if (station.region) {{
                    const region = doc.createElement("span");
                    region.className = "ufs-result-region";
                    region.textContent = station.region;
                    button.appendChild(region);
                }}

                button.addEventListener("click", () => jumpToStation(station));
                return button;
            }}

            function renderResults(query) {{
                const keyword = String(query || "").trim().toLocaleLowerCase("zh-TW");
                const matched = stations
                    .filter((station) => {{
                        const haystack = `${{station.name}} ${{station.region}}`
                            .toLocaleLowerCase("zh-TW");
                        return !keyword || haystack.includes(keyword);
                    }})
                    .slice(0, 10);

                results.replaceChildren();

                if (!matched.length) {{
                    const empty = doc.createElement("div");
                    empty.className = "ufs-empty";
                    empty.textContent = "查無符合的場站";
                    results.appendChild(empty);
                    return;
                }}

                matched.forEach((station) => results.appendChild(createResultButton(station)));
            }}

            trigger.addEventListener("click", () => setOpen(panel.hidden));
            closeButton.addEventListener("click", () => setOpen(false));
            input.addEventListener("input", (event) => renderResults(event.target.value));
            input.addEventListener("keydown", (event) => {{
                if (event.key === "Enter") {{
                    const firstResult = results.querySelector(".ufs-result");
                    if (firstResult) firstResult.click();
                }}
                if (event.key === "Escape") setOpen(false);
            }});

            if (win.__ubikeSearchKeyHandler) {{
                win.removeEventListener("keydown", win.__ubikeSearchKeyHandler);
            }}
            win.__ubikeSearchKeyHandler = (event) => {{
                if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {{
                    event.preventDefault();
                    setOpen(true);
                }}
                if (event.key === "Escape" && !panel.hidden) {{
                    setOpen(false);
                }}
            }};
            win.addEventListener("keydown", win.__ubikeSearchKeyHandler);

            renderResults("");
        }})();
        </script>
        """,
        height=0,
        scrolling=False,
    )


st.set_page_config(
    page_title="臺東 YouBike 智慧調度決策系統",
    page_icon="🚚",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.3rem; padding-bottom: 3rem;}
    [data-testid="stMetricValue"] {font-size: 1.65rem;}
    div[data-testid="stDataEditor"] {border: 1px solid #dddddd; border-radius: 10px;}
    div[data-testid="stNumberInput"] input {text-align: center;}

    /* 手機版：縮小頁面留白，並放大數字輸入框，方便直接叫出九宮格鍵盤。 */
    @media (max-width: 900px) {
        .block-container {
            padding-left: 0.45rem;
            padding-right: 0.45rem;
            padding-top: 0.7rem;
        }

        div[data-testid="stDataEditor"] {
            font-size: 0.78rem;
        }

        div[data-testid="stDataEditor"] [role="columnheader"],
        div[data-testid="stDataEditor"] [role="gridcell"] {
            padding-left: 2px !important;
            padding-right: 2px !important;
        }

        div[data-testid="stNumberInput"] input {
            min-height: 44px;
            font-size: 18px !important;
            font-weight: 700;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def cached_load_workbook(source: bytes | str):
    return load_workbook(source)


st.title("🚚 臺東 YouBike 智慧調度工具")
st.caption("Excel 作為配置基底｜自動區分 D1、D2、D3 與行政區｜分開計算 2.0、2.0E")

mobile_detected = is_mobile_browser()

base_token = get_base_token()
cached_base, cache_expired = load_cached_base(base_token)

if "base_uploader_version" not in st.session_state:
    st.session_state["base_uploader_version"] = 0

if cache_expired:
    clear_base_token()
    st.session_state["base_uploader_version"] += 1
    st.session_state["base_expired_notice"] = True
    base_token = None
    cached_base = None

with st.sidebar:
    st.header("配置基底")

    if st.session_state.pop("base_expired_notice", False):
        st.warning("⏰ 原配置基底已超過 5 小時，請重新上傳。")

    uploaded_excel = st.file_uploader(
        "上傳配置表（必須）",
        type=["xlsx"],
        key=f"base_uploader_{st.session_state['base_uploader_version']}",
        help="上傳後可維持 5 小時；期間重新整理頁面不必重傳，超過時間會自動失效。",
    )
    st.caption("📋 分析結果固定只顯示缺車／多車場站；兩項皆符合者不顯示。")
    mobile_input_mode = st.checkbox(
        "📱 手機九宮格輸入模式",
        value=mobile_detected,
        help="手機通常會自動啟用；若沒有自動偵測，可手動勾選。",
    )
    st.info("現況欄會自動保存，並與目前基底一起保留 5 小時。")

active_base = cached_base

if uploaded_excel is not None:
    uploaded_bytes = uploaded_excel.getvalue()
    uploaded_digest = hashlib.sha256(uploaded_bytes).hexdigest()

    # 同一個上傳元件在每次互動時都會重新執行，不能因此延長 5 小時計時。
    if active_base is None or active_base["sha256"] != uploaded_digest:
        active_base = save_cached_base(uploaded_excel.name, uploaded_bytes)

if st.session_state.pop("full_reset_notice", False):
    st.success("✅ 已完成全部重置：配置基底與所有保留數據都已清除。")

if st.session_state.pop("data_zero_notice", False):
    st.success("✅ 數據已歸零：配置基底保留，所有分區與班別的現況均已恢復為各自的標準配置數字。")

if active_base is None:
    st.info("📤 請先從左側上傳 Excel 配置表；上傳後 5 小時內重新整理頁面都不必重傳。")
    st.stop()

try:
    workbook_data = cached_load_workbook(active_base["bytes"])
    source_caption = f"已使用上傳檔案：{active_base['name']}"
except Exception as exc:
    st.error(f"Excel 讀取失敗：{exc}")
    st.stop()

options = available_sources(workbook_data)
if not options:
    st.error("這份 Excel 找不到 D1／D2／D3 的標準配置區塊。")
    st.stop()

# 頁面最上方的兩種重置功能。
reset_col_1, reset_col_2 = st.columns(2)
with reset_col_1:
    if st.button(
        "🗑️ 全部重置",
        use_container_width=True,
        type="primary",
        help="刪除目前配置基底、網址識別碼，以及所有已保存的現況數據。",
    ):
        reset_token = active_base.get("token")
        if reset_token:
            delete_cached_base(reset_token)
        clear_base_token()
        clear_editor_session_state(reset_token)
        st.session_state["base_uploader_version"] += 1
        st.session_state["full_reset_notice"] = True
        rerun_app()

with reset_col_2:
    if st.button(
        "0️⃣ 數據歸零",
        use_container_width=True,
        help="保留 Excel 配置基底，將所有 D1／D2／D3、所有班別的 2.0 與 2.0E 現況恢復為各自的標準配置數字。",
    ):
        standard_status_cache = build_standard_status_cache(workbook_data, options)
        save_cached_status(
            active_base["token"],
            active_base["expires_at"],
            standard_status_cache,
        )
        clear_editor_session_state(active_base["token"])
        st.session_state["data_zero_notice"] = True
        rerun_app()

st.caption(
    "🗑️ 全部重置：清除基底＋所有保留數據　｜　"
    "0️⃣ 數據歸零：基底不動，所有現況恢復成標準配置數字"
)

with st.sidebar:
    st.success(f"✅ 基底有效中：剩餘約 {format_remaining_time(active_base['expires_at'])}")
    st.caption(f"目前基底：{active_base['name']}")

option_labels = [f"{route}｜{sheet_name}" for sheet_name, route in options]

selector_1, selector_2 = st.columns([2, 1])
with selector_1:
    default_option = "D1｜115_D1暑假配置表"
    default_index = option_labels.index(default_option) if default_option in option_labels else 0
    selected_label = st.selectbox("📚 配置表版本與分區", option_labels, index=default_index)
with selector_2:
    selected_shift = st.selectbox("⏰ 班別", list(SHIFT_COLUMNS.keys()))

selected_index = option_labels.index(selected_label)
selected_sheet, selected_route = options[selected_index]
base_df = parse_route(workbook_data[selected_sheet], selected_route, selected_shift)

status_cache = load_cached_status(active_base["token"], active_base["expires_at"])
current_context_key = status_context_key(selected_sheet, selected_route, selected_shift)
base_df = restore_current_status(
    base_df,
    status_cache["contexts"].get(current_context_key, []),
)

if base_df.empty:
    st.warning("目前選擇的區塊沒有可用場站資料，請改選其他版本或班別。")
    st.stop()


# 多張平板翻拍照片辨識：套用範圍為目前選擇的工作表、D1／D2／D3 與班別。
ocr_state_key = (
    f"ocr_analysis::{active_base['token']}::{selected_sheet}::{selected_route}::{selected_shift}"
)
with st.expander("📷 多張照片辨識場站現況", expanded=False):
    st.caption(
        "可一次上傳多張滾動頁面照片。重複場站數字一致會自動採用；"
        "衝突、缺字、找不到場站及未拍到的場站會集中到可編輯異常區。"
    )

    uploaded_station_photos = st.file_uploader(
        "上傳平板翻拍照片",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key=f"station_photo_uploader::{active_base['token']}::{selected_sheet}::{selected_route}::{selected_shift}",
        help="照片可任意張數，不必固定一頁拍兩張；請盡量避免反光並讓場站名稱與數字保持清楚。",
    )

    number_rule_col_1, number_rule_col_2 = st.columns(2)
    with number_rule_col_1:
        bike_from_right = st.number_input(
            "2.0 取每列右邊第幾個數字",
            min_value=1,
            max_value=8,
            value=2,
            step=1,
            help="預設畫面每列最後兩個數字依序為 2.0、2.0E，所以 2.0 是右邊第 2 個。",
        )
    with number_rule_col_2:
        ebike_from_right = st.number_input(
            "2.0E 取每列右邊第幾個數字",
            min_value=1,
            max_value=8,
            value=1,
            step=1,
            help="預設 2.0E 是每列最右邊的數字。若實際平板畫面欄位不同，可在這裡調整。",
        )

    if st.button(
        "🔍 開始辨識照片",
        use_container_width=True,
        disabled=not uploaded_station_photos,
        key=f"run_ocr::{active_base['token']}::{selected_sheet}::{selected_route}::{selected_shift}",
    ):
        try:
            with st.spinner(f"正在辨識 {len(uploaded_station_photos)} 張照片並比對場站……"):
                analysis_result = analyze_station_photos(
                    uploaded_station_photos,
                    base_df,
                    int(bike_from_right),
                    int(ebike_from_right),
                )
                st.session_state[ocr_state_key] = analysis_result

                # 重複照片結果一致，或單張結果完整且可明確比對時，直接寫入現況。
                normal_updates: dict[str, tuple[int, int]] = {}
                normal_result_df = analysis_result.get("normal_df", pd.DataFrame())
                if not normal_result_df.empty:
                    for _, recognized_row in normal_result_df.iterrows():
                        station_name = str(recognized_row.get("場站名稱", "")).strip()
                        bike_count = safe_optional_count(recognized_row.get("2.0 現況"))
                        ebike_count = safe_optional_count(recognized_row.get("2.0E 現況"))
                        if station_name and bike_count is not None and ebike_count is not None:
                            normal_updates[station_name] = (bike_count, ebike_count)

                if normal_updates:
                    auto_updated_df = apply_ocr_updates_to_dataframe(base_df, normal_updates)
                    status_cache["contexts"][current_context_key] = dataframe_to_status_records(auto_updated_df)
                    save_cached_status(active_base["token"], active_base["expires_at"], status_cache)
                    clear_editor_session_state(active_base["token"])
                    st.session_state["ocr_auto_apply_notice"] = len(normal_updates)
                    rerun_app()
        except ImportError:
            st.error(
                "尚未安裝免費 OCR 套件。請先執行隨附的「安裝OCR.bat」，"
                "或在終端機輸入：python -m pip install rapidocr onnxruntime"
            )
        except Exception as exc:
            st.error(f"照片辨識失敗：{exc}")

    ocr_analysis = st.session_state.get(ocr_state_key)
    if ocr_analysis:
        normal_df = ocr_analysis.get("normal_df", pd.DataFrame()).copy()
        anomaly_df = ocr_analysis.get("anomaly_df", pd.DataFrame()).copy()
        raw_df = ocr_analysis.get("raw_df", pd.DataFrame()).copy()

        summary_col_1, summary_col_2, summary_col_3 = st.columns(3)
        summary_col_1.metric("已上傳照片", f"{ocr_analysis.get('photo_count', 0)} 張")
        summary_col_2.metric("可直接套用", f"{len(normal_df)} 個場站")
        summary_col_3.metric("待確認／未辨識", f"{len(anomaly_df)} 筆")

        st.markdown("#### ✅ 辨識一致，已自動寫入現況")
        if normal_df.empty:
            st.info("目前沒有能自動寫入的場站資料。請查看下方異常區或原始辨識文字。")
        else:
            normal_display_df = normal_df.drop(columns=["套用"], errors="ignore").copy()
            normal_display_df.insert(0, "狀態", "已自動寫入")
            st.dataframe(normal_display_df, hide_index=True, use_container_width=True)

        st.markdown("#### ⚠️ 無法識別或異常資料（可直接調整）")
        if anomaly_df.empty:
            st.success("沒有異常資料。")
            anomaly_edited_df = anomaly_df
        else:
            anomaly_edited_df = st.data_editor(
                anomaly_df,
                hide_index=True,
                use_container_width=True,
                num_rows="dynamic",
                key=f"ocr_anomaly_editor::{ocr_state_key}",
                disabled=["狀態", "說明"],
                column_config={
                    "套用": st.column_config.CheckboxColumn(
                        "套用",
                        default=False,
                        help="確認場站與數字後勾選，才會寫入現況。",
                    ),
                    "狀態": st.column_config.TextColumn("異常狀態", width="medium"),
                    "場站名稱": st.column_config.SelectboxColumn(
                        "場站名稱",
                        options=[""] + [str(value) for value in base_df["場站名稱"].tolist()],
                        required=False,
                        width="large",
                    ),
                    "2.0 現況": st.column_config.NumberColumn("2.0現況", min_value=0, step=1, format="%d"),
                    "2.0E 現況": st.column_config.NumberColumn("2.0E現況", min_value=0, step=1, format="%d"),
                    "說明": st.column_config.TextColumn("辨識內容／處理提示", width="large"),
                },
            )

        if st.button(
            "✅ 套用勾選的異常修正",
            type="primary",
            use_container_width=True,
            key=f"apply_ocr::{ocr_state_key}",
        ):
            updates: dict[str, tuple[int, int]] = {}
            errors: list[str] = []

            if not anomaly_edited_df.empty:
                for row_number, row in anomaly_edited_df.iterrows():
                    if not safe_checkbox_value(row.get("套用", False), default=False):
                        continue
                    station_name = str(row.get("場站名稱", "")).strip()
                    bike_count = safe_optional_count(row.get("2.0 現況"))
                    ebike_count = safe_optional_count(row.get("2.0E 現況"))
                    if not station_name:
                        errors.append(f"異常表第 {row_number + 1} 列尚未選擇場站")
                        continue
                    if bike_count is None or ebike_count is None:
                        errors.append(f"{station_name} 的 2.0 或 2.0E 數量尚未填完整")
                        continue
                    updates[station_name] = (bike_count, ebike_count)

            if errors:
                st.error("無法套用：" + "；".join(errors[:8]))
            elif not updates:
                st.warning("目前沒有勾選任何可套用的場站資料。")
            else:
                updated_full_df = apply_ocr_updates_to_dataframe(base_df, updates)
                status_cache["contexts"][current_context_key] = dataframe_to_status_records(updated_full_df)
                save_cached_status(active_base["token"], active_base["expires_at"], status_cache)
                clear_editor_session_state(active_base["token"])
                st.session_state["ocr_apply_notice"] = len(updates)
                rerun_app()

        with st.expander("查看 OCR 原始辨識列（除錯用）", expanded=False):
            if raw_df.empty:
                st.caption("沒有原始辨識資料。")
            else:
                st.dataframe(raw_df, hide_index=True, use_container_width=True)

auto_applied_ocr_count = st.session_state.pop("ocr_auto_apply_notice", None)
if auto_applied_ocr_count is not None:
    st.success(f"✅ 已自動寫入 {auto_applied_ocr_count} 個辨識一致的場站；異常資料可在照片辨識區內直接修正。")

applied_ocr_count = st.session_state.pop("ocr_apply_notice", None)
if applied_ocr_count is not None:
    st.success(f"✅ 已將 {applied_ocr_count} 個異常修正寫入現況，分析結果已同步更新。")

regions = ["全部"] + list(dict.fromkeys(base_df["行政區"].tolist()))
selected_region = st.selectbox("📍 行政區", regions)
working_df = base_df if selected_region == "全部" else base_df[base_df["行政區"] == selected_region]
working_df = working_df.reset_index(drop=True)

render_floating_station_search(working_df, mobile_input_mode)

st.caption(f"{source_caption}｜工作表：{selected_sheet}｜分區：{selected_route}｜共 {len(working_df)} 個場站")
st.subheader("✏️ 輸入現場現況")

editor_key = (
    f"editor::{active_base['token']}::{selected_sheet}::{selected_route}::"
    f"{selected_shift}::{selected_region}"
)

if mobile_input_mode:
    # st.number_input 會輸出 HTML number 欄位；手機點選時會直接叫出數字九宮格。
    st.caption("📱 手機九宮格輸入已啟用：點選數字框即可輸入現場車數。")
    edited_df = working_df.copy()

    header_station, header_bike, header_ebike = st.columns([1.55, 0.9, 0.95])
    header_station.markdown("**場站／標準**")
    header_bike.markdown("**2.0 現況**")
    header_ebike.markdown("**2.0E 現況**")

    for row_index, row in working_df.iterrows():
        st.markdown(
            f'<div id="station-anchor-{row_index}" class="station-anchor"></div>',
            unsafe_allow_html=True,
        )
        station_col, bike_col, ebike_col = st.columns([1.55, 0.9, 0.95])

        with station_col:
            st.markdown(f"**{row['場站名稱']}**")
            st.caption(
                f"標準：2.0＝{safe_nonnegative_int(row['2.0 標準'])}｜"
                f"2.0E＝{safe_nonnegative_int(row['2.0E 標準'])}"
            )

        with bike_col:
            bike_now = st.number_input(
                f"{row['場站名稱']} 2.0 現況",
                min_value=0,
                value=safe_nonnegative_int(row["2.0 現況"]),
                step=1,
                format="%d",
                key=f"{editor_key}::mobile::2.0::{row_index}",
                label_visibility="collapsed",
            )

        with ebike_col:
            ebike_now = st.number_input(
                f"{row['場站名稱']} 2.0E 現況",
                min_value=0,
                value=safe_nonnegative_int(row["2.0E 現況"]),
                step=1,
                format="%d",
                key=f"{editor_key}::mobile::2.0E::{row_index}",
                label_visibility="collapsed",
            )

        edited_df.at[row_index, "2.0 現況"] = int(bike_now)
        edited_df.at[row_index, "2.0E 現況"] = int(ebike_now)
else:
    # 電腦版維持原本可快速批次編輯的資料表。
    # 依場站筆數自動計算表格高度，完整展開所有列。
    editor_row_height = 38
    editor_header_height = 40
    editor_height = editor_header_height + (len(working_df) * editor_row_height) + 4

    edited_df = st.data_editor(
        working_df,
        key=editor_key,
        hide_index=True,
        use_container_width=True,
        height=editor_height,
        row_height=editor_row_height,
        column_order=["場站名稱", "2.0 現況", "2.0E 現況", "2.0 標準", "2.0E 標準"],
        disabled=["行政區", "場站名稱", "2.0 標準", "2.0E 標準"],
        column_config={
            # 輸入現場現況時隱藏行政區；資料仍保留供篩選、分析、分組及匯出使用。
            "行政區": None,
            "場站名稱": st.column_config.TextColumn("場站", width=120),
            "2.0 現況": st.column_config.NumberColumn(
                "2.0現", width=52, min_value=0, step=1, format="%d"
            ),
            "2.0E 現況": st.column_config.NumberColumn(
                "2.0E現", width=58, min_value=0, step=1, format="%d"
            ),
            "2.0 標準": st.column_config.NumberColumn("2.0標", width=52, format="%d"),
            "2.0E 標準": st.column_config.NumberColumn("2.0E標", width=58, format="%d"),
        },
    )

# 每次數字異動造成 Streamlit 重新執行時，自動把現況寫入與基底綁定的暫存檔。
full_status_df = merge_current_status(base_df, edited_df)
status_cache["contexts"][current_context_key] = dataframe_to_status_records(full_status_df)
save_cached_status(active_base["token"], active_base["expires_at"], status_cache)

st.caption(
    f"💾 現況已自動保存，將與基底同時於約 "
    f"{format_remaining_time(active_base['expires_at'])} 後到期。"
)

st.markdown("---")
st.subheader("📋 調度分析結果")

result_df = build_result(edited_df)

# 分析結果固定排除「2.0、2.0E 都符合」的場站。
# 若只有其中一種車型符合，該場站仍保留，但「符合」不參與排序。
result_df = result_df[
    (result_df["2.0 缺／多幾台"] != "符合")
    | (result_df["2.0E 缺／多幾台"] != "符合")
].reset_index(drop=True)

sort_control_1, sort_control_2 = st.columns([2, 1])
with sort_control_1:
    selected_sort_field = st.selectbox(
        "↕️ 分析結果排序欄位",
        list(SORT_FIELD_OPTIONS.keys()),
        index=0,
        help="排序會同步套用到畫面、CSV 與彩色 Excel。",
    )
with sort_control_2:
    selected_sort_direction = st.selectbox(
        "排序方向",
        ["由大到小／Z→A／倒序", "由小到大／A→Z／正序"],
        index=0,
    )

sort_descending = selected_sort_direction.startswith("由大到小")

# 保留行政區分段；每個行政區內依使用者選擇的方式排序。
sorted_region_frames = []
for region_name in result_df["行政區"].drop_duplicates():
    region_rows = result_df[result_df["行政區"] == region_name]
    sorted_region_frames.append(
        sort_dispatch_results(region_rows, selected_sort_field, sort_descending)
    )

if sorted_region_frames:
    result_df = pd.concat(sorted_region_frames, ignore_index=True)

bike_short, bike_extra = calculate_totals(edited_df, "2.0 現況", "2.0 標準")
ebike_short, ebike_extra = calculate_totals(edited_df, "2.0E 現況", "2.0E 標準")

m1, m2, m3, m4 = st.columns(4)
m1.metric("2.0 缺車合計", f"{bike_short} 台")
m2.metric("2.0 多車合計", f"{bike_extra} 台")
m3.metric("2.0E 缺車合計", f"{ebike_short} 台")
m4.metric("2.0E 多車合計", f"{ebike_extra} 台")

if result_df.empty:
    st.success("✨ 所有場站皆符合配置，目前不需要調度。")
else:
    st.caption("🔴 多車｜🟠 缺車｜下載檔會使用目前畫面的排序")

    # 依行政區分段顯示，但每張表維持使用者指定的三個欄位。
    for region, region_df in result_df.groupby("行政區", sort=False):
        st.markdown(f"#### {selected_route}｜{region}")

        display_df = region_df[["場站名稱", "2.0 缺／多幾台", "2.0E 缺／多幾台"]].copy()
        for status_column in ("2.0 缺／多幾台", "2.0E 缺／多幾台"):
            display_df[status_column] = display_df[status_column].map(add_dispatch_indicator)

        st.dataframe(
            display_df,
            hide_index=True,
            use_container_width=True,
        )

    export_df = make_colored_export_df(result_df)
    csv_data = export_df.to_csv(index=False).encode("utf-8-sig")
    excel_data = build_colored_excel(export_df)

    download_col_1, download_col_2 = st.columns(2)
    with download_col_1:
        st.download_button(
            "⬇️ 下載彩色標記 CSV",
            data=csv_data,
            file_name=f"{selected_route}_{selected_shift}_調度分析_彩色標記.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with download_col_2:
        st.download_button(
            "⬇️ 下載彩色 Excel",
            data=excel_data,
            file_name=f"{selected_route}_{selected_shift}_調度分析_彩色.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.caption("CSV 以 🔴／🟠 圖案表示顏色；Excel 會直接套用紅／橘儲存格底色。")
