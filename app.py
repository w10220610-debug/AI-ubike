from __future__ import annotations

import hashlib
import html
import json
import math
import os
import re
import time
import tempfile
import unicodedata
import uuid
from datetime import datetime
from difflib import SequenceMatcher
from functools import lru_cache
from io import BytesIO
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image, ImageEnhance, ImageOps
import streamlit.components.v1 as components
from openpyxl import Workbook, load_workbook as openpyxl_load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from dispatch_core import (
    SHIFT_COLUMNS,
    available_sources,
    build_result,
    calculate_totals,
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


CURRENT_STATUS_COLUMNS = ("2.0 現況", "2.0E 現況")
RECOGNITION_ERROR_TEXT = "識別錯誤"


def normalize_current_status(value) -> int | None:
    """將現況值正規化；空白與無法轉換的內容保留為 None，不再自動變成 0。"""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    if not text or text == RECOGNITION_ERROR_TEXT:
        return None

    try:
        return max(0, int(float(text)))
    except (TypeError, ValueError, OverflowError):
        return None


def blank_current_status(status_df: pd.DataFrame) -> pd.DataFrame:
    """建立現況預設為空白的資料表。"""
    blank_df = status_df.copy()
    for column in CURRENT_STATUS_COLUMNS:
        blank_df[column] = pd.Series([pd.NA] * len(blank_df), dtype="Int64")
    return blank_df



def coerce_nullable_current_status(status_df: pd.DataFrame) -> pd.DataFrame:
    """讓現況欄維持可輸入空白的 nullable integer 型別；使用向量化避免逐格 Python 迴圈。"""
    normalized_df = status_df.copy()
    for column in CURRENT_STATUS_COLUMNS:
        if column not in normalized_df.columns:
            normalized_df[column] = pd.Series(pd.NA, index=normalized_df.index, dtype="Int64")
            continue
        numeric = pd.to_numeric(
            normalized_df[column].replace(RECOGNITION_ERROR_TEXT, pd.NA),
            errors="coerce",
        )
        normalized_df[column] = pd.Series(
            np.trunc(numeric).clip(lower=0),
            index=normalized_df.index,
        ).astype("Int64")
    return normalized_df


def format_dispatch_status(current_value, standard_value) -> str:
    """現況為空白時顯示識別錯誤，其餘依標準值計算缺車或多車。"""
    current = normalize_current_status(current_value)
    if current is None:
        return RECOGNITION_ERROR_TEXT

    standard = safe_nonnegative_int(standard_value)
    difference = current - standard
    if difference > 0:
        return f"多 {difference} 台"
    if difference < 0:
        return f"缺 {abs(difference)} 台"
    return "符合"



def build_result_with_recognition_errors(status_df: pd.DataFrame) -> pd.DataFrame:
    """建立分析結果；以向量化計算取代逐列 iterrows。"""
    result = status_df[["行政區", "場站名稱"]].astype(str).copy()

    def build_status_series(current_column: str, standard_column: str) -> pd.Series:
        current = pd.Series(
            np.trunc(pd.to_numeric(status_df[current_column], errors="coerce")),
            index=status_df.index,
        ).clip(lower=0)
        standard = pd.Series(
            np.trunc(pd.to_numeric(status_df[standard_column], errors="coerce")),
            index=status_df.index,
        ).fillna(0).clip(lower=0)
        difference = current - standard
        output = pd.Series(RECOGNITION_ERROR_TEXT, index=status_df.index, dtype="object")
        valid = current.notna()
        output.loc[valid & difference.eq(0)] = "符合"
        output.loc[valid & difference.gt(0)] = (
            "多 " + difference.loc[valid & difference.gt(0)].astype(int).astype(str) + " 台"
        )
        output.loc[valid & difference.lt(0)] = (
            "缺 " + difference.loc[valid & difference.lt(0)].abs().astype(int).astype(str) + " 台"
        )
        return output

    result["2.0 缺／多幾台"] = build_status_series("2.0 現況", "2.0 標準")
    result["2.0E 缺／多幾台"] = build_status_series("2.0E 現況", "2.0E 標準")
    return result



def calculate_totals_ignoring_missing(
    status_df: pd.DataFrame,
    current_column: str,
    standard_column: str,
) -> tuple[int, int]:
    """缺／多合計只計算有效現況；使用向量化加速。"""
    current = pd.Series(
        np.trunc(pd.to_numeric(status_df[current_column], errors="coerce")),
        index=status_df.index,
    ).clip(lower=0)
    standard = pd.Series(
        np.trunc(pd.to_numeric(status_df[standard_column], errors="coerce")),
        index=status_df.index,
    ).fillna(0).clip(lower=0)
    difference = (current - standard).dropna()
    short_total = int((-difference[difference.lt(0)]).sum())
    extra_total = int(difference[difference.gt(0)].sum())
    return short_total, extra_total




def calculate_inventory_summary(
    status_df: pd.DataFrame,
    current_column: str,
    standard_column: str,
) -> dict[str, int | str | None]:
    """計算配置總數、目前總數與完整資料下的整體差額。"""
    standard = pd.Series(
        np.trunc(pd.to_numeric(status_df[standard_column], errors="coerce")),
        index=status_df.index,
    ).fillna(0).clip(lower=0)
    current = pd.Series(
        np.trunc(pd.to_numeric(status_df[current_column], errors="coerce")),
        index=status_df.index,
    ).clip(lower=0)

    configured_total = int(standard.sum())
    current_total = int(current.dropna().sum())
    missing_count = int(current.isna().sum())
    station_count = int(len(status_df))

    # 只在所有場站都有現況資料時顯示整體缺／多，避免把空白誤判成缺車。
    difference: int | None = None
    state = "pending"
    state_label = "資料未完整"
    difference_text = f"待補 {missing_count} 筆"
    signed_difference_text = "—"

    if missing_count == 0:
        difference = current_total - configured_total
        if difference > 0:
            state = "extra"
            state_label = "多車"
            difference_text = f"多 {difference} 台"
            signed_difference_text = f"+{difference} 台"
        elif difference < 0:
            state = "short"
            state_label = "缺車"
            difference_text = f"缺 {abs(difference)} 台"
            signed_difference_text = f"−{abs(difference)} 台"
        else:
            state = "balanced"
            state_label = "符合配置"
            difference_text = "剛好 0 台"
            signed_difference_text = "0 台"

    return {
        "configured_total": configured_total,
        "current_total": current_total,
        "missing_count": missing_count,
        "station_count": station_count,
        "difference": difference,
        "state": state,
        "state_label": state_label,
        "difference_text": difference_text,
        "signed_difference_text": signed_difference_text,
    }


def _summary_illustration(vehicle_type: str) -> str:
    """回傳不需額外圖片檔的內嵌 SVG 小插畫。"""
    if vehicle_type == "bike":
        return """
        <svg class="fleet-card-svg" viewBox="0 0 150 120" aria-hidden="true">
          <circle cx="49" cy="79" r="23" fill="none" stroke="currentColor" stroke-width="7"/>
          <circle cx="111" cy="79" r="23" fill="none" stroke="currentColor" stroke-width="7"/>
          <path d="M49 79 L67 45 L86 79 L49 79 M67 45 L94 45 L111 79 M65 45 L58 31 M53 31 H70 M85 79 L101 57"
                fill="none" stroke="currentColor" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M91 39 H111 L116 52 H96 Z" fill="currentColor" opacity=".28"/>
          <path d="M18 105 H132" stroke="currentColor" stroke-width="5" stroke-linecap="round" opacity=".28"/>
          <circle cx="123" cy="24" r="11" fill="currentColor" opacity=".18"/>
        </svg>
        """
    return """
    <svg class="fleet-card-svg" viewBox="0 0 150 120" aria-hidden="true">
      <circle cx="75" cy="51" r="39" fill="currentColor" opacity=".17"/>
      <path d="M83 17 L53 60 H72 L62 93 L101 45 H80 Z" fill="currentColor"/>
      <path d="M18 105 H132" stroke="currentColor" stroke-width="5" stroke-linecap="round" opacity=".25"/>
      <rect x="22" y="72" width="19" height="33" rx="3" fill="currentColor" opacity=".18"/>
      <rect x="45" y="63" width="17" height="42" rx="3" fill="currentColor" opacity=".22"/>
      <rect x="108" y="69" width="20" height="36" rx="3" fill="currentColor" opacity=".16"/>
    </svg>
    """


def render_inventory_summary_card(
    title: str,
    vehicle_type: str,
    summary: dict[str, int | str | None],
) -> None:
    """渲染配置／現況／差額三欄總覽卡。"""
    state = str(summary["state"])
    illustration = _summary_illustration(vehicle_type)
    station_count = safe_nonnegative_int(summary["station_count"])
    configured_total = safe_nonnegative_int(summary["configured_total"])
    current_total = safe_nonnegative_int(summary["current_total"])
    state_label = html.escape(str(summary["state_label"]))
    difference_text = html.escape(str(summary["difference_text"]))
    signed_difference_text = html.escape(str(summary["signed_difference_text"]))

    st.markdown(
        f"""
        <section class="fleet-summary-card fleet-theme-{vehicle_type}">
          <div class="fleet-card-illustration">{illustration}</div>
          <div class="fleet-card-content">
            <div class="fleet-card-heading">
              <div>
                <div class="fleet-card-title">{html.escape(title)}</div>
                <div class="fleet-card-subtitle">目前篩選共 {station_count} 個場站</div>
              </div>
              <span class="fleet-state-badge fleet-state-{state}">{state_label}</span>
            </div>
            <div class="fleet-card-metrics">
              <div class="fleet-metric-block">
                <div class="fleet-metric-label">配置總數</div>
                <div class="fleet-metric-value">{configured_total}<span>台</span></div>
              </div>
              <div class="fleet-metric-block">
                <div class="fleet-metric-label">目前總數</div>
                <div class="fleet-metric-value">{current_total}<span>台</span></div>
              </div>
              <div class="fleet-metric-block fleet-difference-block">
                <div class="fleet-metric-label">差額</div>
                <div class="fleet-difference-chip fleet-difference-{state}">
                  <strong>{difference_text}</strong>
                  <small>{signed_difference_text}</small>
                </div>
              </div>
            </div>
          </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_missing_data_notice(missing_bike_count: int, missing_ebike_count: int) -> None:
    """渲染資料不完整提示；有空白時暫停顯示整體缺／多差額。"""
    st.markdown(
        f"""
        <div class="fleet-data-notice">
          <div class="fleet-notice-icon" aria-hidden="true">!</div>
          <div class="fleet-notice-copy">
            <div class="fleet-notice-title">資料提醒</div>
            <div class="fleet-notice-text">
              尚有識別錯誤／空白資料：2.0 共 <strong>{missing_bike_count}</strong> 筆，
              2.0E 共 <strong>{missing_ebike_count}</strong> 筆。<br>
              配置總數與目前已取得的總數仍會顯示，但整體差額會等資料完整後再計算。
            </div>
          </div>
          <div class="fleet-notice-decoration" aria-hidden="true">📋</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_dispatch_legend() -> None:
    st.markdown(
        """
        <div class="fleet-legend" aria-label="狀態圖例">
          <span><i class="fleet-legend-dot fleet-legend-extra"></i>多車</span>
          <span class="fleet-legend-divider">｜</span>
          <span><i class="fleet-legend-dot fleet-legend-short"></i>缺車</span>
          <span class="fleet-legend-divider">｜</span>
          <span><i class="fleet-legend-dot fleet-legend-balanced"></i>符合</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


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
    """依使用者選擇排序；以向量化字串擷取取代逐格正規表示式。"""
    if result_df.empty:
        return result_df

    sorted_df = result_df.copy()
    bike_counts = pd.to_numeric(
        sorted_df["2.0 缺／多幾台"].astype(str).str.extract(r"(\d+)", expand=False),
        errors="coerce",
    )
    ebike_counts = pd.to_numeric(
        sorted_df["2.0E 缺／多幾台"].astype(str).str.extract(r"(\d+)", expand=False),
        errors="coerce",
    )
    bike_valid = sorted_df["2.0 缺／多幾台"].astype(str).str.contains("多|缺|少", regex=True)
    ebike_valid = sorted_df["2.0E 缺／多幾台"].astype(str).str.contains("多|缺|少", regex=True)
    bike_counts = bike_counts.where(bike_valid)
    ebike_counts = ebike_counts.where(ebike_valid)

    count_frame = pd.concat([bike_counts, ebike_counts], axis=1)
    sorted_df["_排序最大台數"] = count_frame.max(axis=1, skipna=True).fillna(-1).astype(int)
    sorted_df["_排序總台數"] = count_frame.fillna(0).sum(axis=1).astype(int)
    sorted_df["_排序2.0台數"] = bike_counts.fillna(-1).astype(int)
    sorted_df["_排序2.0E台數"] = ebike_counts.fillna(-1).astype(int)
    sorted_df["_原始順序"] = np.arange(len(sorted_df), dtype=int)

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

    sorted_df = sorted_df.sort_values(by=by, ascending=ascending_values, kind="mergesort")
    return sorted_df.drop(
        columns=[
            "_排序最大台數", "_排序總台數", "_排序2.0台數",
            "_排序2.0E台數", "_原始順序",
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


@st.cache_data(show_spinner=False, max_entries=32)
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
    yellow_fill = PatternFill("solid", fgColor="FFF2CC")

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
            elif RECOGNITION_ERROR_TEXT in text:
                cell.fill = yellow_fill
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



@lru_cache(maxsize=8192)
def _normalize_station_key_cached(text: str) -> str:
    normalized = normalize_ocr_text(text).lower()
    return re.sub(r"[^0-9a-z\u3400-\u9fff]", "", normalized)


def normalize_station_key(value) -> str:
    """建立只保留中英數字的場站比對鍵，並快取重複字串。"""
    return _normalize_station_key_cached(str(value or ""))




YOUBIKE_STATION_CATALOG_URL = "https://apis.youbike.com.tw/json/station-min-yb2.json"
YOUBIKE_PARKING_INFO_URL = "https://apis.youbike.com.tw/tw2/parkingInfo"
YOUBIKE_REQUEST_TIMEOUT_SECONDS = 25
YOUBIKE_HTTP_MAX_ATTEMPTS = 3
YOUBIKE_STATION_BATCH_SIZE = 100
YOUBIKE_MATCH_THRESHOLD = 0.82
TAIPEI_TIMEZONE = ZoneInfo("Asia/Taipei")


class YouBikeDataError(RuntimeError):
    """YouBike 官網公開資料連線或格式異常。"""


YOUBIKE_BROWSER_COMPONENT_HTML = r"""<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    :root { color-scheme: light dark; }
    * { box-sizing: border-box; }
    html, body {
      width: 1px; height: 1px; margin: 0; padding: 0; overflow: hidden;
      background: transparent; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }
    /* 同步元件留在背景執行；手動操作改由主頁右側懸浮按鈕觸發。 */
    #syncButton, #status { display: none !important; }
    .error { color: #c62828 !important; }
  </style>
</head>
<body>
  <button id="syncButton" type="button">🔄 由手機／瀏覽器取得 YouBike 即時車數</button>
  <div id="status"></div>
<script>
(() => {
  const API_VERSION = 1;
  const button = document.getElementById("syncButton");
  const statusNode = document.getElementById("status");
  let args = {};
  let busy = false;
  let autoTimer = null;

  function send(type, data = {}) {
    window.parent.postMessage({ isStreamlitMessage: true, type, ...data }, "*");
  }
  function sendHostSyncState(state, detail = {}) {
    window.parent.postMessage({
      source: "ubike-browser-sync",
      type: "ubike:sync-state",
      state,
      ...detail,
    }, "*");
  }
  function setHeight() {
    send("streamlit:setFrameHeight", { height: 1 });
  }
  function setValue(value) {
    send("streamlit:setComponentValue", { value, dataType: "json" });
  }
  function setStatus(text, isError = false) {
    statusNode.textContent = text || "";
    statusNode.className = isError ? "error" : "";
    setHeight();
  }
  function intOrNull(value) {
    if (value === null || value === undefined || value === "") return null;
    const number = Number(value);
    return Number.isFinite(number) ? Math.max(0, Math.trunc(number)) : null;
  }
  function firstNonempty(...values) {
    for (const value of values) {
      if (value === null || value === undefined) continue;
      if (typeof value === "string" && !value.trim()) continue;
      return value;
    }
    return null;
  }
  function extractItems(payload) {
    if (Array.isArray(payload)) return payload.filter(item => item && typeof item === "object");
    if (!payload || typeof payload !== "object") return [];
    const candidates = [payload.data, payload.result, payload.stations, payload.retVal];
    for (const candidate of candidates) {
      if (Array.isArray(candidate)) return candidate.filter(item => item && typeof item === "object");
      if (candidate && typeof candidate === "object" && Array.isArray(candidate.data)) {
        return candidate.data.filter(item => item && typeof item === "object");
      }
    }
    return [];
  }
  function isTaitung(item) {
    const locationText = [
      item.county_tw, item.city_tw, item.scity, item.district_tw,
      item.address_tw, item.name_tw, item.sarea, item.ar, item.sna
    ].map(value => String(value || "")).join(" ").replaceAll("臺", "台");
    if (locationText.includes("台東縣")) return true;
    const lat = Number(firstNonempty(item.lat, item.latitude));
    const lng = Number(firstNonempty(item.lng, item.longitude));
    return Number.isFinite(lat) && Number.isFinite(lng)
      && lat >= 21.85 && lat <= 23.60 && lng >= 120.70 && lng <= 122.20;
  }
  function sleep(milliseconds) {
    return new Promise(resolve => setTimeout(resolve, Math.max(0, milliseconds)));
  }
  async function fetchJson(url, options = {}, maxAttempts = 3) {
    let lastError = null;
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 25000);
      try {
        const response = await fetch(url, {
          cache: "no-store",
          credentials: "omit",
          ...options,
          signal: controller.signal,
        });
        if (!response.ok) {
          const error = new Error(`HTTP ${response.status}`);
          error.status = response.status;
          throw error;
        }
        const responseText = await response.text();
        try { return JSON.parse(responseText); }
        catch (_) { throw new Error("官網回傳的內容不是 JSON"); }
      } catch (error) {
        lastError = error;
        const status = Number(error && error.status);
        const retryable = error && (
          error.name === "AbortError" || status === 429 || status >= 500 || !Number.isFinite(status)
        );
        if (!retryable || attempt >= maxAttempts) throw error;
        // 加入少量隨機退避，避免多個並行請求在同一時間再次撞上官網限制。
        await sleep(180 * attempt + Math.floor(Math.random() * 140));
      } finally {
        clearTimeout(timeout);
      }
    }
    throw lastError || new Error("官網請求失敗");
  }
  function batched(values, size) {
    const output = [];
    for (let index = 0; index < values.length; index += size) output.push(values.slice(index, index + size));
    return output;
  }
  function taipeiTimeText() {
    return new Intl.DateTimeFormat("zh-TW", {
      timeZone: "Asia/Taipei", year: "numeric", month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false,
    }).format(new Date()).replaceAll("-", "/");
  }
  function eventId() {
    if (globalThis.crypto && typeof globalThis.crypto.randomUUID === "function") return globalThis.crypto.randomUUID();
    return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function scheduleAutoSync() {
    if (autoTimer !== null) {
      clearTimeout(autoTimer);
      autoTimer = null;
    }
    if (!args.auto_refresh) return;
    const seconds = Math.max(5, Math.min(60, Number(args.auto_refresh_seconds || 60)));
    autoTimer = setTimeout(() => {
      autoTimer = null;
      if (busy) scheduleAutoSync();
      else runSync();
    }, seconds * 1000);
  }

  async function runSync() {
    if (busy) return;
    if (autoTimer !== null) {
      clearTimeout(autoTimer);
      autoTimer = null;
    }
    busy = true;
    const startedAt = performance.now();
    sendHostSyncState("busy");
    button.disabled = true;
    button.textContent = "⏳ 正在高速分批讀取 YouBike 官網……";
    setStatus("將以最大批次、有限並行及只補漏站的方式取得資料，不經 TDX。", false);

    try {
      const catalogUrl = args.catalog_url || "https://apis.youbike.com.tw/json/station-min-yb2.json";
      const parkingUrl = args.parking_url || "https://apis.youbike.com.tw/tw2/parkingInfo";
      const batchSize = Math.max(1, Math.min(50, Number(args.batch_size || 20)));
      const concurrency = Math.max(1, Math.min(6, Number(args.request_concurrency || 4)));
      const maxBatchRounds = Math.max(1, Math.min(8, Number(args.max_batch_rounds || 4)));
      const maxSingleRounds = Math.max(0, Math.min(4, Number(args.max_single_rounds || 2)));
      const waveDelayMs = Math.max(0, Math.min(1000, Number(args.wave_delay_ms || 70)));

      const catalogPayload = await fetchJson(catalogUrl, {
        method: "GET",
        headers: { "Accept": "application/json, text/plain, */*" },
      });
      const catalog = extractItems(catalogPayload).filter(isTaitung).map(item => {
        const stationId = String(firstNonempty(item.station_no, item.sno, item.station_id) || "").trim();
        const stationName = String(firstNonempty(item.name_tw, item.sna, item.station_name) || "").trim();
        return {
          station_uid: stationId,
          station_id: stationId,
          station_name: stationName,
          service_status: intOrNull(firstNonempty(item.status, item.act, 1)) ?? 1,
          source_update_time: String(firstNonempty(item.updated_at, item.mday, item.time) || "").trim(),
          latitude: firstNonempty(item.lat, item.latitude),
          longitude: firstNonempty(item.lng, item.longitude),
        };
      }).filter(item => item.station_id && item.station_name);

      if (!catalog.length) throw new Error("官網站點清單中找不到臺東候選場站");

      const requestedStationIds = [...new Set(catalog.map(item => item.station_id))];
      const requestedStationIdSet = new Set(requestedStationIds);
      const parkingMap = new Map();
      let requestCount = 0;
      let failedRequestCount = 0;
      let batchRoundCount = 0;
      let singleRoundCount = 0;

      function stationIdOf(item) {
        return String(firstNonempty(item && item.station_no, item && item.sno) || "").trim();
      }

      function mergeParkingItems(items) {
        let addedCount = 0;
        for (const item of items) {
          const stationId = stationIdOf(item);
          if (!stationId || !requestedStationIdSet.has(stationId)) continue;
          if (!parkingMap.has(stationId)) addedCount += 1;
          parkingMap.set(stationId, item);
        }
        return addedCount;
      }

      function currentMissingIds() {
        return requestedStationIds.filter(stationId => !parkingMap.has(stationId));
      }

      async function requestParkingGroup(stationIds) {
        const payload = await fetchJson(parkingUrl, {
          method: "POST",
          headers: {
            "Accept": "application/json, text/plain, */*",
            "Content-Type": "application/json;charset=UTF-8",
          },
          body: JSON.stringify({ station_no: stationIds }),
        });
        return extractItems(payload);
      }

      async function runGroups(groups, phaseText, workerLimit = concurrency) {
        if (!groups.length) return { addedCount: 0, failedGroups: [] };
        let nextIndex = 0;
        let addedCount = 0;
        let completedCount = 0;
        const failedGroups = [];

        async function worker() {
          while (true) {
            const index = nextIndex;
            nextIndex += 1;
            if (index >= groups.length) return;
            const stationIds = groups[index];
            try {
              requestCount += 1;
              const items = await requestParkingGroup(stationIds);
              addedCount += mergeParkingItems(items);
            } catch (error) {
              failedRequestCount += 1;
              failedGroups.push({ stationIds, error: String(error && error.message ? error.message : error) });
            } finally {
              completedCount += 1;
              const missingCount = currentMissingIds().length;
              setStatus(
                `${phaseText}：完成 ${completedCount}／${groups.length} 批，已取得 ` +
                `${parkingMap.size}／${requestedStationIds.length} 站，尚缺 ${missingCount} 站`,
                false,
              );
            }
          }
        }

        const workerCount = Math.min(Math.max(1, workerLimit), groups.length);
        await Promise.all(Array.from({ length: workerCount }, () => worker()));
        return { addedCount, failedGroups };
      }

      let missingStationIds = currentMissingIds();
      let previousMissingCount = missingStationIds.length + 1;

      // 主階段：每一輪都使用設定的最大批次，並行查完後只保留仍缺少的場站進入下一輪。
      for (let round = 1; round <= maxBatchRounds && missingStationIds.length; round += 1) {
        batchRoundCount = round;
        const groups = batched(missingStationIds, batchSize);
        setStatus(
          `高速批次第 ${round} 輪：${missingStationIds.length} 站，分成 ${groups.length} 批並行讀取……`,
          false,
        );
        const result = await runGroups(groups, `高速批次第 ${round} 輪`);
        missingStationIds = currentMissingIds();

        if (!missingStationIds.length) break;
        // 這一輪完全沒有新增資料時，繼續重送相同批次沒有速度效益，立即改走單站補查。
        if (result.addedCount <= 0 || missingStationIds.length >= previousMissingCount) break;
        previousMissingCount = missingStationIds.length;
        if (waveDelayMs) await sleep(waveDelayMs);
      }

      // 最後階段：只對殘留漏站做單站並行查詢，避免一個異常站拖累同批其他場站。
      missingStationIds = currentMissingIds();
      for (let round = 1; round <= maxSingleRounds && missingStationIds.length; round += 1) {
        singleRoundCount = round;
        const singleGroups = missingStationIds.map(stationId => [stationId]);
        setStatus(
          `單站補查第 ${round} 輪：正在並行補齊最後 ${missingStationIds.length} 個場站……`,
          false,
        );
        const beforeCount = parkingMap.size;
        await runGroups(singleGroups, `單站補查第 ${round} 輪`, Math.min(6, concurrency + 1));
        missingStationIds = currentMissingIds();
        if (!missingStationIds.length) break;
        // 即使這一輪暫時沒有新增，也保留後續重試機會；官網可能只是短暫漏回或限流。
        const noProgressDelay = parkingMap.size <= beforeCount ? 260 * round : waveDelayMs + 60;
        if (noProgressDelay) await sleep(noProgressDelay);
      }

      const sourceTimes = [];
      const records = [];
      for (const station of catalog) {
        const parking = parkingMap.get(station.station_id);
        if (!parking) continue;
        let detail = firstNonempty(parking.available_spaces_detail, parking.sbi_detail);
        if (!detail || typeof detail !== "object") detail = {};
        const sourceTime = String(firstNonempty(
          parking.updated_at, parking.mday, parking.time, station.source_update_time
        ) || "").trim();
        if (sourceTime) sourceTimes.push(sourceTime);
        records.push({
          ...station,
          service_status: intOrNull(firstNonempty(parking.status, parking.act, station.service_status, 1)) ?? 1,
          general_bikes: intOrNull(detail.yb2),
          electric_bikes: intOrNull(detail.eyb),
          available_spaces: intOrNull(firstNonempty(parking.available_spaces, parking.sbi)),
          empty_spaces: intOrNull(firstNonempty(parking.empty_spaces, parking.bemp)),
          parking_spaces: intOrNull(firstNonempty(parking.parking_spaces, parking.tot)),
          source_update_time: sourceTime,
        });
      }

      if (!records.length) throw new Error("官網沒有回傳臺東場站即時車數");
      missingStationIds = currentMissingIds();
      setValue({
        ok: true,
        event_id: eventId(),
        records,
        fetched_at: taipeiTimeText(),
        latest_source_time: sourceTimes.length ? sourceTimes.sort().at(-1) : "",
        station_count: records.length,
        requested_station_count: requestedStationIds.length,
        missing_station_count: missingStationIds.length,
        missing_station_ids: missingStationIds,
        request_batch_count: requestCount,
        request_count: requestCount,
        failed_request_count: failedRequestCount,
        batch_round_count: batchRoundCount,
        single_round_count: singleRoundCount,
        batch_size: batchSize,
        request_concurrency: concurrency,
        elapsed_ms: Math.max(0, Math.round(performance.now() - startedAt)),
        source: "YouBike 官網公開接口（高速循環補查，由使用者瀏覽器直接取得，免 TDX）",
      });
      const missingText = missingStationIds.length ? `，仍缺 ${missingStationIds.length} 個` : "，已全數取得";
      setStatus(
        `已取得 ${records.length}／${requestedStationIds.length} 個場站${missingText}，共送出 ${requestCount} 次請求，正在寫入分析系統……`,
        false,
      );
      sendHostSyncState("success", { station_count: records.length });
    } catch (error) {
      const message = error && error.name === "AbortError"
        ? "連線逾時，請檢查手機網路後再試"
        : String(error && error.message ? error.message : error);
      setValue({ ok: false, event_id: eventId(), error: message });
      setStatus(`同步失敗：${message}`, true);
      sendHostSyncState("error", { message });
    } finally {
      busy = false;
      button.disabled = false;
      button.textContent = args.button_label || "🔄 由手機／瀏覽器取得 YouBike 即時車數";
      setHeight();
      scheduleAutoSync();
    }
  }

  button.addEventListener("click", runSync);
  window.addEventListener("message", event => {
    if (!event.data) return;
    if (event.data.type === "ubike:manual-sync") {
      runSync();
      return;
    }
    if (event.data.type !== "streamlit:render") return;
    args = event.data.args || {};
    button.textContent = args.button_label || "🔄 手動更新即時車數";
    button.disabled = Boolean(event.data.disabled) || busy;
    scheduleAutoSync();
    setHeight();
  });

  send("streamlit:componentReady", { apiVersion: API_VERSION });
  setHeight();
})();
</script>
</body>
</html>
"""


_YOUBIKE_BROWSER_SYNC_COMPONENT = None


def get_youbike_browser_sync_component():
    """建立雙向 Streamlit 元件，讓請求從使用者瀏覽器發出以避開雲端主機 503。"""
    global _YOUBIKE_BROWSER_SYNC_COMPONENT
    if _YOUBIKE_BROWSER_SYNC_COMPONENT is not None:
        return _YOUBIKE_BROWSER_SYNC_COMPONENT

    component_dir = Path(tempfile.gettempdir()) / "youbike_browser_sync_component_v2"
    component_dir.mkdir(parents=True, exist_ok=True)
    index_path = component_dir / "index.html"
    try:
        if not index_path.exists() or index_path.read_text(encoding="utf-8") != YOUBIKE_BROWSER_COMPONENT_HTML:
            index_path.write_text(YOUBIKE_BROWSER_COMPONENT_HTML, encoding="utf-8")
    except OSError as exc:
        raise YouBikeDataError(f"無法建立瀏覽器同步元件：{exc}") from exc

    _YOUBIKE_BROWSER_SYNC_COMPONENT = components.declare_component(
        "youbike_browser_sync_v2",
        path=str(component_dir),
    )
    return _YOUBIKE_BROWSER_SYNC_COMPONENT


def normalize_browser_live_payload(payload) -> dict:
    """驗證瀏覽器回傳資料並補齊 Python 端配對所需欄位。"""
    if not isinstance(payload, dict):
        raise YouBikeDataError("瀏覽器沒有回傳有效資料。")
    if not payload.get("ok"):
        raise YouBikeDataError(str(payload.get("error") or "瀏覽器同步失敗。"))

    raw_records = payload.get("records")
    if not isinstance(raw_records, list):
        raise YouBikeDataError("瀏覽器回傳的場站資料格式不正確。")

    records: list[dict] = []
    for raw_record in raw_records:
        if not isinstance(raw_record, dict):
            continue
        station_id = str(raw_record.get("station_id") or raw_record.get("station_uid") or "").strip()
        station_name = str(raw_record.get("station_name") or "").strip()
        if not station_id or not station_name:
            continue
        records.append(
            {
                **raw_record,
                "station_uid": station_id,
                "station_id": station_id,
                "station_name": station_name,
                "station_key": normalize_youbike_station_key(station_name),
                "service_status": safe_nonnegative_int(raw_record.get("service_status", 1)),
                "general_bikes": normalize_current_status(raw_record.get("general_bikes")),
                "electric_bikes": normalize_current_status(raw_record.get("electric_bikes")),
                "available_spaces": normalize_current_status(raw_record.get("available_spaces")),
                "empty_spaces": normalize_current_status(raw_record.get("empty_spaces")),
                "parking_spaces": normalize_current_status(raw_record.get("parking_spaces")),
            }
        )

    if not records:
        raise YouBikeDataError("瀏覽器沒有回傳可用的臺東場站即時車數。")

    return {
        "records": records,
        "fetched_at": str(payload.get("fetched_at") or datetime.now(TAIPEI_TIMEZONE).strftime("%Y/%m/%d %H:%M:%S")),
        "latest_source_time": str(payload.get("latest_source_time") or "").strip(),
        "station_count": len(records),
        "requested_station_count": safe_nonnegative_int(payload.get("requested_station_count")),
        "missing_station_count": safe_nonnegative_int(payload.get("missing_station_count")),
        "request_batch_count": safe_nonnegative_int(payload.get("request_batch_count")),
        "request_count": safe_nonnegative_int(payload.get("request_count")),
        "failed_request_count": safe_nonnegative_int(payload.get("failed_request_count")),
        "batch_round_count": safe_nonnegative_int(payload.get("batch_round_count")),
        "single_round_count": safe_nonnegative_int(payload.get("single_round_count")),
        "batch_size": safe_nonnegative_int(payload.get("batch_size")),
        "request_concurrency": safe_nonnegative_int(payload.get("request_concurrency")),
        "elapsed_ms": safe_nonnegative_int(payload.get("elapsed_ms")),
        "missing_station_ids": [
            str(value).strip() for value in payload.get("missing_station_ids", [])
            if str(value).strip()
        ] if isinstance(payload.get("missing_station_ids"), list) else [],
        "source": str(payload.get("source") or "YouBike 官網公開接口（高速循環補查，由瀏覽器直接取得，免 TDX）"),
        "event_id": str(payload.get("event_id") or "").strip(),
    }


def _first_nonempty(*values):
    """回傳第一個不是 None／空字串的值。"""
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return value
    return None


def _youbike_http_json(
    url: str,
    *,
    method: str = "GET",
    json_body: dict | None = None,
):
    """讀取 YouBike 官網公開 JSON；免 TDX、免帳號、免 API 金鑰。"""
    request_headers = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.7",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/126.0 Safari/537.36"
        ),
        "Referer": "https://www.youbike.com.tw/region/taitung/stations/",
        "Origin": "https://www.youbike.com.tw",
    }

    encoded_body = None
    if json_body is not None:
        encoded_body = json.dumps(json_body, ensure_ascii=False).encode("utf-8")
        request_headers["Content-Type"] = "application/json;charset=UTF-8"

    last_error: Exception | None = None
    for attempt in range(1, YOUBIKE_HTTP_MAX_ATTEMPTS + 1):
        request = Request(
            url,
            data=encoded_body,
            headers=request_headers,
            method=method,
        )

        try:
            with urlopen(request, timeout=YOUBIKE_REQUEST_TIMEOUT_SECONDS) as response:
                raw_body = response.read().decode("utf-8-sig")

            try:
                payload = json.loads(raw_body)
            except json.JSONDecodeError as exc:
                raise YouBikeDataError("YouBike 官網回傳內容不是有效的 JSON。") from exc

            if isinstance(payload, dict):
                ret_code = payload.get("retCode")
                if ret_code not in (None, 1, "1", True):
                    ret_message = str(payload.get("retMsg") or "官方資料服務回傳失敗").strip()
                    raise YouBikeDataError(f"YouBike 官網資料服務錯誤：{ret_message}")
            return payload

        except HTTPError as exc:
            last_error = exc
            try:
                detail = exc.read().decode("utf-8", errors="replace")[:240]
            except Exception:
                detail = ""

            # 429 與 5xx 通常是暫時性錯誤，先短暫退避後自動重試。
            if exc.code == 429 or 500 <= exc.code <= 599:
                if attempt < YOUBIKE_HTTP_MAX_ATTEMPTS:
                    time.sleep(0.8 * attempt)
                    continue
                if exc.code == 429:
                    raise YouBikeDataError(
                        "官方資料請求過於頻繁，請等候約 1 分鐘再試。"
                    ) from exc

            if exc.code in (401, 403):
                raise YouBikeDataError(
                    "YouBike 官網暫時拒絕此主機連線；請稍後再試，照片辨識功能仍可使用。"
                ) from exc
            raise YouBikeDataError(
                f"YouBike 官網資料回傳 HTTP {exc.code}。{detail or '請稍後再試。'}"
            ) from exc
        except (URLError, TimeoutError) as exc:
            last_error = exc
            if attempt < YOUBIKE_HTTP_MAX_ATTEMPTS:
                time.sleep(0.8 * attempt)
                continue

    if isinstance(last_error, URLError):
        reason = getattr(last_error, "reason", last_error)
        raise YouBikeDataError(f"無法連線至 YouBike 官網資料服務：{reason}") from last_error
    if isinstance(last_error, TimeoutError):
        raise YouBikeDataError("連線 YouBike 官網資料服務逾時，請稍後重試。") from last_error
    raise YouBikeDataError("YouBike 官網資料服務暫時無法使用，請稍後重試。")


def _extract_youbike_station_items(payload) -> list[dict]:
    """相容清單陣列、data 包裝，以及 retVal.data 等常見官方格式。"""
    containers = [payload]
    if isinstance(payload, dict):
        containers.append(payload.get("retVal"))

    for container in containers:
        if isinstance(container, list):
            return [item for item in container if isinstance(item, dict)]
        if not isinstance(container, dict):
            continue
        for key in ("data", "result", "stations", "retVal"):
            values = container.get(key)
            if isinstance(values, list):
                return [item for item in values if isinstance(item, dict)]
            if isinstance(values, dict):
                nested_data = values.get("data")
                if isinstance(nested_data, list):
                    return [item for item in nested_data if isinstance(item, dict)]
    return []


def normalize_youbike_station_key(value) -> str:
    """正規化 Excel 與 YouBike 官網站名，處理臺／台與車種前綴差異。"""
    text = normalize_ocr_text(value).lower().replace("臺", "台")
    text = re.sub(
        r"^(?:youbike|ubike)\s*2\s*[.．]?\s*0\s*e?\s*[_\-－—:：]*\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )
    text = text.replace("公共自行車租賃站", "")
    return re.sub(r"[^0-9a-z\u3400-\u9fff]", "", text)


def _youbike_station_similarity(excel_name: str, api_name: str) -> float:
    excel_key = normalize_youbike_station_key(excel_name)
    api_key = normalize_youbike_station_key(api_name)
    if not excel_key or not api_key:
        return 0.0
    if excel_key == api_key:
        return 1.0
    if min(len(excel_key), len(api_key)) >= 4 and (
        excel_key in api_key or api_key in excel_key
    ):
        return 0.96
    return SequenceMatcher(None, excel_key, api_key, autojunk=False).ratio()


def _looks_like_taitung_station(record: dict) -> bool:
    """以官方欄位及經緯度範圍篩出臺東縣候選場站。"""
    location_text = " ".join(
        str(record.get(key) or "")
        for key in (
            "county_tw", "city_tw", "scity", "district_tw", "address_tw",
            "name_tw", "sarea", "ar", "sna",
        )
    ).replace("臺", "台")
    if "台東縣" in location_text:
        return True

    try:
        latitude = float(_first_nonempty(record.get("lat"), record.get("latitude")))
        longitude = float(_first_nonempty(record.get("lng"), record.get("longitude")))
    except (TypeError, ValueError):
        return False

    # 包含臺東本島、綠島與蘭嶼。後續仍會以場站名稱做一對一安全配對。
    return 21.85 <= latitude <= 23.60 and 120.70 <= longitude <= 122.20


@st.cache_data(show_spinner=False, ttl=21600, max_entries=4)
def fetch_youbike_taitung_station_catalog() -> list[dict]:
    """取得 YouBike 全臺站點清單並留下臺東縣候選場站。"""
    payload = _youbike_http_json(YOUBIKE_STATION_CATALOG_URL)
    items = _extract_youbike_station_items(payload)
    records: list[dict] = []

    for item in items:
        if not _looks_like_taitung_station(item):
            continue
        station_no = str(
            _first_nonempty(item.get("station_no"), item.get("sno"), item.get("station_id"))
            or ""
        ).strip()
        station_name = str(
            _first_nonempty(item.get("name_tw"), item.get("sna"), item.get("station_name"))
            or ""
        ).strip()
        if not station_no or not station_name:
            continue

        raw_status = _first_nonempty(item.get("status"), item.get("act"), 1)
        records.append(
            {
                "station_uid": station_no,
                "station_id": station_no,
                "station_name": station_name,
                "station_key": normalize_youbike_station_key(station_name),
                "service_status": safe_nonnegative_int(raw_status),
                "source_update_time": str(
                    _first_nonempty(item.get("updated_at"), item.get("mday"), item.get("time"))
                    or ""
                ).strip(),
                "latitude": _first_nonempty(item.get("lat"), item.get("latitude")),
                "longitude": _first_nonempty(item.get("lng"), item.get("longitude")),
            }
        )

    if not records:
        raise YouBikeDataError("YouBike 官網沒有回傳可辨識的臺東場站清單。")
    return records


def _batched_station_numbers(station_numbers: list[str]):
    """將場站編號分批，避免單次 POST 過大而被官方服務拒絕。"""
    for start_index in range(0, len(station_numbers), YOUBIKE_STATION_BATCH_SIZE):
        yield station_numbers[start_index : start_index + YOUBIKE_STATION_BATCH_SIZE]


@st.cache_data(show_spinner=False, ttl=60, max_entries=8)
def fetch_youbike_taitung_bike_data(refresh_bucket: int) -> dict:
    """取得臺東縣 YouBike 2.0／2.0E 即時可借車數；完全不使用 TDX。"""
    del refresh_bucket  # 讓快取依傳入分鐘批次更新，同分鐘內避免重複打官方接口。
    catalog = fetch_youbike_taitung_station_catalog()
    station_numbers = [record["station_id"] for record in catalog]

    parking_items: list[dict] = []
    station_batches = list(_batched_station_numbers(station_numbers))
    for batch_index, station_batch in enumerate(station_batches):
        payload = _youbike_http_json(
            YOUBIKE_PARKING_INFO_URL,
            method="POST",
            json_body={"station_no": station_batch},
        )
        parking_items.extend(_extract_youbike_station_items(payload))
        if batch_index < len(station_batches) - 1:
            time.sleep(0.15)

    parking_by_station = {
        str(_first_nonempty(item.get("station_no"), item.get("sno")) or "").strip(): item
        for item in parking_items
        if str(_first_nonempty(item.get("station_no"), item.get("sno")) or "").strip()
    }

    records: list[dict] = []
    source_times: list[str] = []
    for station in catalog:
        parking = parking_by_station.get(station["station_id"])
        if not isinstance(parking, dict):
            continue

        detail = _first_nonempty(
            parking.get("available_spaces_detail"),
            parking.get("sbi_detail"),
        )
        if not isinstance(detail, dict):
            detail = {}

        general_bikes = normalize_current_status(detail.get("yb2"))
        electric_bikes = normalize_current_status(detail.get("eyb"))
        source_update_time = str(
            _first_nonempty(
                parking.get("updated_at"),
                parking.get("mday"),
                parking.get("time"),
                station.get("source_update_time"),
            )
            or ""
        ).strip()
        if source_update_time:
            source_times.append(source_update_time)

        raw_service_status = _first_nonempty(
            parking.get("status"),
            parking.get("act"),
            station.get("service_status"),
            1,
        )
        records.append(
            {
                **station,
                "service_status": safe_nonnegative_int(raw_service_status),
                "general_bikes": general_bikes,
                "electric_bikes": electric_bikes,
                "available_spaces": normalize_current_status(
                    _first_nonempty(parking.get("available_spaces"), parking.get("sbi"))
                ),
                "empty_spaces": normalize_current_status(
                    _first_nonempty(parking.get("empty_spaces"), parking.get("bemp"))
                ),
                "parking_spaces": normalize_current_status(
                    _first_nonempty(parking.get("parking_spaces"), parking.get("tot"))
                ),
                "source_update_time": source_update_time,
            }
        )

    if not records:
        raise YouBikeDataError(
            "YouBike 官網沒有回傳臺東場站即時車數，可能是官方資料服務暫時異常。"
        )

    fetched_at = datetime.now(TAIPEI_TIMEZONE).strftime("%Y/%m/%d %H:%M:%S")
    return {
        "records": records,
        "fetched_at": fetched_at,
        "latest_source_time": max(source_times) if source_times else "",
        "station_count": len(records),
        "request_batch_count": len(station_batches),
        "source": "YouBike 官網公開接口（免 TDX）",
    }


def match_youbike_station(
    excel_name: str,
    live_records: list[dict],
) -> tuple[dict | None, float, bool]:
    """配對 Excel 與 YouBike 官網站名；過於接近時不自動寫入。"""
    excel_key = normalize_youbike_station_key(excel_name)
    if not excel_key:
        return None, 0.0, False

    exact_matches = [record for record in live_records if record.get("station_key") == excel_key]
    if len(exact_matches) == 1:
        return exact_matches[0], 1.0, False
    if len(exact_matches) > 1:
        return None, 1.0, True

    ranked = sorted(
        (
            (_youbike_station_similarity(excel_name, record.get("station_name", "")), record)
            for record in live_records
        ),
        key=lambda item: item[0],
        reverse=True,
    )
    if not ranked:
        return None, 0.0, False

    best_score, best_record = ranked[0]
    if best_score < YOUBIKE_MATCH_THRESHOLD:
        return None, best_score, False

    second_score = ranked[1][0] if len(ranked) > 1 else 0.0
    ambiguous = best_score < 0.96 and second_score >= best_score - 0.035
    if ambiguous:
        return None, best_score, True
    return best_record, best_score, False


def apply_youbike_updates_to_dataframe(
    base_df: pd.DataFrame,
    live_records: list[dict],
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """將 YouBike 官網即時車數寫入目前配置，只寫入安全且資料完整的配對。"""
    updated_df = base_df.copy()
    report_rows: list[dict] = []
    used_station_ids: set[str] = set()
    matched_count = 0
    skipped_count = 0
    unmatched_count = 0

    for row_index, row in updated_df.iterrows():
        excel_name = str(row.get("場站名稱", "") or "").strip()
        matched_record, score, ambiguous = match_youbike_station(excel_name, live_records)

        if matched_record is None:
            unmatched_count += 1
            report_rows.append(
                {
                    "Excel 場站": excel_name,
                    "YouBike 場站": "",
                    "2.0": pd.NA,
                    "2.0E": pd.NA,
                    "結果": "名稱可能重複，未寫入" if ambiguous else "找不到安全配對",
                    "相似度": round(score, 3),
                }
            )
            continue

        station_id = str(matched_record.get("station_id", "") or "")
        if station_id in used_station_ids:
            unmatched_count += 1
            report_rows.append(
                {
                    "Excel 場站": excel_name,
                    "YouBike 場站": matched_record.get("station_name", ""),
                    "2.0": pd.NA,
                    "2.0E": pd.NA,
                    "結果": "同一官網場站已配對，未重複寫入",
                    "相似度": round(score, 3),
                }
            )
            continue

        bike_count = normalize_current_status(matched_record.get("general_bikes"))
        ebike_count = normalize_current_status(matched_record.get("electric_bikes"))
        total_available = normalize_current_status(matched_record.get("available_spaces"))

        # 官網偶爾只缺其中一個車種明細；若總可借數完整，可用加總關係安全補算。
        if total_available is not None:
            if bike_count is None and ebike_count is not None and total_available >= ebike_count:
                bike_count = total_available - ebike_count
            elif ebike_count is None and bike_count is not None and total_available >= bike_count:
                ebike_count = total_available - bike_count

        service_status = safe_nonnegative_int(matched_record.get("service_status"))
        if service_status != 1 or bike_count is None or ebike_count is None:
            skipped_count += 1
            reason = "場站目前非正常服務" if service_status != 1 else "官網車種明細不完整"
            report_rows.append(
                {
                    "Excel 場站": excel_name,
                    "YouBike 場站": matched_record.get("station_name", ""),
                    "2.0": pd.NA if bike_count is None else bike_count,
                    "2.0E": pd.NA if ebike_count is None else ebike_count,
                    "結果": f"{reason}，未寫入",
                    "相似度": round(score, 3),
                }
            )
            continue

        used_station_ids.add(station_id)
        updated_df.at[row_index, "2.0 現況"] = bike_count
        updated_df.at[row_index, "2.0E 現況"] = ebike_count
        matched_count += 1
        report_rows.append(
            {
                "Excel 場站": excel_name,
                "YouBike 場站": matched_record.get("station_name", ""),
                "2.0": bike_count,
                "2.0E": ebike_count,
                "結果": "已寫入",
                "相似度": round(score, 3),
            }
        )

    report_df = pd.DataFrame(report_rows)
    summary = {
        "matched_count": matched_count,
        "skipped_count": skipped_count,
        "unmatched_count": unmatched_count,
        "total_count": len(updated_df),
    }
    return coerce_nullable_current_status(updated_df), report_df, summary


def partial_similarity(needle: str, haystack: str) -> float:
    """計算局部相似度；正規化結果會快取，並對高相似結果提前結束。"""
    needle_key = normalize_station_key(needle)
    haystack_key = normalize_station_key(haystack)
    if not needle_key or not haystack_key:
        return 0.0
    if needle_key in haystack_key:
        return 1.0

    direct = SequenceMatcher(None, needle_key, haystack_key, autojunk=False).ratio()
    if direct >= 0.86 or len(haystack_key) <= len(needle_key):
        return direct

    best = direct
    minimum = max(1, len(needle_key) - 2)
    maximum = min(len(haystack_key), len(needle_key) + 2)
    for window_size in range(minimum, maximum + 1):
        for start in range(0, len(haystack_key) - window_size + 1):
            window = haystack_key[start : start + window_size]
            score = SequenceMatcher(None, needle_key, window, autojunk=False).ratio()
            if score > best:
                best = score
                if best >= 0.96:
                    return best
    return best


def parse_count_token(value) -> int | None:
    """只接受獨立的整數儲存格，排除 D1、站名門牌與 2.0 欄位標題。"""
    text = normalize_ocr_text(value)
    if not text or any(symbol in text for symbol in (":", "/", "%")):
        return None

    # 2.0、2.0E 是欄位名稱，不是車數。
    if re.search(r"2\s*[.．,，]\s*0\s*[Ee]?", text):
        return None

    compact = re.sub(r"\s+", "", text)
    compact = compact.translate(str.maketrans({"O": "0", "o": "0", "〇": "0", "○": "0"}))
    compact = compact.strip("|[](){}<>，,。.;；")

    # 必須是單獨數字；這可避免把 D1 或「238巷」誤認成車數。
    match = re.fullmatch(r"([0-9]{1,3})(?:台)?", compact)
    if not match:
        return None

    number = int(match.group(1))
    return number if 0 <= number <= OCR_MAX_COUNT else None


def extract_numbers_from_row(row_items: list[dict]) -> list[tuple[int, float]]:
    """依畫面由左到右，取出一列中的獨立整數儲存格與水平位置。"""
    numbers: list[tuple[int, float]] = []
    for item in sorted(row_items, key=lambda current: current["x"]):
        number = parse_count_token(item["text"])
        if number is not None:
            numbers.append((number, float(item["x"])))
    return numbers


def detect_numeric_column_centers(rows: list[list[dict]]) -> list[float]:
    """從整張表格的重複數字位置推定「單位、在站2.0、在站2.0E」欄位。"""
    all_items = [item for row_items in rows for item in row_items]
    if not all_items:
        return []

    right_edge = max(float(item["x"]) + float(item.get("width", 1.0)) for item in all_items)
    unit_headers = [
        item
        for item in all_items
        if "單位" in normalize_ocr_text(item.get("text", ""))
    ]

    if unit_headers:
        # 以「單位」標題為起點，排除左側責任區 D1 與站名中的門牌號碼。
        unit_header_x = float(np.median([float(item["x"]) for item in unit_headers]))
        minimum_numeric_x = unit_header_x - max(20.0, right_edge * 0.025)
    else:
        # 固定版面中，單位欄位位於畫面約三分之一之後。
        minimum_numeric_x = right_edge * 0.30

    x_positions: list[float] = []
    widths: list[float] = []
    for item in all_items:
        if float(item["x"]) < minimum_numeric_x:
            continue
        if parse_count_token(item.get("text")) is None:
            continue
        x_positions.append(float(item["x"]))
        widths.append(max(1.0, float(item.get("width", 1.0))))

    if not x_positions:
        return []

    median_width = float(np.median(widths)) if widths else 12.0
    tolerance = max(12.0, median_width * 1.55)
    clusters: list[dict] = []

    for x_value in sorted(x_positions):
        nearest = None
        nearest_distance = None
        for cluster in clusters:
            distance = abs(x_value - float(cluster["center"]))
            if distance <= tolerance and (nearest_distance is None or distance < nearest_distance):
                nearest = cluster
                nearest_distance = distance
        if nearest is None:
            clusters.append({"values": [x_value], "center": x_value})
        else:
            nearest["values"].append(x_value)
            nearest["center"] = float(np.median(nearest["values"]))

    # 表格欄位會在多列重複；單次出現的時間、總數、座標等雜訊不採用。
    repeated = [cluster for cluster in clusters if len(cluster["values"]) >= 2]
    repeated.sort(key=lambda cluster: float(cluster["center"]))
    return [float(cluster["center"]) for cluster in repeated]


def choose_station_counts(
    numbers: list[tuple[int, float]],
    numeric_column_centers: list[float],
) -> tuple[int | None, int | None, int | None]:
    """固定讀取「單位」右側的在站 2.0、2.0E，而不是最右側綁車欄位。"""
    if len(numbers) < 3:
        return None, None, None

    # 表格由左至右前三個重複數字欄，依序是：單位、在站2.0、在站2.0E。
    if len(numeric_column_centers) >= 3:
        targets = numeric_column_centers[:3]
        gaps = [targets[index + 1] - targets[index] for index in range(2)]
        max_distance = max(14.0, min(gaps) * 0.46)
        selected: list[int] = []
        used_indexes: set[int] = set()

        for target in targets:
            candidates = [
                (abs(x_value - target), index, number)
                for index, (number, x_value) in enumerate(numbers)
                if index not in used_indexes
            ]
            if not candidates:
                return None, None, None
            distance, selected_index, selected_number = min(candidates, key=lambda item: item[0])
            if distance > max_distance:
                return None, None, None
            used_indexes.add(selected_index)
            selected.append(selected_number)

        unit_count, bike_count, ebike_count = selected
    else:
        # 後備規則：嚴格數字框由左至右前三個欄位。
        unit_count, bike_count, ebike_count = [number for number, _ in numbers[:3]]

    # 單位是場站總柱數；在站兩車種合計不應超過單位，異常列直接忽略。
    if unit_count <= 0 or bike_count < 0 or ebike_count < 0:
        return None, None, None
    if bike_count + ebike_count > unit_count:
        return None, None, None

    return unit_count, bike_count, ebike_count

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
    """校正方向與清晰度；先限制手機超高解析度照片，降低 OCR 計算量。"""
    image = Image.open(BytesIO(file_bytes))
    image = ImageOps.exif_transpose(image).convert("RGB")

    # 手機原圖常超過 4000px；縮到 2400px 長邊通常仍足以辨識表格文字，速度明顯較快。
    max_long_edge = 2400
    long_edge = max(image.width, image.height)
    if long_edge > max_long_edge:
        scale = max_long_edge / long_edge
        image = image.resize(
            (max(1, int(image.width * scale)), max(1, int(image.height * scale))),
            Image.Resampling.LANCZOS,
        )
    elif image.width < 1500:
        scale = min(1.75, 1500 / max(1, image.width))
        target_size = (int(image.width * scale), int(image.height * scale))
        if max(target_size) > max_long_edge:
            scale = max_long_edge / long_edge
            target_size = (int(image.width * scale), int(image.height * scale))
        image = image.resize(target_size, Image.Resampling.LANCZOS)

    image = ImageOps.autocontrast(image, cutoff=1)
    image = ImageEnhance.Sharpness(image).enhance(1.25)
    return np.asarray(image)



@st.cache_data(show_spinner=False, max_entries=48)
def _cached_ocr_geometry(file_bytes: bytes) -> tuple[list[list[dict]], list[dict]]:
    """依照片內容快取 OCR 結果；相同照片再次辨識時不重跑模型。"""
    engine = get_local_ocr_engine()
    image_array = prepare_image_for_ocr(file_bytes)
    result = engine(image_array)
    items = [
        item for item in decode_rapidocr_output(result)
        if item["text"] and item["score"] >= OCR_MIN_TEXT_CONFIDENCE
    ]
    return group_ocr_items_into_rows(items), items


def run_ocr_on_photo(file_name: str, file_bytes: bytes) -> tuple[list[list[dict]], list[dict]]:
    """辨識單張照片；幾何結果依照片內容快取，檔名在回傳時再附加。"""
    cached_rows, cached_items = _cached_ocr_geometry(file_bytes)
    items = [{**item, "file_name": file_name} for item in cached_items]
    item_lookup = {
        (item["text"], item["x"], item["y"], item["width"], item["height"]): item
        for item in items
    }
    rows = []
    for cached_row in cached_rows:
        rows.append([
            item_lookup.get(
                (item["text"], item["x"], item["y"], item["width"], item["height"]),
                {**item, "file_name": file_name},
            )
            for item in cached_row
        ])
    return rows, items



def pick_station_for_row(
    row_text: str,
    station_index: list[tuple[str, str]],
) -> tuple[str | None, float]:
    """從預先正規化的場站索引中找最接近的場站，避免每列重算所有站名。"""
    row_key = normalize_station_key(row_text)
    best_station = None
    best_score = 0.0
    for station_name, station_key in station_index:
        if not station_key:
            continue
        if station_key in row_key:
            return station_name, 1.0
        score = partial_similarity(station_key, row_key)
        if score > best_score:
            best_station = station_name
            best_score = score
            if best_score >= 0.96:
                break
    if best_score < OCR_MATCH_THRESHOLD:
        return None, best_score
    return best_station, best_score


def uploaded_photo_name_and_bytes(uploaded_photo) -> tuple[str, bytes]:
    """同時支援 Streamlit UploadedFile 與暫存在 session_state 的照片紀錄。"""
    if isinstance(uploaded_photo, dict):
        file_name = str(uploaded_photo.get("name") or "未命名照片")
        file_bytes = uploaded_photo.get("bytes", b"")
    else:
        file_name = str(getattr(uploaded_photo, "name", "未命名照片"))
        file_bytes = uploaded_photo.getvalue()

    if not isinstance(file_bytes, bytes):
        file_bytes = bytes(file_bytes)
    return file_name, file_bytes


def analyze_station_photos(uploaded_photos, station_df: pd.DataFrame) -> dict:
    """辨識照片中的在站數；成功資料直接更新，失敗資料明確標示識別錯誤。"""
    station_names = [
        str(value).strip()
        for value in station_df["場站名稱"].tolist()
        if str(value).strip()
    ]
    station_index = [(name, normalize_station_key(name)) for name in station_names]
    best_by_station: dict[str, dict] = {}
    errors_by_key: dict[tuple[str, str], dict] = {}
    scanned_row_count = 0

    for uploaded_photo in uploaded_photos:
        file_name, file_bytes = uploaded_photo_name_and_bytes(uploaded_photo)
        rows, _ = run_ocr_on_photo(file_name, file_bytes)
        numeric_column_centers = detect_numeric_column_centers(rows)
        photo_has_station_match = False
        photo_has_success = False

        for row_items in rows:
            row_text = " ".join(item["text"] for item in row_items).strip()
            if not row_text:
                continue
            scanned_row_count += 1

            station_name, match_score = pick_station_for_row(row_text, station_index)
            if not station_name:
                continue

            photo_has_station_match = True
            numbers = extract_numbers_from_row(row_items)
            unit_count, bike_count, ebike_count = choose_station_counts(
                numbers, numeric_column_centers
            )
            error_key = (file_name, station_name)

            if bike_count is None or ebike_count is None:
                errors_by_key[error_key] = {
                    "場站名稱": station_name,
                    "2.0 現況": RECOGNITION_ERROR_TEXT,
                    "2.0E 現況": RECOGNITION_ERROR_TEXT,
                    "來源": file_name,
                }
                continue

            photo_has_success = True
            errors_by_key.pop(error_key, None)
            row_score = float(np.mean([item["score"] for item in row_items]))
            evidence_score = (row_score * 0.45) + (match_score * 0.55)
            record = {
                "場站名稱": station_name,
                "單位": unit_count,
                "2.0 現況": bike_count,
                "2.0E 現況": ebike_count,
                "來源": file_name,
                "信心": round(evidence_score, 3),
            }

            # 重疊照片出現相同場站時，採用綜合信心較高的一列。
            previous = best_by_station.get(station_name)
            if previous is None or evidence_score > float(previous["信心"]):
                best_by_station[station_name] = record

        if not photo_has_station_match and not photo_has_success:
            errors_by_key[(file_name, "未識別場站")] = {
                "場站名稱": "未識別場站",
                "2.0 現況": RECOGNITION_ERROR_TEXT,
                "2.0E 現況": RECOGNITION_ERROR_TEXT,
                "來源": file_name,
            }

    recognized_rows = sorted(
        best_by_station.values(),
        key=lambda record: station_names.index(record["場站名稱"]),
    )
    updates = {
        record["場站名稱"]: (int(record["2.0 現況"]), int(record["2.0E 現況"]))
        for record in recognized_rows
    }

    return {
        "updates": updates,
        "recognized_df": pd.DataFrame(recognized_rows),
        "error_df": pd.DataFrame(list(errors_by_key.values())),
        "photo_count": len(uploaded_photos),
        "scanned_row_count": scanned_row_count,
    }



def apply_ocr_updates_to_dataframe(base_df: pd.DataFrame, updates: dict[str, tuple[int, int]]) -> pd.DataFrame:
    """只更新 OCR 成功辨識的場站；使用 map 避免逐列寫入。"""
    updated_df = base_df.copy()
    if not updates:
        return updated_df

    bike_updates = {name: values[0] for name, values in updates.items()}
    ebike_updates = {name: values[1] for name, values in updates.items()}
    station_series = updated_df["場站名稱"].astype(str)
    bike_mapped = station_series.map(bike_updates)
    ebike_mapped = station_series.map(ebike_updates)
    bike_mask = bike_mapped.notna()
    ebike_mask = ebike_mapped.notna()
    updated_df.loc[bike_mask, "2.0 現況"] = bike_mapped.loc[bike_mask].astype(int)
    updated_df.loc[ebike_mask, "2.0E 現況"] = ebike_mapped.loc[ebike_mask].astype(int)
    return coerce_nullable_current_status(updated_df)


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
    """讀取已保存的配置基底；不設定自動失效時間。"""
    if not token:
        return None, False

    excel_path, metadata_path = base_cache_paths(token)
    if not excel_path.exists() or not metadata_path.exists():
        return None, False

    try:
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        file_bytes = excel_path.read_bytes()
        return {
            "token": token,
            "name": str(metadata.get("name") or "配置基底.xlsx"),
            "bytes": file_bytes,
            "sha256": str(
                metadata.get("sha256") or hashlib.sha256(file_bytes).hexdigest()
            ),
            "uploaded_at": float(metadata.get("uploaded_at") or 0.0),
            "expires_at": None,
        }, False
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return None, False


def load_cached_status(token: str, expires_at: float | None = None) -> dict:
    """讀取與基底綁定的現況資料與同步資訊；不設定自動失效時間。"""
    status_path = current_status_cache_path(token)
    if not status_path.exists():
        return {"contexts": {}, "metadata": {}}

    try:
        payload = json.loads(status_path.read_text(encoding="utf-8"))
        contexts = payload.get("contexts", {})
        metadata = payload.get("metadata", {})
        if not isinstance(contexts, dict):
            contexts = {}
        if not isinstance(metadata, dict):
            metadata = {}
        return {"contexts": contexts, "metadata": metadata}
    except (OSError, TypeError, ValueError, json.JSONDecodeError):
        return {"contexts": {}, "metadata": {}}


def save_cached_status(
    token: str,
    expires_at: float | None,
    payload: dict,
) -> None:
    """保存現況資料；不設定自動失效時間。"""
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
            normalize_current_status(record.get("2.0 現況")),
            normalize_current_status(record.get("2.0E 現況")),
        )

    for row_index, row in restored_df.iterrows():
        key = (str(row.get("行政區", "")), str(row.get("場站名稱", "")))
        if key not in saved_values:
            continue
        bike_now, ebike_now = saved_values[key]
        restored_df.at[row_index, "2.0 現況"] = pd.NA if bike_now is None else bike_now
        restored_df.at[row_index, "2.0E 現況"] = pd.NA if ebike_now is None else ebike_now

    return coerce_nullable_current_status(restored_df)


def merge_current_status(full_df: pd.DataFrame, edited_df: pd.DataFrame) -> pd.DataFrame:
    """把目前畫面修改的現況合併回完整分區資料，避免切換行政區時遺失。"""
    merged_df = full_df.copy()
    edited_values = {
        (str(row.get("行政區", "")), str(row.get("場站名稱", ""))): (
            normalize_current_status(row.get("2.0 現況")),
            normalize_current_status(row.get("2.0E 現況")),
        )
        for _, row in edited_df.iterrows()
    }

    for row_index, row in merged_df.iterrows():
        key = (str(row.get("行政區", "")), str(row.get("場站名稱", "")))
        if key not in edited_values:
            continue
        bike_now, ebike_now = edited_values[key]
        merged_df.at[row_index, "2.0 現況"] = pd.NA if bike_now is None else bike_now
        merged_df.at[row_index, "2.0E 現況"] = pd.NA if ebike_now is None else ebike_now

    return coerce_nullable_current_status(merged_df)



def dataframe_to_status_records(status_df: pd.DataFrame) -> list[dict]:
    """將完整現況轉成 JSON 紀錄；以 DataFrame 向量化處理。"""
    columns = ["行政區", "場站名稱", "2.0 現況", "2.0E 現況"]
    records_df = coerce_nullable_current_status(status_df[columns].copy())
    records_df["行政區"] = records_df["行政區"].astype(str)
    records_df["場站名稱"] = records_df["場站名稱"].astype(str)
    object_df = records_df.astype(object).where(records_df.notna(), None)
    return object_df.to_dict(orient="records")


def save_cached_base(file_name: str, file_bytes: bytes) -> dict:
    """將新上傳的 Excel 保存於伺服器，不設定自動失效時間。"""
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
        "expires_at": None,
    }


def format_remaining_time(expires_at: float | None) -> str:
    """相容舊介面；目前保存期限為無時間限制。"""
    return "無時間限制"


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


def build_blank_status_cache(workbook_data: dict, options: list[tuple[str, str]]) -> dict:
    """把所有工作表、分區與班別的現況清成空白。"""
    blank_contexts: dict[str, list[dict]] = {}

    for sheet_name, route in options:
        for shift in SHIFT_COLUMNS.keys():
            try:
                route_df = parse_route(workbook_data[sheet_name], route, shift)
            except Exception:
                continue

            if route_df.empty:
                continue

            context_key = status_context_key(sheet_name, route, shift)
            blank_contexts[context_key] = dataframe_to_status_records(
                blank_current_status(route_df)
            )

    return {"contexts": blank_contexts, "metadata": {}}


def dispatch_status_class(value) -> str:
    """回傳分析結果儲存格的視覺狀態類別。"""
    text = str(value)
    if "多" in text:
        return "analysis-status-extra"
    if "缺" in text or "少" in text:
        return "analysis-status-short"
    if RECOGNITION_ERROR_TEXT in text:
        return "analysis-status-error"
    return "analysis-status-ok"


def render_analysis_result_table(region_df: pd.DataFrame) -> None:
    """以響應式 HTML 表格完整展開分析結果，避免表格內水平或垂直捲動。"""
    rows = []
    for row_index, row in region_df.iterrows():
        station_name = html.escape(str(row.get("場站名稱", "")))
        bike_status = add_dispatch_indicator(row.get("2.0 缺／多幾台", ""))
        ebike_status = add_dispatch_indicator(row.get("2.0E 缺／多幾台", ""))
        bike_class = dispatch_status_class(bike_status)
        ebike_class = dispatch_status_class(ebike_status)
        rows.append(
            f'<tr id="analysis-result-anchor-{int(row_index)}" class="analysis-result-row">'
            f'<td class="analysis-station-cell">{station_name}</td>'
            f'<td class="analysis-status-cell {bike_class}">{html.escape(bike_status)}</td>'
            f'<td class="analysis-status-cell {ebike_class}">{html.escape(ebike_status)}</td>'
            "</tr>"
        )

    table_html = (
        '<div class="analysis-result-table-wrap">'
        '<table class="analysis-result-table">'
        '<colgroup>'
        '<col class="analysis-col-station" />'
        '<col class="analysis-col-status" />'
        '<col class="analysis-col-status" />'
        '</colgroup>'
        '<thead><tr>'
        '<th>場站名稱</th>'
        '<th>2.0 缺／多</th>'
        '<th>2.0E 缺／多</th>'
        '</tr></thead>'
        f'<tbody>{"".join(rows)}</tbody>'
        '</table></div>'
    )
    st.markdown(table_html, unsafe_allow_html=True)


def render_floating_station_search(result_df: pd.DataFrame, mobile_mode: bool) -> None:
    """建立搜尋分析結果、回頁首與跳至分析區的垂直懸浮按鍵。"""
    stations = []
    for row_index, row in result_df.reset_index(drop=True).iterrows():
        station_name = str(row.get("場站名稱", "")).strip()
        if not station_name:
            continue
        stations.append(
            {
                "name": station_name,
                "region": str(row.get("行政區", "")).strip(),
                "bike": str(row.get("2.0 缺／多幾台", "")).strip(),
                "ebike": str(row.get("2.0E 缺／多幾台", "")).strip(),
                "anchor": f"analysis-result-anchor-{row_index}",
            }
        )

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

            const oldRoot = doc.getElementById("ubike-float-tools");
            if (oldRoot) oldRoot.remove();
            const oldStyle = doc.getElementById("ubike-float-tools-style");
            if (oldStyle) oldStyle.remove();

            const style = doc.createElement("style");
            style.id = "ubike-float-tools-style";
            style.textContent = `
                #ubike-float-tools {{
                    position: fixed;
                    right: 18px;
                    bottom: 22px;
                    z-index: 2147483000;
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI",
                                 "Microsoft JhengHei", sans-serif;
                }}
                #ubike-float-tools * {{ box-sizing: border-box; }}
                #ubike-float-tools .uft-actions {{
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                    gap: 8px;
                }}
                #ubike-float-tools .uft-button {{
                    width: 56px;
                    height: 56px;
                    flex: 0 0 56px;
                    border: 0;
                    border-radius: 50%;
                    color: #151515;
                    font-weight: 850;
                    cursor: pointer;
                    box-shadow: 0 8px 28px rgba(0, 0, 0, 0.28);
                    transition: transform 0.16s ease, box-shadow 0.16s ease;
                }}
                #ubike-float-tools .uft-button:hover {{
                    transform: translateY(-2px);
                    box-shadow: 0 11px 32px rgba(0, 0, 0, 0.34);
                }}
                #ubike-float-tools .uft-top {{ background: #f1f3f5; font-size: 12px; }}
                #ubike-float-tools .uft-analysis {{ background: #8bd3ff; font-size: 13px; }}
                #ubike-float-tools .uft-search {{ background: #ffbf00; font-size: 24px; }}
                #ubike-float-tools .uft-refresh {{ background: #6ee7b7; font-size: 13px; }}
                #ubike-float-tools .uft-refresh.is-syncing {{
                    cursor: wait;
                    animation: ubikeRefreshPulse 0.9s ease-in-out infinite;
                }}
                @keyframes ubikeRefreshPulse {{
                    0%, 100% {{ transform: scale(1); }}
                    50% {{ transform: scale(0.91); }}
                }}
                #ubike-float-tools .uft-panel {{
                    position: absolute;
                    right: 66px;
                    bottom: 0;
                    width: min(340px, calc(100vw - 96px));
                    padding: 13px;
                    border: 1px solid rgba(0, 0, 0, 0.13);
                    border-radius: 16px;
                    background: rgba(255, 255, 255, 0.98);
                    color: #171717;
                    box-shadow: 0 16px 46px rgba(0, 0, 0, 0.28);
                    backdrop-filter: blur(8px);
                }}
                #ubike-float-tools .uft-panel[hidden] {{ display: none !important; }}
                #ubike-float-tools .uft-title-row {{
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    margin-bottom: 9px;
                }}
                #ubike-float-tools .uft-title {{ font-size: 15px; font-weight: 800; }}
                #ubike-float-tools .uft-close {{
                    border: 0;
                    background: transparent;
                    color: #555;
                    font-size: 20px;
                    cursor: pointer;
                }}
                #ubike-float-tools .uft-input {{
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
                #ubike-float-tools .uft-input:focus {{
                    border-color: #ffbf00;
                    box-shadow: 0 0 0 3px rgba(255, 191, 0, 0.18);
                }}
                #ubike-float-tools .uft-hint {{
                    margin: 7px 2px 8px;
                    color: #666;
                    font-size: 12px;
                }}
                #ubike-float-tools .uft-results {{
                    display: flex;
                    flex-direction: column;
                    gap: 5px;
                    max-height: min(310px, 48vh);
                    overflow-y: auto;
                    overflow-x: hidden;
                }}
                #ubike-float-tools .uft-result {{
                    width: 100%;
                    padding: 9px 10px;
                    border: 1px solid #e4e4e4;
                    border-radius: 10px;
                    background: #fff;
                    text-align: left;
                    cursor: pointer;
                }}
                #ubike-float-tools .uft-result:hover,
                #ubike-float-tools .uft-result:focus {{
                    border-color: #ffbf00;
                    background: #fff9df;
                    outline: none;
                }}
                #ubike-float-tools .uft-result-name {{
                    display: block;
                    color: #111;
                    font-size: 14px;
                    font-weight: 750;
                }}
                #ubike-float-tools .uft-result-region,
                #ubike-float-tools .uft-result-status {{
                    display: block;
                    margin-top: 2px;
                    color: #777;
                    font-size: 12px;
                    line-height: 1.35;
                }}
                #ubike-float-tools .uft-result-status {{ color: #444; }}
                #ubike-float-tools .uft-empty {{
                    padding: 12px 4px 5px;
                    color: #777;
                    font-size: 13px;
                    text-align: center;
                }}
                #ubike-search-toast {{
                    position: fixed;
                    left: 50%;
                    bottom: var(--ubike-float-toast-bottom, 210px);
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
                .ubike-analysis-row-focus {{
                    position: relative;
                    z-index: 2;
                    outline: 3px solid #ff9f00;
                    outline-offset: -3px;
                    animation: ubikeAnalysisPulse 0.72s ease-in-out 3;
                }}
                @keyframes ubikeAnalysisPulse {{
                    0%, 100% {{ filter: brightness(1); }}
                    50% {{ filter: brightness(1.16); }}
                }}
                @media (max-width: 700px) {{
                    #ubike-float-tools {{
                        right: 10px;
                        bottom: calc(72px + env(safe-area-inset-bottom, 0px));
                    }}
                    #ubike-float-tools .uft-button {{
                        width: 52px;
                        height: 52px;
                        flex-basis: 52px;
                    }}
                    #ubike-float-tools .uft-panel {{
                        right: 60px;
                        bottom: 0;
                        width: min(305px, calc(100vw - 82px));
                        max-height: 72vh;
                    }}
                }}
            `;
            doc.head.appendChild(style);

            const root = doc.createElement("div");
            root.id = "ubike-float-tools";
            root.innerHTML = `
                <div class="uft-panel" hidden>
                    <div class="uft-title-row">
                        <div class="uft-title">🔎 搜尋分析結果</div>
                        <button class="uft-close" type="button" aria-label="關閉搜尋">×</button>
                    </div>
                    <input class="uft-input" type="search"
                           placeholder="輸入場站名稱或行政區" autocomplete="off" />
                    <div class="uft-hint">只搜尋目前的調度分析結果｜快捷鍵 Ctrl + K</div>
                    <div class="uft-results"></div>
                </div>
                <div class="uft-actions">
                    <button class="uft-button uft-top" type="button" title="回到頁面最上方">TOP</button>
                    <button class="uft-button uft-analysis" type="button" title="跳到調度分析結果">分析</button>
                    <button class="uft-button uft-search" type="button" title="搜尋分析結果（Ctrl + K）">🔎</button>
                    <button class="uft-button uft-refresh" type="button" title="手動更新 YouBike 即時車數">更新</button>
                </div>
            `;
            doc.body.appendChild(root);

            const topButton = root.querySelector(".uft-top");
            const analysisButton = root.querySelector(".uft-analysis");
            const searchButton = root.querySelector(".uft-search");
            const refreshButton = root.querySelector(".uft-refresh");
            const panel = root.querySelector(".uft-panel");
            const closeButton = root.querySelector(".uft-close");
            const input = root.querySelector(".uft-input");
            const results = root.querySelector(".uft-results");

            function isMobileLayout() {{
                return displayMode === "mobile" || win.matchMedia("(max-width: 700px)").matches;
            }}

            function updateFloatingPosition() {{
                // 固定成本：不再掃描整個頁面 DOM。手機保留底部安全距離即可。
                const bottomGap = isMobileLayout() ? 72 : 22;
                root.style.bottom = isMobileLayout()
                    ? `calc(${{bottomGap}}px + env(safe-area-inset-bottom, 0px))`
                    : `${{bottomGap}}px`;
                doc.documentElement.style.setProperty(
                    "--ubike-float-toast-bottom",
                    `${{bottomGap + 252}}px`,
                );
            }}

            function setRefreshButtonState(syncing) {{
                refreshButton.disabled = Boolean(syncing);
                refreshButton.classList.toggle("is-syncing", Boolean(syncing));
                refreshButton.textContent = syncing ? "更新中" : "更新";
                refreshButton.title = syncing
                    ? "正在更新 YouBike 即時車數"
                    : "手動更新 YouBike 即時車數";
            }}

            function requestManualSync() {{
                let postedCount = 0;
                for (const frame of doc.querySelectorAll("iframe")) {{
                    try {{
                        if (!frame.contentWindow) continue;
                        const frameTitle = String(frame.getAttribute("title") || "").toLowerCase();
                        const frameSource = String(frame.getAttribute("src") || "").toLowerCase();
                        let isSyncFrame = frameTitle.includes("youbike_browser_sync")
                            || frameSource.includes("youbike_browser_sync");
                        try {{
                            isSyncFrame = isSyncFrame
                                || Boolean(frame.contentDocument?.getElementById("syncButton"));
                        }} catch (_accessError) {{
                            // 跨來源時改以 title／src 判斷。
                        }}
                        if (!isSyncFrame) continue;
                        frame.contentWindow.postMessage({{ type: "ubike:manual-sync" }}, "*");
                        postedCount += 1;
                    }} catch (_error) {{
                        // 略過無法存取的其他 iframe。
                    }}
                }}
                if (!postedCount) {{
                    showToast("同步元件尚未準備完成，請稍後再按一次");
                    return;
                }}
                setRefreshButtonState(true);
                showToast("正在手動更新 YouBike 即時車數");
                // 防止外部網路錯誤造成按鈕永久鎖住；元件回報時會更早解除。
                win.clearTimeout(win.__ubikeManualSyncFallbackTimer);
                win.__ubikeManualSyncFallbackTimer = win.setTimeout(() => {{
                    setRefreshButtonState(false);
                }}, 45000);
            }}

            function getScrollTargets() {{
                const candidates = [
                    doc.scrollingElement,
                    doc.documentElement,
                    doc.body,
                    doc.querySelector('[data-testid="stAppViewContainer"]'),
                    doc.querySelector('[data-testid="stMain"]'),
                    doc.querySelector('section.main'),
                    doc.querySelector('.stApp'),
                ];
                return Array.from(new Set(candidates.filter(Boolean)));
            }}

            function scrollPageToTop() {{
                setOpen(false);
                const topAnchor = doc.getElementById("ubike-page-top-anchor");
                if (topAnchor) topAnchor.scrollIntoView({{ behavior: "smooth", block: "start" }});

                for (const target of getScrollTargets()) {{
                    try {{
                        target.scrollTo({{ top: 0, left: 0, behavior: "smooth" }});
                    }} catch (_error) {{
                        target.scrollTop = 0;
                        target.scrollLeft = 0;
                    }}
                }}
                try {{
                    win.scrollTo({{ top: 0, left: 0, behavior: "smooth" }});
                }} catch (_error) {{
                    win.scrollTo(0, 0);
                }}
                win.setTimeout(() => {{
                    for (const target of getScrollTargets()) {{
                        target.scrollTop = 0;
                        target.scrollLeft = 0;
                    }}
                    win.scrollTo(0, 0);
                }}, 420);
                showToast("已回到頁面最上方");
            }}

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

            function jumpToStation(station) {{
                setOpen(false);
                const anchor = doc.getElementById(station.anchor);
                if (!anchor) {{
                    showToast("找不到這個分析結果，請重新整理後再試");
                    return;
                }}

                anchor.scrollIntoView({{ behavior: "smooth", block: "center" }});
                anchor.classList.remove("ubike-analysis-row-focus");
                void anchor.offsetWidth;
                anchor.classList.add("ubike-analysis-row-focus");
                win.setTimeout(() => anchor.classList.remove("ubike-analysis-row-focus"), 2400);
                showToast(`已找到分析結果：${{station.name}}`);
            }}

            function createResultButton(station) {{
                const button = doc.createElement("button");
                button.type = "button";
                button.className = "uft-result";

                const name = doc.createElement("span");
                name.className = "uft-result-name";
                name.textContent = station.name;
                button.appendChild(name);

                if (station.region) {{
                    const region = doc.createElement("span");
                    region.className = "uft-result-region";
                    region.textContent = station.region;
                    button.appendChild(region);
                }}

                const status = doc.createElement("span");
                status.className = "uft-result-status";
                status.textContent = `2.0：${{station.bike || "—"}}｜2.0E：${{station.ebike || "—"}}`;
                button.appendChild(status);

                button.addEventListener("click", () => jumpToStation(station));
                return button;
            }}

            function renderResults(query) {{
                const keyword = String(query || "").trim().toLocaleLowerCase("zh-TW");
                const matched = stations
                    .filter((station) => {{
                        const haystack = `${{station.name}} ${{station.region}} ${{station.bike}} ${{station.ebike}}`
                            .toLocaleLowerCase("zh-TW");
                        return !keyword || haystack.includes(keyword);
                    }})
                    .slice(0, 12);

                results.replaceChildren();
                if (!matched.length) {{
                    const empty = doc.createElement("div");
                    empty.className = "uft-empty";
                    empty.textContent = stations.length
                        ? "分析結果中查無符合的場站"
                        : "目前沒有需要調度的分析結果";
                    results.appendChild(empty);
                    return;
                }}
                matched.forEach((station) => results.appendChild(createResultButton(station)));
            }}

            topButton.addEventListener("click", scrollPageToTop);
            analysisButton.addEventListener("click", () => {{
                setOpen(false);
                const anchor = doc.getElementById("analysis-results-anchor");
                if (!anchor) {{
                    showToast("找不到分析結果區");
                    return;
                }}
                anchor.scrollIntoView({{ behavior: "smooth", block: "start" }});
            }});
            searchButton.addEventListener("click", () => setOpen(panel.hidden));
            refreshButton.addEventListener("click", requestManualSync);
            closeButton.addEventListener("click", () => setOpen(false));
            input.addEventListener("input", (event) => renderResults(event.target.value));
            input.addEventListener("keydown", (event) => {{
                if (event.key === "Enter") {{
                    const firstResult = results.querySelector(".uft-result");
                    if (firstResult) firstResult.click();
                }}
                if (event.key === "Escape") setOpen(false);
            }});

            if (win.__ubikeSyncStateHandler) {{
                win.removeEventListener("message", win.__ubikeSyncStateHandler);
            }}
            win.__ubikeSyncStateHandler = (event) => {{
                const data = event.data || {{}};
                if (data.source !== "ubike-browser-sync" || data.type !== "ubike:sync-state") return;
                if (data.state === "busy") {{
                    setRefreshButtonState(true);
                    return;
                }}
                win.clearTimeout(win.__ubikeManualSyncFallbackTimer);
                setRefreshButtonState(false);
                if (data.state === "success") {{
                    const countText = Number(data.station_count) > 0 ? `（${{Number(data.station_count)}} 站）` : "";
                    showToast(`即時數據更新完成${{countText}}`);
                }} else if (data.state === "error") {{
                    showToast(`即時數據更新失敗：${{String(data.message || "請稍後再試")}}`);
                }}
            }};
            win.addEventListener("message", win.__ubikeSyncStateHandler);

            if (win.__ubikeSearchKeyHandler) {{
                win.removeEventListener("keydown", win.__ubikeSearchKeyHandler);
            }}
            win.__ubikeSearchKeyHandler = (event) => {{
                if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {{
                    event.preventDefault();
                    setOpen(true);
                }}
                if (event.key === "Escape" && !panel.hidden) setOpen(false);
            }};
            win.addEventListener("keydown", win.__ubikeSearchKeyHandler);

            if (win.__ubikeFloatingResizeHandler) {{
                win.removeEventListener("resize", win.__ubikeFloatingResizeHandler);
                if (win.visualViewport) {{
                    win.visualViewport.removeEventListener("resize", win.__ubikeFloatingResizeHandler);
                    win.visualViewport.removeEventListener("scroll", win.__ubikeFloatingResizeHandler);
                }}
            }}
            win.__ubikeFloatingResizeHandler = () => updateFloatingPosition();
            win.addEventListener("resize", win.__ubikeFloatingResizeHandler, {{ passive: true }});
            if (win.visualViewport) {{
                win.visualViewport.addEventListener("resize", win.__ubikeFloatingResizeHandler, {{ passive: true }});
                win.visualViewport.addEventListener("scroll", win.__ubikeFloatingResizeHandler, {{ passive: true }});
            }}

            // 舊版若已建立 MutationObserver，先關閉；之後只在視窗尺寸改變時更新。
            if (win.__ubikeFloatingObserver) {{
                win.__ubikeFloatingObserver.disconnect();
                win.__ubikeFloatingObserver = null;
            }}

            updateFloatingPosition();
            win.setTimeout(updateFloatingPosition, 350);
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
    html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
        max-width: 100%;
        overflow-x: hidden !important;
    }
    .block-container {
        width: 100%;
        max-width: 100%;
        padding-top: 1.3rem;
        padding-bottom: 3rem;
        overflow-x: clip;
    }
    [data-testid="stMetricValue"] {font-size: 1.65rem;}
    div[data-testid="stDataEditor"] {border: 1px solid #dddddd; border-radius: 10px;}
    div[data-testid="stNumberInput"] input {text-align: center;}

    .analysis-result-table-wrap {
        width: 100%;
        max-width: 100%;
        margin: 0.15rem 0 0.75rem;
        overflow: visible;
    }
    .analysis-result-table {
        width: 100%;
        max-width: 100%;
        table-layout: fixed;
        border-collapse: separate;
        border-spacing: 0;
        border: 1px solid #d9dee5;
        border-radius: 10px;
        overflow: hidden;
        background: white;
    }
    .analysis-result-table .analysis-col-station {width: 42%;}
    .analysis-result-table .analysis-col-status {width: 29%;}
    .analysis-result-table th,
    .analysis-result-table td {
        padding: 8px 6px;
        border-right: 1px solid #e4e8ed;
        border-bottom: 1px solid #e4e8ed;
        vertical-align: middle;
        line-height: 1.3;
        white-space: normal;
        overflow-wrap: anywhere;
        word-break: break-word;
    }
    .analysis-result-table th:last-child,
    .analysis-result-table td:last-child {border-right: 0;}
    .analysis-result-table tbody tr:last-child td {border-bottom: 0;}
    .analysis-result-table th {
        background: #f4f6f8;
        color: #222;
        font-size: 0.88rem;
        font-weight: 800;
        text-align: center;
    }
    .analysis-result-table td {
        font-size: 0.9rem;
    }
    .analysis-station-cell {
        font-weight: 750;
        text-align: left;
    }
    .analysis-status-cell {
        text-align: center;
        font-weight: 750;
    }
    .analysis-status-extra {background: #fce1e1;}
    .analysis-status-short {background: #fff0df;}
    .analysis-status-error {background: #fff7cc;}
    .analysis-status-ok {background: #e7f5e9;}

    .fleet-summary-card {
        --fleet-accent: #d88900;
        --fleet-soft: #fff8e8;
        --fleet-border: #f5d88f;
        position: relative;
        display: grid;
        grid-template-columns: 118px minmax(0, 1fr);
        gap: 0.7rem;
        width: 100%;
        margin: 0.8rem 0;
        padding: 1rem 1.05rem;
        overflow: hidden;
        border: 1px solid var(--fleet-border);
        border-radius: 20px;
        background:
            radial-gradient(circle at 12% 20%, rgba(255,255,255,.95), rgba(255,255,255,0) 35%),
            linear-gradient(135deg, var(--fleet-soft), #ffffff 70%);
        box-shadow: 0 10px 28px rgba(31, 41, 55, 0.07);
    }
    .fleet-theme-ebike {
        --fleet-accent: #2468c9;
        --fleet-soft: #eef5ff;
        --fleet-border: #bdd4fb;
    }
    .fleet-card-illustration {
        min-width: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        color: var(--fleet-accent);
        border-radius: 18px;
        background: color-mix(in srgb, var(--fleet-accent) 10%, white);
    }
    .fleet-card-svg {
        width: 100%;
        max-width: 110px;
        height: auto;
        filter: drop-shadow(0 7px 9px rgba(31, 41, 55, .08));
    }
    .fleet-card-content {min-width: 0;}
    .fleet-card-heading {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 0.8rem;
        margin-bottom: 0.7rem;
    }
    .fleet-card-title {
        color: var(--fleet-accent);
        font-size: 1.55rem;
        font-weight: 900;
        line-height: 1.1;
        letter-spacing: .01em;
    }
    .fleet-card-subtitle {
        margin-top: .22rem;
        color: #7b8494;
        font-size: .78rem;
        font-weight: 650;
    }
    .fleet-state-badge {
        flex: 0 0 auto;
        display: inline-flex;
        align-items: center;
        min-height: 34px;
        padding: .32rem .72rem;
        border-radius: 999px;
        font-size: .85rem;
        font-weight: 850;
        white-space: nowrap;
        border: 1px solid transparent;
    }
    .fleet-state-short {
        color: #d46700;
        background: #fff1d7;
        border-color: #ffc96c;
    }
    .fleet-state-extra {
        color: #d7243f;
        background: #ffe5e9;
        border-color: #ffabb7;
    }
    .fleet-state-balanced {
        color: #187642;
        background: #e4f7ec;
        border-color: #9fd8b8;
    }
    .fleet-state-pending {
        color: #806000;
        background: #fff7cc;
        border-color: #ead27a;
    }
    .fleet-card-metrics {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        align-items: stretch;
    }
    .fleet-metric-block {
        min-width: 0;
        padding: 0 .8rem;
        text-align: center;
        border-left: 1px dashed #cfd5df;
    }
    .fleet-metric-block:first-child {
        padding-left: 0;
        border-left: 0;
    }
    .fleet-metric-block:last-child {padding-right: 0;}
    .fleet-metric-label {
        min-height: 1.4rem;
        color: #555f6f;
        font-size: .84rem;
        font-weight: 750;
    }
    .fleet-metric-value {
        margin-top: .32rem;
        color: #151922;
        font-size: clamp(1.55rem, 4vw, 2.25rem);
        font-weight: 900;
        line-height: 1;
        white-space: nowrap;
    }
    .fleet-metric-value span {
        margin-left: .12rem;
        font-size: .56em;
        font-weight: 800;
    }
    .fleet-difference-chip {
        display: flex;
        min-height: 58px;
        margin-top: .25rem;
        padding: .4rem .3rem;
        flex-direction: column;
        justify-content: center;
        border-radius: 13px;
        line-height: 1.06;
    }
    .fleet-difference-chip strong {
        font-size: clamp(1.05rem, 3.4vw, 1.55rem);
        font-weight: 950;
        white-space: nowrap;
    }
    .fleet-difference-chip small {
        margin-top: .22rem;
        font-size: .72rem;
        font-weight: 750;
        opacity: .72;
    }
    .fleet-difference-short {color: #ef7200; background: #fff0d5;}
    .fleet-difference-extra {color: #df203b; background: #ffe1e6;}
    .fleet-difference-balanced {color: #197a45; background: #e2f6ea;}
    .fleet-difference-pending {color: #876900; background: #fff7d4;}

    .fleet-data-notice {
        position: relative;
        display: grid;
        grid-template-columns: 44px minmax(0, 1fr) 56px;
        align-items: center;
        gap: .7rem;
        margin: 1rem 0 .8rem;
        padding: .95rem 1rem;
        overflow: hidden;
        border: 1px solid #f1d67a;
        border-radius: 17px;
        background: linear-gradient(135deg, #fff9dc, #fffdf2);
        box-shadow: 0 7px 20px rgba(91, 67, 0, .05);
    }
    .fleet-notice-icon {
        display: grid;
        width: 40px;
        height: 40px;
        place-items: center;
        color: white;
        font-size: 1.45rem;
        font-weight: 950;
        border-radius: 13px 13px 17px 17px;
        background: #f0a300;
        clip-path: polygon(50% 0, 100% 100%, 0 100%);
        padding-top: 9px;
    }
    .fleet-notice-title {
        color: #8d6100;
        font-size: 1.02rem;
        font-weight: 900;
    }
    .fleet-notice-text {
        margin-top: .2rem;
        color: #7b5a0b;
        font-size: .88rem;
        line-height: 1.55;
    }
    .fleet-notice-decoration {
        font-size: 2.25rem;
        text-align: right;
        filter: saturate(.8);
        opacity: .75;
    }
    .fleet-legend {
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: .35rem;
        margin: .6rem 0 .4rem;
        color: #707887;
        font-size: .9rem;
        font-weight: 700;
    }
    .fleet-legend span {display: inline-flex; align-items: center; gap: .3rem;}
    .fleet-legend-dot {
        display: inline-block;
        width: 16px;
        height: 16px;
        border-radius: 50%;
        box-shadow: inset 0 1px 2px rgba(255,255,255,.7), 0 2px 4px rgba(31,41,55,.14);
    }
    .fleet-legend-extra {background: linear-gradient(#ff8a96, #ef4457);}
    .fleet-legend-short {background: linear-gradient(#ffd07e, #f1a43a);}
    .fleet-legend-balanced {background: linear-gradient(#8cd6a9, #39a96b);}
    .fleet-legend-divider {opacity: .48;}

    /* 手機版：縮小頁面留白，並放大數字輸入框，方便直接叫出九宮格鍵盤。 */
    @media (max-width: 900px) {
        .block-container {
            padding-left: 0.22rem;
            padding-right: 0.22rem;
            padding-top: 0.7rem;
        }

        .analysis-result-table-wrap {
            margin-left: 0;
            margin-right: 0;
        }
        .analysis-result-table .analysis-col-station {width: 40%;}
        .analysis-result-table .analysis-col-status {width: 30%;}
        .analysis-result-table th,
        .analysis-result-table td {
            padding: 7px 3px;
            font-size: 0.78rem;
            line-height: 1.24;
        }
        .analysis-result-table th {
            font-size: 0.75rem;
        }

        .fleet-summary-card {
            grid-template-columns: 78px minmax(0, 1fr);
            gap: .48rem;
            margin: .65rem 0;
            padding: .8rem .72rem;
            border-radius: 16px;
        }
        .fleet-card-illustration {border-radius: 14px;}
        .fleet-card-svg {max-width: 72px;}
        .fleet-card-heading {gap: .42rem; margin-bottom: .62rem;}
        .fleet-card-title {font-size: 1.25rem;}
        .fleet-card-subtitle {font-size: .67rem;}
        .fleet-state-badge {min-height: 29px; padding: .25rem .5rem; font-size: .72rem;}
        .fleet-metric-block {padding: 0 .32rem;}
        .fleet-metric-label {font-size: .7rem; min-height: 1.2rem;}
        .fleet-metric-value {font-size: clamp(1.22rem, 6vw, 1.7rem);}
        .fleet-difference-chip {min-height: 49px; padding: .3rem .14rem; border-radius: 10px;}
        .fleet-difference-chip strong {font-size: clamp(.84rem, 4vw, 1.12rem);}
        .fleet-difference-chip small {font-size: .61rem;}
        .fleet-data-notice {
            grid-template-columns: 35px minmax(0, 1fr) 38px;
            gap: .48rem;
            padding: .78rem .72rem;
            border-radius: 14px;
        }
        .fleet-notice-icon {width: 34px; height: 34px; font-size: 1.15rem; padding-top: 8px;}
        .fleet-notice-title {font-size: .92rem;}
        .fleet-notice-text {font-size: .77rem; line-height: 1.45;}
        .fleet-notice-decoration {font-size: 1.7rem;}

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




DISPATCH_REJECTION_REASONS = (
    "路況壅塞／繞路",
    "無法臨停或停車",
    "現場暫時無法作業",
    "即時資料與現場不符",
    "目前貨車載量或車種不適合",
    "想先處理其他場站",
    "其他原因",
)
DISPATCH_REASON_SCORE_MULTIPLIERS = {
    "路況壅塞／繞路": 0.82,
    "無法臨停或停車": 0.76,
    "現場暫時無法作業": 0.80,
    "即時資料與現場不符": 0.86,
    "目前貨車載量或車種不適合": 0.90,
    "想先處理其他場站": 0.95,
    "其他原因": 0.93,
}
DISPATCH_COOLDOWN_SECONDS = 10 * 60
DISPATCH_ESTIMATED_SPEED_KMH = 32.0
DISPATCH_ROAD_DISTANCE_FACTOR = 1.22
DISPATCH_OPERATION_BASE_MINUTES = 2.0
DISPATCH_OPERATION_MINUTES_PER_BIKE = 0.75


DISPATCH_GEOLOCATION_COMPONENT_HTML = r"""<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; background: transparent; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
    #locateButton {
      width: 100%; min-height: 50px; border: 0; border-radius: 13px;
      padding: 10px 14px; font-size: 16px; font-weight: 750;
      color: #fff; background: #1677ff; cursor: pointer;
      -webkit-tap-highlight-color: transparent; touch-action: manipulation;
    }
    #locateButton:disabled { opacity: .68; cursor: wait; }
    #status { min-height: 18px; margin-top: 6px; padding: 0 3px; font-size: 13px; color: #6b7280; }
    .error { color: #c62828 !important; }
  </style>
</head>
<body>
  <button id="locateButton" type="button">📍 取得／更新目前位置</button>
  <div id="status">第一次使用時，瀏覽器會詢問定位權限。</div>
<script>
(() => {
  const API_VERSION = 1;
  const button = document.getElementById("locateButton");
  const statusNode = document.getElementById("status");

  function send(type, data = {}) {
    window.parent.postMessage({ isStreamlitMessage: true, type, ...data }, "*");
  }
  function setHeight() {
    send("streamlit:setFrameHeight", { height: Math.max(78, document.body.scrollHeight + 2) });
  }
  function setValue(value) {
    send("streamlit:setComponentValue", { value, dataType: "json" });
  }
  function eventId() {
    if (globalThis.crypto && typeof globalThis.crypto.randomUUID === "function") return globalThis.crypto.randomUUID();
    return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }
  function setStatus(text, isError = false) {
    statusNode.textContent = text || "";
    statusNode.className = isError ? "error" : "";
    setHeight();
  }
  button.addEventListener("click", () => {
    if (!navigator.geolocation) {
      setStatus("此瀏覽器不支援定位。", true);
      setValue({ ok: false, event_id: eventId(), error: "此瀏覽器不支援定位" });
      return;
    }
    button.disabled = true;
    button.textContent = "⏳ 正在取得目前位置……";
    setStatus("請允許瀏覽器使用定位；室外或靠近窗邊通常較準。", false);
    navigator.geolocation.getCurrentPosition(
      position => {
        const payload = {
          ok: true,
          event_id: eventId(),
          latitude: Number(position.coords.latitude),
          longitude: Number(position.coords.longitude),
          accuracy: Number(position.coords.accuracy || 0),
          timestamp: Number(position.timestamp || Date.now()),
        };
        setValue(payload);
        setStatus(`定位完成，誤差約 ${Math.round(payload.accuracy)} 公尺。`, false);
        button.disabled = false;
        button.textContent = "📍 重新取得目前位置";
      },
      error => {
        const message = error && error.message ? error.message : "定位失敗";
        setValue({ ok: false, event_id: eventId(), error: message });
        setStatus(`定位失敗：${message}`, true);
        button.disabled = false;
        button.textContent = "📍 再試一次";
      },
      { enableHighAccuracy: true, timeout: 15000, maximumAge: 15000 }
    );
  });
  window.addEventListener("message", event => {
    if (!event.data || event.data.type !== "streamlit:render") return;
    button.disabled = Boolean(event.data.disabled);
    setHeight();
  });
  send("streamlit:componentReady", { apiVersion: API_VERSION });
  setHeight();
})();
</script>
</body>
</html>
"""


_DISPATCH_GEOLOCATION_COMPONENT = None


def get_dispatch_geolocation_component():
    """建立可由手機瀏覽器回傳目前經緯度的 Streamlit 雙向元件。"""
    global _DISPATCH_GEOLOCATION_COMPONENT
    if _DISPATCH_GEOLOCATION_COMPONENT is not None:
        return _DISPATCH_GEOLOCATION_COMPONENT

    component_dir = Path(tempfile.gettempdir()) / "dispatch_geolocation_component_v1"
    component_dir.mkdir(parents=True, exist_ok=True)
    index_path = component_dir / "index.html"
    try:
        if not index_path.exists() or index_path.read_text(encoding="utf-8") != DISPATCH_GEOLOCATION_COMPONENT_HTML:
            index_path.write_text(DISPATCH_GEOLOCATION_COMPONENT_HTML, encoding="utf-8")
    except OSError as exc:
        raise RuntimeError(f"無法建立定位元件：{exc}") from exc

    _DISPATCH_GEOLOCATION_COMPONENT = components.declare_component(
        "dispatch_geolocation_v1",
        path=str(component_dir),
    )
    return _DISPATCH_GEOLOCATION_COMPONENT


def normalize_coordinate(value, minimum: float, maximum: float) -> float | None:
    try:
        coordinate = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    if not math.isfinite(coordinate) or not minimum <= coordinate <= maximum:
        return None
    return coordinate


def build_youbike_station_location_map(
    base_df: pd.DataFrame,
    live_records: list[dict],
) -> dict[str, dict]:
    """把 Excel 場站名稱安全配對到官方站號及經緯度，供智慧調度計算距離。"""
    location_map: dict[str, dict] = {}
    for station_name in base_df["場站名稱"].astype(str).drop_duplicates():
        matched_record, _score, ambiguous = match_youbike_station(station_name, live_records)
        if matched_record is None or ambiguous:
            continue
        latitude = normalize_coordinate(matched_record.get("latitude"), -90.0, 90.0)
        longitude = normalize_coordinate(matched_record.get("longitude"), -180.0, 180.0)
        if latitude is None or longitude is None:
            continue
        location_map[station_name] = {
            "station_id": str(matched_record.get("station_id") or "").strip(),
            "official_name": str(matched_record.get("station_name") or station_name).strip(),
            "latitude": latitude,
            "longitude": longitude,
        }
    return location_map


def haversine_distance_km(
    origin_latitude: float,
    origin_longitude: float,
    destination_latitude: float,
    destination_longitude: float,
) -> float:
    """計算兩個座標間的大圓直線距離。"""
    earth_radius_km = 6371.0088
    lat1 = math.radians(origin_latitude)
    lat2 = math.radians(destination_latitude)
    delta_lat = math.radians(destination_latitude - origin_latitude)
    delta_lon = math.radians(destination_longitude - origin_longitude)
    value = (
        math.sin(delta_lat / 2) ** 2
        + math.cos(lat1) * math.cos(lat2) * math.sin(delta_lon / 2) ** 2
    )
    return earth_radius_km * 2 * math.atan2(math.sqrt(value), math.sqrt(max(0.0, 1 - value)))


def build_dispatch_plan_for_station(
    row: pd.Series,
    *,
    truck_bike: int,
    truck_ebike: int,
    max_capacity: int,
    global_bike_shortage: int,
    global_ebike_shortage: int,
) -> dict | None:
    """依車上存量及總容量，模擬本站先下車後上車的最大可行調度量。"""
    current_bike = normalize_current_status(row.get("2.0 現況"))
    current_ebike = normalize_current_status(row.get("2.0E 現況"))
    if current_bike is None or current_ebike is None:
        return None

    standard_bike = safe_nonnegative_int(row.get("2.0 標準"))
    standard_ebike = safe_nonnegative_int(row.get("2.0E 標準"))

    bike_shortage = max(0, standard_bike - current_bike)
    ebike_shortage = max(0, standard_ebike - current_ebike)
    bike_extra = max(0, current_bike - standard_bike)
    ebike_extra = max(0, current_ebike - standard_ebike)

    unload_bike = min(bike_shortage, truck_bike)
    unload_ebike = min(ebike_shortage, truck_ebike)
    bike_after_unload = truck_bike - unload_bike
    ebike_after_unload = truck_ebike - unload_ebike
    free_capacity = max(0, max_capacity - bike_after_unload - ebike_after_unload)

    pickup = {"bike": 0, "ebike": 0}
    pickup_needs = [
        ("bike", bike_extra, global_bike_shortage),
        ("ebike", ebike_extra, global_ebike_shortage),
    ]
    # 空間不足時，優先將整體缺口較大的車種上車帶走；仍以本站總調度量最大化為第一目標。
    pickup_needs.sort(key=lambda item: (item[2], item[1]), reverse=True)
    for vehicle_type, extra_count, _network_shortage in pickup_needs:
        amount = min(extra_count, free_capacity)
        pickup[vehicle_type] = amount
        free_capacity -= amount

    pickup_bike = pickup["bike"]
    pickup_ebike = pickup["ebike"]
    dispatch_count = unload_bike + unload_ebike + pickup_bike + pickup_ebike
    if dispatch_count <= 0:
        return None

    final_bike = bike_after_unload + pickup_bike
    final_ebike = ebike_after_unload + pickup_ebike
    if final_bike + final_ebike > max_capacity:
        return None

    return {
        "station_name": str(row.get("場站名稱") or "").strip(),
        "region": str(row.get("行政區") or "").strip(),
        "current_bike": current_bike,
        "current_ebike": current_ebike,
        "standard_bike": standard_bike,
        "standard_ebike": standard_ebike,
        "unload_bike": unload_bike,
        "unload_ebike": unload_ebike,
        "pickup_bike": pickup_bike,
        "pickup_ebike": pickup_ebike,
        "dispatch_count": dispatch_count,
        "truck_before_bike": safe_nonnegative_int(truck_bike),
        "truck_before_ebike": safe_nonnegative_int(truck_ebike),
        "truck_after_bike": final_bike,
        "truck_after_ebike": final_ebike,
        "max_capacity": safe_nonnegative_int(max_capacity),
    }


def calculate_dispatch_candidates(
    dispatch_df: pd.DataFrame,
    *,
    station_locations: dict[str, dict],
    current_location: dict,
    truck_bike: int,
    truck_ebike: int,
    max_capacity: int,
    cooldowns: dict[str, dict],
    rejection_history: list[dict],
    now_timestamp: float,
) -> list[dict]:
    """依可調度量、估計道路距離、行車時間及拒絕原因修正，建立候選排名。"""
    valid_df = dispatch_df.copy()
    current_bike_series = pd.to_numeric(valid_df["2.0 現況"], errors="coerce")
    current_ebike_series = pd.to_numeric(valid_df["2.0E 現況"], errors="coerce")
    standard_bike_series = pd.to_numeric(valid_df["2.0 標準"], errors="coerce").fillna(0)
    standard_ebike_series = pd.to_numeric(valid_df["2.0E 標準"], errors="coerce").fillna(0)
    global_bike_shortage = int((standard_bike_series - current_bike_series).clip(lower=0).fillna(0).sum())
    global_ebike_shortage = int((standard_ebike_series - current_ebike_series).clip(lower=0).fillna(0).sum())

    last_rejection_by_station: dict[str, dict] = {}
    one_hour_ago = now_timestamp - 3600
    for event in rejection_history:
        if event.get("action") != "rejected":
            continue
        if float(event.get("timestamp") or 0) < one_hour_ago:
            continue
        station_name = str(event.get("station_name") or "")
        if station_name:
            last_rejection_by_station[station_name] = event

    origin_lat = normalize_coordinate(current_location.get("latitude"), -90.0, 90.0)
    origin_lon = normalize_coordinate(current_location.get("longitude"), -180.0, 180.0)
    if origin_lat is None or origin_lon is None:
        return []

    candidates: list[dict] = []
    for _row_index, row in valid_df.iterrows():
        station_name = str(row.get("場站名稱") or "").strip()
        if not station_name:
            continue
        cooldown = cooldowns.get(station_name)
        if isinstance(cooldown, dict) and float(cooldown.get("until") or 0) > now_timestamp:
            continue
        location = station_locations.get(station_name)
        if not isinstance(location, dict):
            continue
        destination_lat = normalize_coordinate(location.get("latitude"), -90.0, 90.0)
        destination_lon = normalize_coordinate(location.get("longitude"), -180.0, 180.0)
        if destination_lat is None or destination_lon is None:
            continue

        plan = build_dispatch_plan_for_station(
            row,
            truck_bike=truck_bike,
            truck_ebike=truck_ebike,
            max_capacity=max_capacity,
            global_bike_shortage=global_bike_shortage,
            global_ebike_shortage=global_ebike_shortage,
        )
        if plan is None:
            continue

        straight_distance_km = haversine_distance_km(
            origin_lat,
            origin_lon,
            destination_lat,
            destination_lon,
        )
        estimated_road_distance_km = max(0.05, straight_distance_km * DISPATCH_ROAD_DISTANCE_FACTOR)
        estimated_drive_minutes = max(
            1.0,
            estimated_road_distance_km / DISPATCH_ESTIMATED_SPEED_KMH * 60,
        )
        estimated_operation_minutes = (
            DISPATCH_OPERATION_BASE_MINUTES
            + plan["dispatch_count"] * DISPATCH_OPERATION_MINUTES_PER_BIKE
        )
        estimated_total_minutes = estimated_drive_minutes + estimated_operation_minutes
        raw_efficiency = plan["dispatch_count"] / max(1.0, estimated_total_minutes)

        reason_multiplier = 1.0
        last_rejection = last_rejection_by_station.get(station_name)
        if last_rejection:
            reason_multiplier = DISPATCH_REASON_SCORE_MULTIPLIERS.get(
                str(last_rejection.get("reason") or ""),
                0.95,
            )
        score = raw_efficiency * reason_multiplier

        plan.update(
            {
                "station_id": str(location.get("station_id") or ""),
                "official_name": str(location.get("official_name") or station_name),
                "latitude": destination_lat,
                "longitude": destination_lon,
                "straight_distance_km": straight_distance_km,
                "estimated_distance_km": estimated_road_distance_km,
                "estimated_drive_minutes": estimated_drive_minutes,
                "estimated_operation_minutes": estimated_operation_minutes,
                "estimated_total_minutes": estimated_total_minutes,
                "raw_efficiency": raw_efficiency,
                "reason_multiplier": reason_multiplier,
                "score": score,
                "last_rejection_reason": str(last_rejection.get("reason") or "") if last_rejection else "",
            }
        )
        candidates.append(plan)

    return sorted(
        candidates,
        key=lambda item: (
            item["score"],
            item["dispatch_count"],
            -item["estimated_distance_km"],
        ),
        reverse=True,
    )


def dispatch_action_text(plan: dict) -> str:
    parts: list[str] = []
    if safe_nonnegative_int(plan.get("unload_bike")):
        parts.append(f"下車 2.0 × {safe_nonnegative_int(plan['unload_bike'])}")
    if safe_nonnegative_int(plan.get("unload_ebike")):
        parts.append(f"下車 2.0E × {safe_nonnegative_int(plan['unload_ebike'])}")
    if safe_nonnegative_int(plan.get("pickup_bike")):
        parts.append(f"上車 2.0 × {safe_nonnegative_int(plan['pickup_bike'])}")
    if safe_nonnegative_int(plan.get("pickup_ebike")):
        parts.append(f"上車 2.0E × {safe_nonnegative_int(plan['pickup_ebike'])}")
    return "｜".join(parts) if parts else "無可行調度"


def _dispatch_truck_before_counts(plan: dict) -> tuple[int, int]:
    """取得推薦前貨車載量；相容舊版已鎖定但尚未含 before 欄位的行程。"""
    before_bike = plan.get("truck_before_bike")
    before_ebike = plan.get("truck_before_ebike")
    if before_bike is None:
        before_bike = (
            safe_nonnegative_int(plan.get("truck_after_bike"))
            + safe_nonnegative_int(plan.get("unload_bike"))
            - safe_nonnegative_int(plan.get("pickup_bike"))
        )
    if before_ebike is None:
        before_ebike = (
            safe_nonnegative_int(plan.get("truck_after_ebike"))
            + safe_nonnegative_int(plan.get("unload_ebike"))
            - safe_nonnegative_int(plan.get("pickup_ebike"))
        )
    return max(0, int(before_bike)), max(0, int(before_ebike))


def render_dispatch_plan_card(plan: dict, *, title: str) -> None:
    """呈現單一推薦場站，並用作業前後對照快速確認貨車載量。"""
    before_bike, before_ebike = _dispatch_truck_before_counts(plan)
    after_bike = safe_nonnegative_int(plan.get("truck_after_bike"))
    after_ebike = safe_nonnegative_int(plan.get("truck_after_ebike"))
    max_capacity = max(
        1,
        safe_nonnegative_int(plan.get("max_capacity"))
        or before_bike + before_ebike
        or after_bike + after_ebike
        or 1,
    )
    before_total = before_bike + before_ebike
    after_total = after_bike + after_ebike
    after_free = max(0, max_capacity - after_total)

    multiplier_text = ""
    if float(plan.get("reason_multiplier") or 1.0) < 0.999:
        multiplier_text = (
            f'<div class="dispatch-plan-note">⚠️ 曾因「'
            f'{html.escape(str(plan.get("last_rejection_reason") or "其他原因"))}」跳過，'
            '本次效益已納入原因修正。</div>'
        )

    st.markdown(
        f"""
        <section class="dispatch-plan-card">
          <div class="dispatch-plan-header">
            <div>
              <div class="dispatch-plan-kicker">{html.escape(title)}</div>
              <div class="dispatch-plan-title">{html.escape(str(plan['station_name']))}</div>
              <div class="dispatch-plan-region">{html.escape(str(plan.get('region') or ''))}</div>
            </div>
            <div class="dispatch-plan-badge">可調度 <strong>{safe_nonnegative_int(plan['dispatch_count'])}</strong> 台</div>
          </div>

          <div class="dispatch-plan-grid">
            <div><span>道路距離</span><strong>{float(plan['estimated_distance_km']):.1f} km</strong></div>
            <div><span>行車時間</span><strong>{float(plan['estimated_drive_minutes']):.0f} 分</strong></div>
            <div><span>預估總時間</span><strong>{float(plan['estimated_total_minutes']):.0f} 分</strong></div>
            <div><span>綜合效益</span><strong>{float(plan['score']):.2f}</strong><small>台／分</small></div>
          </div>

          <div class="dispatch-plan-action-label">本站作業</div>
          <div class="dispatch-plan-action">{html.escape(dispatch_action_text(plan))}</div>

          <div class="dispatch-truck-compare">
            <div class="dispatch-truck-row dispatch-truck-before">
              <span>目前車上數量</span>
              <strong>2.0＝{before_bike} 台｜2.0E＝{before_ebike} 台</strong>
              <small>合計 {before_total}／{max_capacity} 台</small>
            </div>
            <div class="dispatch-truck-arrow" aria-hidden="true">↓</div>
            <div class="dispatch-truck-row dispatch-truck-after">
              <span>完成後貨車</span>
              <strong>2.0＝{after_bike} 台｜2.0E＝{after_ebike} 台</strong>
              <small>合計 {after_total}／{max_capacity} 台・剩餘 {after_free} 格</small>
            </div>
          </div>
          {multiplier_text}
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_dispatch_truck_status(
    *,
    truck_bike: int,
    truck_ebike: int,
    max_capacity: int,
    locked: bool,
) -> None:
    """在調度區頂端顯示精簡貨車狀態，避免反覆打開設定確認。"""
    total = truck_bike + truck_ebike
    remaining = max(0, max_capacity - total)
    load_percent = min(100.0, total / max(1, max_capacity) * 100)
    lock_text = "🔒 行程中已鎖定" if locked else "可調整"
    st.markdown(
        f"""
        <section class="dispatch-truck-status">
          <div class="dispatch-truck-status-head">
            <span>🚚 目前貨車</span><small>{lock_text}</small>
          </div>
          <div class="dispatch-truck-status-grid">
            <div><span>2.0</span><strong>{truck_bike}<small>台</small></strong></div>
            <div><span>2.0E</span><strong>{truck_ebike}<small>台</small></strong></div>
            <div><span>合計</span><strong>{total}<small>／{max_capacity}</small></strong></div>
            <div><span>剩餘空位</span><strong>{remaining}<small>格</small></strong></div>
          </div>
          <div class="dispatch-load-track"><i style="width:{load_percent:.1f}%"></i></div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_dispatch_auxiliary_panels(
    *,
    dispatch_prefix: str,
    cooldown_key: str,
    cooldowns: dict[str, dict],
    history: list[dict],
    now_timestamp: float,
) -> None:
    """把冷卻與紀錄移到主要決策區後方，讓下一站卡片優先出現。"""
    if cooldowns:
        cooldown_lines = []
        for station_name, data in sorted(
            cooldowns.items(),
            key=lambda item: float(item[1].get("until") or 0),
        ):
            remaining_seconds = max(0, int(float(data.get("until") or 0) - now_timestamp))
            minutes, seconds = divmod(remaining_seconds, 60)
            cooldown_lines.append(
                f"{station_name}：{minutes:02d}:{seconds:02d}（{data.get('reason') or '未填原因'}）"
            )
        with st.expander(f"冷卻中的場站（{len(cooldowns)}）", expanded=False):
            st.write("  \n".join(cooldown_lines))
            if st.button(
                "清除全部冷卻",
                key=f"{dispatch_prefix}::clear_cooldowns",
                use_container_width=True,
            ):
                st.session_state[cooldown_key] = {}
                rerun_app()

    if history:
        with st.expander(f"本班次調度紀錄（{len(history)}）", expanded=False):
            history_rows = []
            for event in reversed(history[-30:]):
                event_time = datetime.fromtimestamp(
                    float(event.get("timestamp") or now_timestamp),
                    TAIPEI_TIMEZONE,
                ).strftime("%H:%M:%S")
                if event.get("action") == "rejected":
                    detail = str(event.get("reason") or "未填原因")
                    if event.get("note"):
                        detail += f"｜{event['note']}"
                    result_text = "不前往／冷卻10分鐘"
                else:
                    detail = (
                        f"下車2.0 {safe_nonnegative_int(event.get('unload_bike'))}、"
                        f"上車2.0 {safe_nonnegative_int(event.get('pickup_bike'))}、"
                        f"下車2.0E {safe_nonnegative_int(event.get('unload_ebike'))}、"
                        f"上車2.0E {safe_nonnegative_int(event.get('pickup_ebike'))}"
                    )
                    result_text = "已完成"
                history_rows.append(
                    {
                        "時間": event_time,
                        "場站": event.get("station_name", ""),
                        "結果": result_text,
                        "原因／作業": detail,
                    }
                )
            st.dataframe(pd.DataFrame(history_rows), hide_index=True, use_container_width=True)

def render_smart_dispatch(
    *,
    full_status_df: pd.DataFrame,
    selected_region: str,
    status_cache: dict,
    current_context_key: str,
    active_base: dict,
) -> None:
    """智慧動態調度：逐站詢問、10分鐘冷卻、載量限制、現場變數及完成後重算。"""
    st.markdown("---")
    st.markdown('<div id="smart-dispatch-anchor"></div>', unsafe_allow_html=True)
    st.subheader("智慧動態調度")
    st.caption(
        "每次只安排一站；同意後鎖定目的地，完成本站才重新計算。距離與時間為官方座標搭配道路係數的估算值，導航請以地圖實際路況為準。"
    )

    st.markdown(
        """
        <style>
        .dispatch-plan-card {
            border: 1px solid rgba(22,119,255,.25); border-radius: 20px;
            padding: 1rem; margin: .45rem 0 .7rem;
            background: linear-gradient(145deg, rgba(22,119,255,.085), rgba(16,185,129,.055));
            box-shadow: 0 10px 28px rgba(15,23,42,.07);
        }
        .dispatch-plan-header {display:flex; justify-content:space-between; align-items:flex-start; gap:.8rem;}
        .dispatch-plan-kicker {font-size:.76rem; font-weight:850; letter-spacing:.06em; color:#1677ff;}
        .dispatch-plan-title {font-size:1.55rem; line-height:1.2; font-weight:900; margin-top:.18rem;}
        .dispatch-plan-region {font-size:.8rem; opacity:.7; margin-top:.25rem;}
        .dispatch-plan-badge {flex:0 0 auto; padding:.42rem .62rem; border-radius:999px; background:rgba(22,119,255,.12); color:#0b63ce; font-size:.72rem; font-weight:750; white-space:nowrap;}
        .dispatch-plan-badge strong {font-size:1.02rem;}
        .dispatch-plan-grid {display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:.42rem; margin-top:.78rem;}
        .dispatch-plan-grid div {padding:.58rem .38rem; border-radius:12px; background:rgba(255,255,255,.76); text-align:center; min-width:0;}
        .dispatch-plan-grid span {display:block; font-size:.66rem; opacity:.66; white-space:nowrap;}
        .dispatch-plan-grid strong {display:inline-block; font-size:.98rem; margin-top:.16rem;}
        .dispatch-plan-grid small {font-size:.62rem; margin-left:.12rem; opacity:.7;}
        .dispatch-plan-action-label {margin-top:.68rem; font-size:.67rem; font-weight:800; opacity:.66;}
        .dispatch-plan-action {margin-top:.22rem; padding:.72rem .78rem; border-radius:12px; font-size:1.02rem; font-weight:900; background:rgba(22,119,255,.12);}
        .dispatch-truck-compare {margin-top:.58rem; padding:.55rem; border-radius:14px; background:rgba(255,255,255,.56);}
        .dispatch-truck-row {display:grid; grid-template-columns:6.2rem 1fr auto; gap:.45rem; align-items:center; padding:.44rem .5rem; border-radius:10px;}
        .dispatch-truck-row span {font-size:.72rem; font-weight:800; opacity:.68;}
        .dispatch-truck-row strong {font-size:.82rem;}
        .dispatch-truck-row small {font-size:.66rem; opacity:.66; white-space:nowrap;}
        .dispatch-truck-before {background:rgba(148,163,184,.08);}
        .dispatch-truck-after {background:rgba(16,185,129,.11);}
        .dispatch-truck-arrow {height:.55rem; line-height:.55rem; text-align:center; font-size:.7rem; opacity:.5;}
        .dispatch-plan-note {margin-top:.48rem; padding:.48rem .58rem; border-radius:10px; font-size:.72rem; line-height:1.45; background:rgba(245,158,11,.1);}
        .dispatch-truck-status {margin:.35rem 0 .7rem; padding:.72rem .78rem; border:1px solid rgba(148,163,184,.22); border-radius:16px; background:rgba(148,163,184,.06);}
        .dispatch-truck-status-head {display:flex; justify-content:space-between; align-items:center; margin-bottom:.48rem;}
        .dispatch-truck-status-head span {font-weight:850; font-size:.84rem;}
        .dispatch-truck-status-head small {font-size:.66rem; opacity:.62;}
        .dispatch-truck-status-grid {display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:.35rem;}
        .dispatch-truck-status-grid div {padding:.38rem .3rem; text-align:center; border-radius:10px; background:rgba(255,255,255,.68);}
        .dispatch-truck-status-grid span {display:block; font-size:.63rem; opacity:.62;}
        .dispatch-truck-status-grid strong {font-size:.98rem;}
        .dispatch-truck-status-grid strong small {font-size:.62rem; margin-left:.08rem; opacity:.7;}
        .dispatch-load-track {height:5px; margin-top:.5rem; border-radius:999px; overflow:hidden; background:rgba(148,163,184,.24);}
        .dispatch-load-track i {display:block; height:100%; border-radius:999px; background:linear-gradient(90deg,#1677ff,#10b981);}
        @media (max-width: 700px) {
          .dispatch-plan-card {padding:.82rem; border-radius:18px;}
          .dispatch-plan-title {font-size:1.42rem;}
          .dispatch-plan-grid {grid-template-columns:repeat(2,minmax(0,1fr));}
          .dispatch-truck-row {grid-template-columns:5.7rem 1fr; gap:.2rem .4rem;}
          .dispatch-truck-row small {grid-column:2; white-space:normal;}
          .dispatch-truck-status {padding:.65rem;}
          .dispatch-truck-status-grid {grid-template-columns:repeat(4,minmax(0,1fr)); gap:.24rem;}
          .dispatch-truck-status-grid strong {font-size:.9rem;}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    dispatch_prefix = f"smart_dispatch::{active_base['token']}::{current_context_key}"
    location_state_key = f"{dispatch_prefix}::location"
    active_trip_key = f"{dispatch_prefix}::active_trip"
    cooldown_key = f"{dispatch_prefix}::cooldowns"
    history_key = f"{dispatch_prefix}::history"

    st.session_state.setdefault(cooldown_key, {})
    st.session_state.setdefault(history_key, [])
    pending_truck_key = f"{dispatch_prefix}::pending_truck_counts"
    pending_truck_counts = st.session_state.pop(pending_truck_key, None)
    if isinstance(pending_truck_counts, dict):
        st.session_state[f"{dispatch_prefix}::truck_bike"] = safe_nonnegative_int(
            pending_truck_counts.get("bike")
        )
        st.session_state[f"{dispatch_prefix}::truck_ebike"] = safe_nonnegative_int(
            pending_truck_counts.get("ebike")
        )
    now_timestamp = time.time()

    # 移除已滿10分鐘的冷卻，讓場站重新加入候選評估。
    cooldowns = dict(st.session_state.get(cooldown_key, {}))
    cooldowns = {
        station_name: data
        for station_name, data in cooldowns.items()
        if isinstance(data, dict) and float(data.get("until") or 0) > now_timestamp
    }
    st.session_state[cooldown_key] = cooldowns
    history = list(st.session_state.get(history_key, []))

    active_trip = st.session_state.get(active_trip_key)
    trip_locked = isinstance(active_trip, dict)
    max_capacity_input_key = f"{dispatch_prefix}::max_capacity"
    truck_bike_input_key = f"{dispatch_prefix}::truck_bike"
    truck_ebike_input_key = f"{dispatch_prefix}::truck_ebike"
    has_saved_truck_settings = all(
        key in st.session_state
        for key in (max_capacity_input_key, truck_bike_input_key, truck_ebike_input_key)
    )

    settings_title = "🚚 貨車載量設定" + ("（前往中已鎖定）" if trip_locked else "")
    with st.expander(settings_title, expanded=not has_saved_truck_settings and not trip_locked):
        control_col_1, control_col_2, control_col_3 = st.columns(3)
        with control_col_1:
            max_capacity = int(st.number_input(
                "最高載量",
                min_value=1,
                max_value=100,
                value=14,
                step=1,
                key=max_capacity_input_key,
                disabled=trip_locked,
            ))
        with control_col_2:
            truck_bike = int(st.number_input(
                "車上 2.0",
                min_value=0,
                max_value=100,
                value=0,
                step=1,
                key=truck_bike_input_key,
                disabled=trip_locked,
            ))
        with control_col_3:
            truck_ebike = int(st.number_input(
                "車上 2.0E",
                min_value=0,
                max_value=100,
                value=0,
                step=1,
                key=truck_ebike_input_key,
                disabled=trip_locked,
            ))

    total_on_truck = truck_bike + truck_ebike
    render_dispatch_truck_status(
        truck_bike=truck_bike,
        truck_ebike=truck_ebike,
        max_capacity=max_capacity,
        locked=trip_locked,
    )
    if total_on_truck > max_capacity:
        st.error(
            f"目前車上合計 {total_on_truck} 台，已超過最高載量 {max_capacity} 台。請先修正數量，系統不會安排路線。"
        )
        return

    location_payload = None
    try:
        geolocation_component = get_dispatch_geolocation_component()
        location_payload = geolocation_component(
            key=f"dispatch_geolocation::{dispatch_prefix}",
            default=None,
        )
    except Exception as exc:
        st.error(f"定位功能建立失敗：{exc}")

    if isinstance(location_payload, dict):
        location_event_id = str(location_payload.get("event_id") or "").strip()
        processed_location_event_key = f"{dispatch_prefix}::processed_location_event"
        if location_event_id and st.session_state.get(processed_location_event_key) != location_event_id:
            st.session_state[processed_location_event_key] = location_event_id
            if location_payload.get("ok"):
                latitude = normalize_coordinate(location_payload.get("latitude"), -90.0, 90.0)
                longitude = normalize_coordinate(location_payload.get("longitude"), -180.0, 180.0)
                if latitude is not None and longitude is not None:
                    st.session_state[location_state_key] = {
                        "latitude": latitude,
                        "longitude": longitude,
                        "accuracy": max(0.0, float(location_payload.get("accuracy") or 0)),
                        "updated_at": now_timestamp,
                        "source": "gps",
                    }
                    st.success("目前位置已更新，下一站排名已重新計算。")
            else:
                st.warning(f"目前位置尚未更新：{location_payload.get('error') or '定位失敗'}")

    current_location = st.session_state.get(location_state_key)
    if isinstance(current_location, dict):
        source_text = "GPS定位" if current_location.get("source") == "gps" else "上一個完成場站"
        accuracy = float(current_location.get("accuracy") or 0)
        accuracy_text = f"｜誤差約 {accuracy:.0f} 公尺" if accuracy and current_location.get("source") == "gps" else ""
        st.caption(
            f"計算起點：{source_text}｜{float(current_location['latitude']):.6f}, "
            f"{float(current_location['longitude']):.6f}{accuracy_text}"
        )
    else:
        st.info("請先按「取得／更新目前位置」，系統才能把距離與行車時間納入下一站評估。")
        return

    metadata = status_cache.get("metadata", {}).get(current_context_key, {})
    station_locations = metadata.get("station_locations", {}) if isinstance(metadata, dict) else {}
    if not isinstance(station_locations, dict) or not station_locations:
        st.info("尚未取得場站官方座標。請先在上方執行一次「高速取得全部 YouBike 場站車數」。")
        return


    dispatch_df = full_status_df.copy()
    if selected_region != "全部":
        dispatch_df = dispatch_df[dispatch_df["行政區"] == selected_region].copy()

    if isinstance(active_trip, dict):
        trip_id = str(active_trip.get("trip_id") or normalize_station_key(active_trip.get("station_name")))
        render_dispatch_plan_card(active_trip, title="已同意前往／目的地已鎖定")
        maps_query = urlencode(
            {
                "api": 1,
                "destination": f"{active_trip['latitude']},{active_trip['longitude']}",
                "travelmode": "driving",
            }
        )
        st.markdown(f"[🧭 開啟 Google Maps 導航](https://www.google.com/maps/dir/?{maps_query})")
        st.info("到場作業後，可依現場變數修改實際上下車數量；只有按下完成本站，系統才會安排下一站。")

        with st.expander("現場變數／修改實際上下車數量", expanded=True):
            with st.form(key=f"{dispatch_prefix}::complete_trip_form::{trip_id}", clear_on_submit=False):
                action_col_1, action_col_2 = st.columns(2)
                with action_col_1:
                    actual_unload_bike = int(st.number_input(
                        "實際下車 2.0",
                        min_value=0,
                        value=safe_nonnegative_int(active_trip.get("unload_bike")),
                        step=1,
                        key=f"{dispatch_prefix}::actual_unload_bike::{trip_id}",
                    ))
                    actual_pickup_bike = int(st.number_input(
                        "實際上車 2.0",
                        min_value=0,
                        value=safe_nonnegative_int(active_trip.get("pickup_bike")),
                        step=1,
                        key=f"{dispatch_prefix}::actual_pickup_bike::{trip_id}",
                    ))
                with action_col_2:
                    actual_unload_ebike = int(st.number_input(
                        "實際下車 2.0E",
                        min_value=0,
                        value=safe_nonnegative_int(active_trip.get("unload_ebike")),
                        step=1,
                        key=f"{dispatch_prefix}::actual_unload_ebike::{trip_id}",
                    ))
                    actual_pickup_ebike = int(st.number_input(
                        "實際上車 2.0E",
                        min_value=0,
                        value=safe_nonnegative_int(active_trip.get("pickup_ebike")),
                        step=1,
                        key=f"{dispatch_prefix}::actual_pickup_ebike::{trip_id}",
                    ))

                completed = st.form_submit_button(
                    "✅ 完成本站並安排下一站",
                    type="primary",
                    use_container_width=True,
                )

            if completed:
                operation_start_bike, operation_start_ebike = _dispatch_truck_before_counts(active_trip)
                final_truck_bike = operation_start_bike - actual_unload_bike + actual_pickup_bike
                final_truck_ebike = operation_start_ebike - actual_unload_ebike + actual_pickup_ebike
                final_total = final_truck_bike + final_truck_ebike
                station_final_bike = (
                    safe_nonnegative_int(active_trip.get("current_bike"))
                    + actual_unload_bike
                    - actual_pickup_bike
                )
                station_final_ebike = (
                    safe_nonnegative_int(active_trip.get("current_ebike"))
                    + actual_unload_ebike
                    - actual_pickup_ebike
                )

                errors: list[str] = []
                if actual_unload_bike > operation_start_bike:
                    errors.append("實際下車的 2.0 超過車上現有數量")
                if actual_unload_ebike > operation_start_ebike:
                    errors.append("實際下車的 2.0E 超過車上現有數量")
                if final_truck_bike < 0 or final_truck_ebike < 0:
                    errors.append("作業後車上數量不可為負數")
                if final_total > max_capacity:
                    errors.append(f"作業後合計 {final_total} 台，超過最高載量 {max_capacity} 台")
                if station_final_bike < 0 or station_final_ebike < 0:
                    errors.append("實際上車數量超過本站作業前可用車數")

                if errors:
                    for error in errors:
                        st.error(error)
                else:
                    updated_full_df = full_status_df.copy()
                    station_mask = (
                        updated_full_df["場站名稱"].astype(str).eq(str(active_trip["station_name"]))
                        & updated_full_df["行政區"].astype(str).eq(str(active_trip.get("region") or ""))
                    )
                    if not station_mask.any():
                        st.error("找不到目前目的地在配置表中的資料，未寫入完成結果。")
                    else:
                        updated_full_df.loc[station_mask, "2.0 現況"] = station_final_bike
                        updated_full_df.loc[station_mask, "2.0E 現況"] = station_final_ebike
                        status_cache["contexts"][current_context_key] = dataframe_to_status_records(updated_full_df)
                        save_cached_status(
                            active_base["token"],
                            active_base["expires_at"],
                            status_cache,
                        )
                        st.session_state[pending_truck_key] = {
                            "bike": final_truck_bike,
                            "ebike": final_truck_ebike,
                        }
                        st.session_state[location_state_key] = {
                            "latitude": float(active_trip["latitude"]),
                            "longitude": float(active_trip["longitude"]),
                            "accuracy": 0.0,
                            "updated_at": time.time(),
                            "source": "completed_station",
                        }
                        history.append(
                            {
                                "action": "completed",
                                "station_name": active_trip["station_name"],
                                "timestamp": time.time(),
                                "unload_bike": actual_unload_bike,
                                "unload_ebike": actual_unload_ebike,
                                "pickup_bike": actual_pickup_bike,
                                "pickup_ebike": actual_pickup_ebike,
                            }
                        )
                        st.session_state[history_key] = history[-100:]
                        st.session_state.pop(active_trip_key, None)
                        clear_editor_session_state(active_base["token"])
                        rerun_app()
        render_dispatch_auxiliary_panels(
            dispatch_prefix=dispatch_prefix,
            cooldown_key=cooldown_key,
            cooldowns=cooldowns,
            history=history,
            now_timestamp=now_timestamp,
        )
        return

    candidates = calculate_dispatch_candidates(
        dispatch_df,
        station_locations=station_locations,
        current_location=current_location,
        truck_bike=truck_bike,
        truck_ebike=truck_ebike,
        max_capacity=max_capacity,
        cooldowns=cooldowns,
        rejection_history=history,
        now_timestamp=now_timestamp,
    )
    if not candidates:
        st.warning(
            "目前找不到可執行的下一站。可能原因：全部符合配置、車上車種不足、貨車已滿、場站仍在冷卻，或部分場站缺少座標／現況資料。"
        )
        render_dispatch_auxiliary_panels(
            dispatch_prefix=dispatch_prefix,
            cooldown_key=cooldown_key,
            cooldowns=cooldowns,
            history=history,
            now_timestamp=now_timestamp,
        )
        return

    recommended = candidates[0]
    render_dispatch_plan_card(recommended, title="下一站最高效益推薦")
    if len(candidates) > 1:
        with st.expander("查看其他候選場站", expanded=False):
            candidate_rows = []
            for rank, candidate in enumerate(candidates[1:6], start=2):
                candidate_rows.append(
                    {
                        "順位": rank,
                        "場站": candidate["station_name"],
                        "調度量": candidate["dispatch_count"],
                        "距離(km)": round(candidate["estimated_distance_km"], 1),
                        "總時間(分)": round(candidate["estimated_total_minutes"]),
                        "效益(台/分)": round(candidate["score"], 2),
                    }
                )
            st.dataframe(pd.DataFrame(candidate_rows), hide_index=True, use_container_width=True)

    with st.form(key=f"{dispatch_prefix}::recommendation_decision", clear_on_submit=False):
        rejection_reason = st.selectbox(
            "若不前往，請選擇原因",
            DISPATCH_REJECTION_REASONS,
            key=f"{dispatch_prefix}::rejection_reason",
        )
        rejection_note = st.text_input(
            "補充說明（選填）",
            placeholder="例如：入口施工、臨停位置被占用、現場數據差異……",
            key=f"{dispatch_prefix}::rejection_note",
        )
        decision_col_1, decision_col_2 = st.columns(2)
        with decision_col_1:
            accepted = st.form_submit_button(
                "✅ 前往此站",
                type="primary",
                use_container_width=True,
            )
        with decision_col_2:
            rejected = st.form_submit_button(
                "⏭️ 跳過並找下一站",
                use_container_width=True,
            )

    if accepted:
        locked_trip = dict(recommended)
        locked_trip["trip_id"] = uuid.uuid4().hex
        locked_trip["accepted_at"] = time.time()
        st.session_state[active_trip_key] = locked_trip
        rerun_app()

    if rejected:
        rejected_at = time.time()
        cooldowns[recommended["station_name"]] = {
            "until": rejected_at + DISPATCH_COOLDOWN_SECONDS,
            "reason": rejection_reason,
            "note": rejection_note.strip(),
            "rejected_at": rejected_at,
        }
        st.session_state[cooldown_key] = cooldowns
        history.append(
            {
                "action": "rejected",
                "station_name": recommended["station_name"],
                "timestamp": rejected_at,
                "reason": rejection_reason,
                "note": rejection_note.strip(),
                "score_before_rejection": recommended["score"],
                "dispatch_count": recommended["dispatch_count"],
            }
        )
        st.session_state[history_key] = history[-100:]
        rerun_app()

    render_dispatch_auxiliary_panels(
        dispatch_prefix=dispatch_prefix,
        cooldown_key=cooldown_key,
        cooldowns=cooldowns,
        history=history,
        now_timestamp=now_timestamp,
    )


@st.cache_data(show_spinner=False)
def cached_load_workbook(source: bytes | str) -> dict[str, pd.DataFrame]:
    """只讀取 Excel 中可見的工作表；hidden 與 veryHidden 一律排除。"""
    metadata_source = BytesIO(source) if isinstance(source, bytes) else source
    metadata_book = openpyxl_load_workbook(
        metadata_source,
        read_only=True,
        data_only=False,
    )
    try:
        visible_sheet_names = [
            worksheet.title
            for worksheet in metadata_book.worksheets
            if worksheet.sheet_state == "visible"
        ]
    finally:
        metadata_book.close()

    excel_source = BytesIO(source) if isinstance(source, bytes) else source
    with pd.ExcelFile(excel_source, engine="openpyxl") as book:
        available_visible_names = [
            sheet_name
            for sheet_name in visible_sheet_names
            if sheet_name in book.sheet_names
        ]
        return {
            sheet_name: pd.read_excel(
                book,
                sheet_name=sheet_name,
                header=None,
                dtype=object,
                engine="openpyxl",
            )
            for sheet_name in available_visible_names
        }


@st.cache_data(show_spinner=False, max_entries=48)
def cached_parse_route(source: bytes | str, sheet_name: str, route: str, shift: str) -> pd.DataFrame:
    """快取已解析的分區資料，避免每次輸入都重新掃描 Excel 工作表。"""
    workbook = cached_load_workbook(source)
    return parse_route(workbook[sheet_name], route, shift)


st.markdown(
    '<div id="ubike-page-top-anchor" aria-hidden="true"></div>',
    unsafe_allow_html=True,
)
st.title("臺東 YouBike 調度")

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
    st.header("配置")

    if st.session_state.pop("base_expired_notice", False):
        st.warning("原配置無法讀取，請重新上傳。")

    uploaded_excel = st.file_uploader(
        "配置表",
        type=["xlsx"],
        key=f"base_uploader_{st.session_state['base_uploader_version']}",
    )
    mobile_input_mode = st.checkbox(
        "手機數字輸入",
        value=mobile_detected,
    )

active_base = cached_base

if uploaded_excel is not None:
    uploaded_bytes = uploaded_excel.getvalue()
    uploaded_digest = hashlib.sha256(uploaded_bytes).hexdigest()

    if active_base is None or active_base["sha256"] != uploaded_digest:
        active_base = save_cached_base(uploaded_excel.name, uploaded_bytes)

if st.session_state.pop("full_reset_notice", False):
    st.success("已全部重置。")

if st.session_state.pop("data_zero_notice", False):
    st.success("現況已清空。")

if active_base is None:
    st.info("請上傳配置表。")
    st.stop()

try:
    workbook_data = cached_load_workbook(active_base["bytes"])
except Exception as exc:
    st.error(f"Excel 讀取失敗：{exc}")
    st.stop()

options = available_sources(workbook_data)
if not options:
    st.error("找不到可使用的 D1／D2／D3 配置。")
    st.stop()

with st.sidebar:
    st.caption(active_base["name"])
    with st.expander("資料管理", expanded=False):
        if st.button(
            "清空現況",
            use_container_width=True,
        ):
            blank_status_cache = build_blank_status_cache(workbook_data, options)
            save_cached_status(
                active_base["token"],
                active_base["expires_at"],
                blank_status_cache,
            )
            clear_editor_session_state(active_base["token"])
            st.session_state["data_zero_notice"] = True
            rerun_app()

        if st.button(
            "全部重置",
            use_container_width=True,
            type="primary",
        ):
            reset_token = active_base.get("token")
            if reset_token:
                delete_cached_base(reset_token)
            clear_base_token()
            clear_editor_session_state(reset_token)
            st.session_state["base_uploader_version"] += 1
            st.session_state["full_reset_notice"] = True
            rerun_app()

option_labels = [f"{route}｜{sheet_name}" for sheet_name, route in options]
selected_label = st.selectbox(
    "配置版本",
    option_labels,
    index=None,
    placeholder="選擇配置版本",
    key=f"config_version::{active_base['token']}",
)

# 未主動選擇配置前，不解析、不恢復現況，也不執行任何同步或辨識。
if selected_label is None:
    st.stop()

selector_1, selector_2 = st.columns([1, 1])
with selector_1:
    selected_shift = st.selectbox("班別", list(SHIFT_COLUMNS.keys()))
with selector_2:
    selected_region_placeholder = st.empty()

selected_index = option_labels.index(selected_label)
selected_sheet, selected_route = options[selected_index]
base_df = cached_parse_route(active_base["bytes"], selected_sheet, selected_route, selected_shift)

status_cache = load_cached_status(active_base["token"], active_base["expires_at"])
current_context_key = status_context_key(selected_sheet, selected_route, selected_shift)
saved_current_status = status_cache["contexts"].get(current_context_key)
base_df = blank_current_status(base_df)
if saved_current_status is not None:
    base_df = restore_current_status(base_df, saved_current_status)

if base_df.empty:
    st.warning("此配置沒有可用場站。")
    st.stop()


with st.expander("YouBike 即時車數", expanded=True):
    previous_live_meta = status_cache.get("metadata", {}).get(current_context_key, {})
    if isinstance(previous_live_meta, dict) and previous_live_meta.get("fetched_at"):
        previous_source_time = str(previous_live_meta.get("latest_source_time") or "").strip()
        source_time_text = f"｜官方資料時間 {previous_source_time}" if previous_source_time else ""
        st.caption(
            f"上次同步：{previous_live_meta['fetched_at']}"
            f"{source_time_text}｜{safe_nonnegative_int(previous_live_meta.get('matched_count'))} 站"
        )

    st.caption(
        "即時數據預設每 1 分鐘自動更新一次；右側懸浮「更新」按鈕可隨時手動更新。"
        "目的地一旦同意前往會保持鎖定，不會因即時數據變動自行換站。"
    )

    browser_payload = None
    try:
        browser_sync_component = get_youbike_browser_sync_component()
        browser_payload = browser_sync_component(
            catalog_url=YOUBIKE_STATION_CATALOG_URL,
            parking_url=YOUBIKE_PARKING_INFO_URL,
            # 第一輪每次最多查 20 站，最多 4 個請求並行；後續只重查漏站。
            batch_size=20,
            request_concurrency=4,
            max_batch_rounds=8,
            max_single_rounds=3,
            wave_delay_ms=70,
            button_label="🔄 手動更新即時車數",
            auto_refresh=True,
            auto_refresh_seconds=60,
            key=f"browser_youbike_sync::{current_context_key}",
            default=None,
        )
    except YouBikeDataError as exc:
        st.error(f"瀏覽器同步元件建立失敗：{exc}")
    except Exception as exc:
        st.error(f"瀏覽器同步元件發生未預期錯誤：{exc}")

    if isinstance(browser_payload, dict):
        browser_event_id = str(browser_payload.get("event_id") or "").strip()
        processed_event_key = f"processed_browser_youbike_event::{current_context_key}"
        already_processed = bool(
            browser_event_id
            and st.session_state.get(processed_event_key) == browser_event_id
        )

        if not already_processed:
            try:
                if browser_event_id:
                    # 先登記事件，避免 Streamlit 元件保留上次回傳值時重複寫入。
                    st.session_state[processed_event_key] = browser_event_id

                with st.spinner("正在配對臺東場站名稱並寫入 2.0／2.0E 現況……"):
                    live_payload = normalize_browser_live_payload(browser_payload)
                    previous_location_map = (
                        dict(previous_live_meta.get("station_locations", {}))
                        if isinstance(previous_live_meta, dict)
                        and isinstance(previous_live_meta.get("station_locations"), dict)
                        else {}
                    )
                    previous_location_map.update(
                        build_youbike_station_location_map(base_df, live_payload["records"])
                    )
                    live_updated_df, live_report_df, live_summary = apply_youbike_updates_to_dataframe(
                        base_df,
                        live_payload["records"],
                    )

                    if live_summary["matched_count"] <= 0:
                        st.error("沒有任何場站通過安全配對，因此未修改現況資料。")
                    else:
                        base_df = live_updated_df
                        status_cache["contexts"][current_context_key] = dataframe_to_status_records(base_df)
                        status_cache.setdefault("metadata", {})[current_context_key] = {
                            "source": live_payload.get("source", "YouBike 官網公開接口（瀏覽器直連，免 TDX）"),
                            "fetched_at": live_payload["fetched_at"],
                            "latest_source_time": live_payload.get("latest_source_time", ""),
                            "matched_count": live_summary["matched_count"],
                            "skipped_count": live_summary["skipped_count"],
                            "unmatched_count": live_summary["unmatched_count"],
                            "station_count": live_payload.get("station_count", 0),
                            "requested_station_count": live_payload.get("requested_station_count", 0),
                            "missing_station_count": live_payload.get("missing_station_count", 0),
                            "request_count": live_payload.get("request_count", 0),
                            "batch_round_count": live_payload.get("batch_round_count", 0),
                            "single_round_count": live_payload.get("single_round_count", 0),
                            "station_locations": previous_location_map,
                        }
                        save_cached_status(
                            active_base["token"],
                            active_base["expires_at"],
                            status_cache,
                        )
                        clear_editor_session_state(active_base["token"])
                        official_time = str(live_payload.get("latest_source_time") or "").strip()
                        official_time_text = f"｜官方資料時間：{official_time}" if official_time else ""
                        returned_count = safe_nonnegative_int(live_payload.get("station_count"))
                        requested_count = safe_nonnegative_int(live_payload.get("requested_station_count"))
                        missing_count = safe_nonnegative_int(live_payload.get("missing_station_count"))
                        request_count = safe_nonnegative_int(live_payload.get("request_count"))
                        failed_request_count = safe_nonnegative_int(live_payload.get("failed_request_count"))
                        batch_round_count = safe_nonnegative_int(live_payload.get("batch_round_count"))
                        single_round_count = safe_nonnegative_int(live_payload.get("single_round_count"))
                        fetch_text = (
                            f"｜官網即時資料：{returned_count}／{requested_count} 站"
                            if requested_count else f"｜官網即時資料：{returned_count} 站"
                        )
                        elapsed_seconds = safe_nonnegative_int(live_payload.get("elapsed_ms")) / 1000
                        request_text = (
                            f"｜共 {request_count} 次請求（批次 {batch_round_count} 輪、"
                            f"單站補查 {single_round_count} 輪）｜耗時 {elapsed_seconds:.1f} 秒"
                        )
                        st.success(
                            f"✅ 高速同步完成：已寫入 {live_summary['matched_count']}／"
                            f"{live_summary['total_count']} 個 Excel 場站{fetch_text}{request_text}｜系統取得時間："
                            f"{live_payload['fetched_at']}{official_time_text}"
                        )
                        if missing_count:
                            st.warning(
                                f"官網本次仍有 {missing_count} 個站號未回傳即時資料；系統已完成多輪批次與"
                                "單站補查，未取得者會保留原本數字，不會用 0 覆蓋。"
                            )
                        elif failed_request_count:
                            st.info(
                                f"所有場站皆已取得；過程中有 {failed_request_count} 次暫時失敗，"
                                "已由自動重試或後續補查補齊。"
                            )

                    if not live_report_df.empty:
                        problem_df = live_report_df[live_report_df["結果"] != "已寫入"]
                        with st.expander(
                            f"查看配對明細（未寫入 {len(problem_df)} 筆）",
                            expanded=not problem_df.empty,
                        ):
                            report_height = min(520, max(110, 42 + len(live_report_df) * 35))
                            st.dataframe(
                                live_report_df,
                                hide_index=True,
                                use_container_width=True,
                                height=report_height,
                                row_height=35,
                            )
            except YouBikeDataError as exc:
                st.error(f"YouBike 官網同步失敗：{exc}")
            except Exception as exc:
                st.error(f"YouBike 官網同步發生未預期錯誤：{exc}")



# 多張平板翻拍照片辨識：只讀取「單位」後方的在站 2.0／2.0E。
# 照片先累積在 session_state，因此手機即使每次只能加入一張，也能分批加入後一次辨識。
photo_context_id = (
    f"{active_base['token']}::{selected_sheet}::{selected_route}::{selected_shift}"
)
photo_pool_key = f"station_photo_pool::{photo_context_id}"
photo_uploader_version_key = f"station_photo_uploader_version::{photo_context_id}"

if photo_pool_key not in st.session_state:
    st.session_state[photo_pool_key] = []
if photo_uploader_version_key not in st.session_state:
    st.session_state[photo_uploader_version_key] = 0

with st.expander("照片辨識", expanded=False):
    newly_uploaded_photos = st.file_uploader(
        "上傳照片",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key=(
            f"station_photo_uploader::{photo_context_id}::"
            f"{st.session_state[photo_uploader_version_key]}"
        ),
    )

    # 將本次選取加入照片池；以內容雜湊去重，避免 Streamlit 每次重跑重複累加。
    if newly_uploaded_photos:
        photo_pool = list(st.session_state.get(photo_pool_key, []))
        known_digests = {str(record.get("sha256", "")) for record in photo_pool}

        for uploaded_photo in newly_uploaded_photos:
            photo_bytes = uploaded_photo.getvalue()
            photo_digest = hashlib.sha256(photo_bytes).hexdigest()
            if photo_digest in known_digests:
                continue
            photo_pool.append(
                {
                    "name": str(uploaded_photo.name),
                    "bytes": photo_bytes,
                    "sha256": photo_digest,
                }
            )
            known_digests.add(photo_digest)

        st.session_state[photo_pool_key] = photo_pool

    queued_station_photos = list(st.session_state.get(photo_pool_key, []))

    if queued_station_photos:
        st.caption(f"已加入 {len(queued_station_photos)} 張")

    photo_action_col_1, photo_action_col_2 = st.columns(2)
    with photo_action_col_1:
        run_photo_ocr = st.button(
            "辨識並寫入",
            type="primary",
            use_container_width=True,
            disabled=not queued_station_photos,
            key=f"run_ocr::{photo_context_id}",
        )

    with photo_action_col_2:
        clear_all_photos = st.button(
            "清除照片",
            use_container_width=True,
            disabled=not queued_station_photos,
            key=f"clear_photos::{photo_context_id}",
        )

    if clear_all_photos:
        st.session_state[photo_pool_key] = []
        st.session_state[photo_uploader_version_key] += 1
        rerun_app()

    if run_photo_ocr:
        try:
            with st.spinner(f"正在辨識 {len(queued_station_photos)} 張照片並更新現況……"):
                analysis_result = analyze_station_photos(queued_station_photos, base_df)
                recognized_updates = analysis_result.get("updates", {})

                error_df = analysis_result.get("error_df", pd.DataFrame())

                if recognized_updates:
                    updated_full_df = apply_ocr_updates_to_dataframe(base_df, recognized_updates)
                    status_cache["contexts"][current_context_key] = dataframe_to_status_records(updated_full_df)
                    save_cached_status(active_base["token"], active_base["expires_at"], status_cache)
                    clear_editor_session_state(active_base["token"])
                    base_df = updated_full_df

                    st.success(f"✅ 已直接更新 {len(recognized_updates)} 個場站。")
                    recognized_df = analysis_result.get("recognized_df", pd.DataFrame())
                    if not recognized_df.empty:
                        recognized_table_height = max(105, 40 + (len(recognized_df) * 35) + 4)
                        st.dataframe(
                            recognized_df[["場站名稱", "單位", "2.0 現況", "2.0E 現況", "來源"]],
                            hide_index=True,
                            use_container_width=True,
                            height=recognized_table_height,
                            row_height=35,
                        )
                else:
                    st.error("識別錯誤：這批照片沒有讀到可安全寫入的場站資料。")

                if not error_df.empty:
                    st.error(f"⚠️ 共有 {len(error_df)} 筆識別錯誤，未寫入現況。")
                    error_table_height = max(105, 40 + (len(error_df) * 35) + 4)
                    st.dataframe(
                        error_df[["場站名稱", "2.0 現況", "2.0E 現況", "來源"]],
                        hide_index=True,
                        use_container_width=True,
                        height=error_table_height,
                        row_height=35,
                    )
        except ImportError:
            st.error(
                "尚未安裝免費 OCR 套件。請確認 requirements.txt 內含 rapidocr 與 onnxruntime。"
            )
        except Exception as exc:
            st.error(f"照片辨識失敗：{exc}")

regions = ["全部"] + list(dict.fromkeys(base_df["行政區"].tolist()))
with selected_region_placeholder.container():
    selected_region = st.selectbox("行政區", regions)
working_df = base_df if selected_region == "全部" else base_df[base_df["行政區"] == selected_region]
working_df = working_df.reset_index(drop=True)
st.markdown('<div id="current-status-input-anchor"></div>', unsafe_allow_html=True)

editor_key = (
    f"editor::{active_base['token']}::{selected_sheet}::{selected_route}::"
    f"{selected_shift}::{selected_region}"
)

with st.expander("手動輸入現況", expanded=False):
    with st.form(key=f"status_form::{editor_key}", clear_on_submit=False):
        working_df = coerce_nullable_current_status(working_df)
    
        if mobile_input_mode:
            # value=None 會保留空白；手機點選後仍會叫出數字鍵盤。
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
                        value=normalize_current_status(row["2.0 現況"]),
                        step=1,
                        format="%d",
                        placeholder="空白",
                        key=f"{editor_key}::mobile::2.0::{row_index}",
                        label_visibility="collapsed",
                    )
    
                with ebike_col:
                    ebike_now = st.number_input(
                        f"{row['場站名稱']} 2.0E 現況",
                        min_value=0,
                        value=normalize_current_status(row["2.0E 現況"]),
                        step=1,
                        format="%d",
                        placeholder="空白",
                        key=f"{editor_key}::mobile::2.0E::{row_index}",
                        label_visibility="collapsed",
                    )
    
                edited_df.at[row_index, "2.0 現況"] = (
                    pd.NA if bike_now is None else int(bike_now)
                )
                edited_df.at[row_index, "2.0E 現況"] = (
                    pd.NA if ebike_now is None else int(ebike_now)
                )
    
            edited_df = coerce_nullable_current_status(edited_df)
        else:
            # 電腦版維持可快速批次編輯的資料表，空白儲存格可直接輸入。
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
                    "行政區": None,
                    "場站名稱": st.column_config.TextColumn("場站", width=120),
                    "2.0 現況": st.column_config.NumberColumn(
                        "2.0現", width=52, min_value=0, step=1, format="%d", required=False
                    ),
                    "2.0E 現況": st.column_config.NumberColumn(
                        "2.0E現", width=58, min_value=0, step=1, format="%d", required=False
                    ),
                    "2.0 標準": st.column_config.NumberColumn("2.0標", width=52, format="%d"),
                    "2.0E 標準": st.column_config.NumberColumn("2.0E標", width=58, format="%d"),
                },
            )
            edited_df = coerce_nullable_current_status(edited_df)

        status_form_submitted = st.form_submit_button(
            "套用並儲存",
            type="primary",
            use_container_width=True,
        )

# 表單送出時才會重跑；只有資料真的改變才序列化並寫入磁碟。
full_status_df = merge_current_status(base_df, edited_df)
current_records = dataframe_to_status_records(full_status_df)
previous_records = status_cache["contexts"].get(current_context_key)
if current_records != previous_records:
    status_cache["contexts"][current_context_key] = current_records
    save_cached_status(active_base["token"], active_base["expires_at"], status_cache)

if status_form_submitted:
    st.success("✅ 現況已套用並儲存，分析結果已更新。")

render_smart_dispatch(
    full_status_df=full_status_df,
    selected_region=selected_region,
    status_cache=status_cache,
    current_context_key=current_context_key,
    active_base=active_base,
)

st.markdown('<div id="analysis-results-anchor"></div>', unsafe_allow_html=True)
st.markdown("---")
st.subheader("調度結果")

result_df = build_result_with_recognition_errors(edited_df)

# 分析結果固定排除「2.0、2.0E 都符合」的場站。
# 若只有其中一種車型符合，該場站仍保留，但「符合」不參與排序。
result_df = result_df[
    (result_df["2.0 缺／多幾台"] != "符合")
    | (result_df["2.0E 缺／多幾台"] != "符合")
].reset_index(drop=True)

sort_control_1, sort_control_2 = st.columns([2, 1])
with sort_control_1:
    selected_sort_field = st.selectbox(
        "排序",
        list(SORT_FIELD_OPTIONS.keys()),
        index=0,
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

bike_summary = calculate_inventory_summary(
    edited_df, "2.0 現況", "2.0 標準"
)
ebike_summary = calculate_inventory_summary(
    edited_df, "2.0E 現況", "2.0E 標準"
)

render_inventory_summary_card("2.0 總覽", "bike", bike_summary)
render_inventory_summary_card("2.0E 總覽", "ebike", ebike_summary)

missing_bike_count = safe_nonnegative_int(bike_summary["missing_count"])
missing_ebike_count = safe_nonnegative_int(ebike_summary["missing_count"])
if missing_bike_count or missing_ebike_count:
    render_missing_data_notice(missing_bike_count, missing_ebike_count)

if result_df.empty:
    st.success("✨ 所有場站皆符合配置，目前不需要調度。")
else:
    render_dispatch_legend()

    # 使用固定寬度的響應式表格：完整向下展開，不在表格內產生左右或上下捲動。
    for region, region_df in result_df.groupby("行政區", sort=False):
        st.markdown(f"#### {selected_route}｜{region}")
        render_analysis_result_table(region_df)

    export_df = make_colored_export_df(result_df)
    csv_data = export_df.to_csv(index=False).encode("utf-8-sig")
    excel_data = build_colored_excel(export_df)

    download_col_1, download_col_2 = st.columns(2)
    with download_col_1:
        st.download_button(
            "下載 CSV",
            data=csv_data,
            file_name=f"{selected_route}_{selected_shift}_調度分析_彩色標記.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with download_col_2:
        st.download_button(
            "下載 Excel",
            data=excel_data,
            file_name=f"{selected_route}_{selected_shift}_調度分析_彩色.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

st.caption("即時車數來源：YouBike 官網；以現場狀況為準。")

# 懸浮搜尋只讀取目前排序完成、實際顯示的分析結果。
render_floating_station_search(result_df, mobile_input_mode)
