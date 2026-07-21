--- /mnt/data/app_v24_1_road_optimized.py	2026-07-21 16:17:34.205971262 +0000
+++ /mnt/data/app_v25_all_regions_smart_dispatch.py	2026-07-21 16:25:58.560860570 +0000
@@ -1,7 +1,8 @@
 from __future__ import annotations
 
-# 版本：v24.1｜道路路網智慧調度優化版
+# 版本：v25.0｜三區整合智慧調度版
 
+import base64
 import hashlib
 import html
 import json
@@ -38,9 +39,9 @@
 )
 
 
-APP_VERSION = "v24.1"
-APP_VERSION_NAME = "道路路網智慧調度優化版"
-APP_BUILD_DATE = "2026-07-21"
+APP_VERSION = "v25.0"
+APP_VERSION_NAME = "三區整合智慧調度版"
+APP_BUILD_DATE = "2026-07-22"
 
 
 def is_mobile_browser() -> bool:
@@ -319,6 +320,156 @@
     )
 
 
+
+def render_region_inventory_overview(region_name: str, status_df: pd.DataFrame) -> None:
+    """在每個行政區標題下方顯示 2.0／2.0E 的區域車輛總覽。"""
+    bike_summary = calculate_inventory_summary(status_df, "2.0 現況", "2.0 標準")
+    ebike_summary = calculate_inventory_summary(status_df, "2.0E 現況", "2.0E 標準")
+
+    def metric_html(label: str, summary: dict[str, int | str | None]) -> str:
+        state = html.escape(str(summary.get("state") or "pending"))
+        state_label = html.escape(str(summary.get("state_label") or "資料未完整"))
+        configured_total = safe_nonnegative_int(summary.get("configured_total"))
+        current_total = safe_nonnegative_int(summary.get("current_total"))
+        difference_text = html.escape(str(summary.get("difference_text") or "—"))
+        return f"""
+          <div class="region-fleet-metric region-fleet-{state}">
+            <div class="region-fleet-metric-head">
+              <strong>{html.escape(label)}</strong>
+              <span>{state_label}</span>
+            </div>
+            <div class="region-fleet-numbers">
+              <div><small>配置</small><b>{configured_total}<em>台</em></b></div>
+              <div><small>目前</small><b>{current_total}<em>台</em></b></div>
+              <div><small>差額</small><b>{difference_text}</b></div>
+            </div>
+          </div>
+        """
+
+    st.markdown(
+        f"""
+        <style>
+        .region-fleet-overview{{margin:.35rem 0 .72rem;padding:.72rem;border:1px solid rgba(148,163,184,.24);border-radius:16px;background:rgba(248,250,252,.74)}}
+        .region-fleet-title{{font-size:.78rem;font-weight:850;opacity:.68;margin:0 0 .48rem .08rem}}
+        .region-fleet-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:.48rem}}
+        .region-fleet-metric{{padding:.62rem .66rem;border-radius:13px;background:rgba(255,255,255,.82);border:1px solid rgba(148,163,184,.18)}}
+        .region-fleet-metric-head{{display:flex;justify-content:space-between;align-items:center;gap:.4rem}}
+        .region-fleet-metric-head strong{{font-size:.93rem}}
+        .region-fleet-metric-head span{{font-size:.68rem;font-weight:800;padding:.2rem .4rem;border-radius:999px;background:rgba(148,163,184,.12)}}
+        .region-fleet-extra .region-fleet-metric-head span{{color:#d9363e;background:rgba(244,63,94,.12)}}
+        .region-fleet-short .region-fleet-metric-head span{{color:#b96b00;background:rgba(245,158,11,.15)}}
+        .region-fleet-balanced .region-fleet-metric-head span{{color:#087f5b;background:rgba(16,185,129,.13)}}
+        .region-fleet-numbers{{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.3rem;margin-top:.48rem}}
+        .region-fleet-numbers div{{text-align:center;min-width:0}}
+        .region-fleet-numbers small{{display:block;font-size:.62rem;opacity:.6}}
+        .region-fleet-numbers b{{display:block;font-size:.9rem;margin-top:.12rem;white-space:nowrap}}
+        .region-fleet-numbers em{{font-size:.62rem;font-style:normal;margin-left:.08rem;opacity:.65}}
+        @media(max-width:700px){{.region-fleet-grid{{grid-template-columns:1fr}}}}
+        </style>
+        <section class="region-fleet-overview">
+          <div class="region-fleet-title">{html.escape(region_name)}｜行政區車輛總覽</div>
+          <div class="region-fleet-grid">
+            {metric_html("2.0", bike_summary)}
+            {metric_html("2.0E", ebike_summary)}
+          </div>
+        </section>
+        """,
+        unsafe_allow_html=True,
+    )
+
+
+def render_new_window_download_panel(
+    *,
+    csv_data: bytes,
+    csv_filename: str,
+    excel_data: bytes,
+    excel_filename: str,
+) -> None:
+    """以新視窗／新分頁開啟下載，避免 iOS App 被檔案預覽頁取代後無法返回。"""
+    file_payload = {
+        "csv": {
+            "name": csv_filename,
+            "mime": "text/csv;charset=utf-8",
+            "data": base64.b64encode(csv_data).decode("ascii"),
+        },
+        "excel": {
+            "name": excel_filename,
+            "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
+            "data": base64.b64encode(excel_data).decode("ascii"),
+        },
+    }
+    payload_json = json.dumps(file_payload, ensure_ascii=False).replace("</", "<\\/")
+    components.html(
+        f"""
+        <!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">
+        <meta name="viewport" content="width=device-width,initial-scale=1">
+        <style>
+          *{{box-sizing:border-box}} body{{margin:0;background:transparent;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}}
+          .download-note{{font-size:12px;line-height:1.45;color:#64748b;margin:0 0 8px}}
+          .download-grid{{display:grid;grid-template-columns:1fr 1fr;gap:10px}}
+          button{{width:100%;border:0;border-radius:12px;padding:12px 10px;font-size:15px;font-weight:800;cursor:pointer}}
+          .csv{{background:#e8f3ff;color:#075fb8}} .excel{{background:#e7f8ef;color:#087f5b}}
+        </style></head><body>
+          <p class="download-note">下載會另開新頁或新分頁；看完檔案後關閉下載頁，即可直接回到原本分析畫面。</p>
+          <div class="download-grid">
+            <button class="csv" type="button" onclick="openDownload('csv')">下載 CSV</button>
+            <button class="excel" type="button" onclick="openDownload('excel')">下載 Excel</button>
+          </div>
+        <script>
+          const files = {payload_json};
+          function openDownload(key) {{
+            const file = files[key];
+            const binary = atob(file.data);
+            const bytes = new Uint8Array(binary.length);
+            for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
+            const blob = new Blob([bytes], {{type:file.mime}});
+            const url = URL.createObjectURL(blob);
+            const popup = window.open('', '_blank');
+            if (popup) {{
+              popup.document.title = `下載 ${{file.name}}`;
+              popup.document.body.style.margin = '0';
+              popup.document.body.style.padding = '24px';
+              popup.document.body.style.fontFamily = '-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif';
+              const title = popup.document.createElement('h2');
+              title.textContent = '報表下載';
+              const note = popup.document.createElement('p');
+              note.textContent = '檔案已開始下載。完成後請按下方按鈕關閉此頁，返回原本分析畫面。';
+              const downloadLink = popup.document.createElement('a');
+              downloadLink.href = url;
+              downloadLink.download = file.name;
+              downloadLink.textContent = `再次下載：${{file.name}}`;
+              downloadLink.style.display = 'block';
+              downloadLink.style.margin = '20px 0';
+              const closeButton = popup.document.createElement('button');
+              closeButton.type = 'button';
+              closeButton.textContent = '關閉下載頁，返回系統';
+              closeButton.style.padding = '12px 18px';
+              closeButton.style.border = '0';
+              closeButton.style.borderRadius = '12px';
+              closeButton.style.fontSize = '16px';
+              closeButton.style.fontWeight = '800';
+              closeButton.onclick = () => popup.close();
+              popup.document.body.append(title, note, downloadLink, closeButton);
+              downloadLink.click();
+            }} else {{
+              const anchor = document.createElement('a');
+              anchor.href = url;
+              anchor.download = file.name;
+              anchor.target = '_blank';
+              anchor.rel = 'noopener';
+              document.body.appendChild(anchor);
+              anchor.click();
+              anchor.remove();
+            }}
+            window.setTimeout(() => URL.revokeObjectURL(url), 120000);
+          }}
+        </script></body></html>
+        """,
+        height=106,
+        scrolling=False,
+    )
+
+
 def render_missing_data_notice(missing_bike_count: int, missing_ebike_count: int) -> None:
     """渲染資料不完整提示；有空白時暫停顯示整體缺／多差額。"""
     st.markdown(
@@ -2307,11 +2458,11 @@
     """顯示目前工作情境，讓使用者不用回頭確認選項。"""
     live_meta = live_meta if isinstance(live_meta, dict) else {}
     fetched_at = html.escape(str(live_meta.get("fetched_at") or "尚未同步"))
-    mode_label = "AI 路線" if "AI" in page_mode else "一般分析"
+    mode_label = "智慧調度" if page_mode == "智慧調度" else "一般分析"
     st.markdown(
         f"""
         <div class="dispatch-context-strip">
-          <span><b>配置</b>{html.escape(route)}</span>
+          <span><b>範圍</b>{html.escape(route)}</span>
           <span><b>班別</b>{html.escape(shift)}</span>
           <span><b>場站</b>{safe_nonnegative_int(station_count)} 站</span>
           <span><b>模式</b>{html.escape(mode_label)}</span>
@@ -3646,21 +3797,44 @@
         "longitude": 121.219459,
     },
 }
+ALL_DISPATCH_ZONES = ("D1", "D2", "D3")
 LONG_DISTANCE_ROUTE_ZONES = ("D2", "D3")
 LONG_DISTANCE_LOOP_DIRECTION_OPTIONS = ("AI 自動選擇", "D2 先行", "D3 先行")
 LONG_DISTANCE_TRANSFER_LABEL = "玉長公路"
 SHARED_GEOLOCATION_REFRESH_SECONDS = 30
 
 
-def normalize_long_distance_zone(value) -> str | None:
-    """把配置中的路線名稱辨識為 D2／D3。"""
+def normalize_dispatch_zone(value) -> str | None:
+    """把配置中的路線名稱辨識為 D1／D2／D3。"""
     normalized = re.sub(r"\s+", "", str(value or "").upper())
-    for zone in LONG_DISTANCE_ROUTE_ZONES:
+    for zone in ALL_DISPATCH_ZONES:
         if zone in normalized:
             return zone
     return None
 
 
+def normalize_long_distance_zone(value) -> str | None:
+    """長途環狀邏輯只接受 D2／D3。"""
+    zone = normalize_dispatch_zone(value)
+    return zone if zone in LONG_DISTANCE_ROUTE_ZONES else None
+
+
+def preferred_configuration_sheet(options: list[tuple[str, str]]) -> str:
+    """自動選出同時涵蓋 D1／D2／D3 最完整的工作表，不再要求使用者選配置版本。"""
+    sheet_order: list[str] = []
+    coverage: dict[str, set[str]] = {}
+    for sheet_name, route in options:
+        if sheet_name not in coverage:
+            coverage[sheet_name] = set()
+            sheet_order.append(sheet_name)
+        zone = normalize_dispatch_zone(route)
+        if zone:
+            coverage[sheet_name].add(zone)
+    if not sheet_order:
+        return ""
+    return max(sheet_order, key=lambda sheet: (len(coverage.get(sheet, set())), -sheet_order.index(sheet)))
+
+
 def location_payload_is_valid(location) -> bool:
     if not isinstance(location, dict):
         return False
@@ -3765,22 +3939,22 @@
     selected_zones: list[str],
     status_cache: dict,
 ) -> tuple[pd.DataFrame, dict[str, dict]]:
-    """整合 D2／D3 配置、既有現況、最新即時車數與官方座標。"""
-    selected_zone_set = {zone for zone in selected_zones if zone in LONG_DISTANCE_ROUTE_ZONES}
+    """整合指定 D1／D2／D3 配置、既有現況、最新即時車數與官方座標。"""
+    selected_zone_set = {zone for zone in selected_zones if zone in ALL_DISPATCH_ZONES}
     if not selected_zone_set:
         return pd.DataFrame(), {}
 
-    # D2、D3 可能分別放在不同工作表。每一區都先找目前所選工作表，
-    # 找不到時再獨立到其他可用配置中尋找，避免單選 D2 時把 D3 一起漏掉。
+    # D1、D2、D3 可能分別放在不同工作表。每一區先找自動選定的主工作表，
+    # 找不到時再到其他可用資料中補齊，確保三區都會被讀取。
     chosen_by_zone: dict[str, tuple[str, str]] = {}
-    for zone in LONG_DISTANCE_ROUTE_ZONES:
+    for zone in ALL_DISPATCH_ZONES:
         if zone not in selected_zone_set:
             continue
         same_sheet_match = next(
             (
                 (sheet_name, route)
                 for sheet_name, route in options
-                if sheet_name == selected_sheet and normalize_long_distance_zone(route) == zone
+                if sheet_name == selected_sheet and normalize_dispatch_zone(route) == zone
             ),
             None,
         )
@@ -3788,7 +3962,7 @@
             (
                 (sheet_name, route)
                 for sheet_name, route in options
-                if normalize_long_distance_zone(route) == zone
+                if normalize_dispatch_zone(route) == zone
             ),
             None,
         )
@@ -3810,7 +3984,7 @@
     combined_locations: dict[str, dict] = {}
     cache_changed = False
 
-    for zone in LONG_DISTANCE_ROUTE_ZONES:
+    for zone in ALL_DISPATCH_ZONES:
         if zone not in selected_zone_set or zone not in chosen_by_zone:
             continue
         sheet_name, route = chosen_by_zone[zone]
@@ -3921,8 +4095,10 @@
             endpoint_distance_km = float(endpoint_metric.get("road_distance_km") or 0.0)
             endpoint_drive_minutes = float(endpoint_metric.get("drive_minutes") or 0.0)
 
-        # 單趟終點採 35% 方向性；來回完整計入返程；環狀僅保留少量回維調成本。
-        if trip_mode == "來回":
+        # 一般模式只評估「目前位置 → 單一場站」的即時效益，不安排後續路線。
+        if trip_mode == "一般模式":
+            endpoint_weight = 0.0
+        elif trip_mode == "來回":
             endpoint_weight = 1.0
         elif trip_mode == "環狀一圈":
             endpoint_weight = 0.20 if endpoint_valid else 0.0
@@ -5120,8 +5296,9 @@
     station_locations_override: dict[str, dict] | None = None,
     allow_manual_station_choice: bool = False,
     show_route_preview: bool = False,
+    require_external_location: bool = False,
 ) -> None:
-    """逐站詢問、載量限制、動態重算；長途頁支援單趟、來回及 D2／D3 環狀分階段路線。"""
+    """逐站詢問、載量限制、動態重算；一般模式可強制只採用背景 GPS。"""
     st.markdown('<div id="smart-dispatch-anchor"></div>', unsafe_allow_html=True)
     st.subheader(page_title)
     st.caption(
@@ -5292,7 +5469,7 @@
         )
         return
 
-    shared_location_mode = fallback_location is not None or external_location is not None
+    shared_location_mode = require_external_location or fallback_location is not None or external_location is not None
     location_payload = None
     if not shared_location_mode:
         try:
@@ -5334,7 +5511,10 @@
                     st.warning(f"目前位置尚未更新：{location_payload.get('error') or '定位失敗'}")
 
     stored_location = st.session_state.get(location_state_key)
-    current_location = newest_valid_location(stored_location, external_location, fallback_location)
+    if require_external_location:
+        current_location = dict(external_location) if location_payload_is_valid(external_location) else None
+    else:
+        current_location = newest_valid_location(stored_location, external_location, fallback_location)
     location_request_pending = bool(st.session_state.get(location_request_pending_key, False))
     if location_request_pending and not shared_location_mode:
         st.info("正在讀取取消配置後的目前位置；定位完成後會自動重新安排下一個場站。")
@@ -5677,10 +5857,16 @@
         else:
             road_status = st.session_state.get(ROAD_ROUTER_STATUS_STATE_KEY, {})
             if isinstance(road_status, dict) and road_status.get("ok") is False:
-                st.error(
-                    "道路路網服務暫時無法使用。為避免用直線誤判、橫切山脈，本次不產生 AI 路線；"
-                    "請稍後按更新重新計算。"
-                )
+                if trip_mode == "一般模式":
+                    st.error(
+                        "道路路網服務暫時無法使用。為避免用直線距離誤判效率，本次不產生場站推薦；"
+                        "請稍後按更新重新計算。"
+                    )
+                else:
+                    st.error(
+                        "道路路網服務暫時無法使用。為避免用直線誤判、橫切山脈，本次不產生 AI 路線；"
+                        "請稍後按更新重新計算。"
+                    )
             else:
                 st.warning(
                     "目前找不到可執行的下一站。可能原因：全部符合配置、車上車種不足、貨車已滿、"
@@ -5736,7 +5922,7 @@
             manual_col_1, manual_col_2 = st.columns(2)
             with manual_col_1:
                 manual_confirmed = st.button(
-                    "指定此站並重排後續路線",
+                    "指定此站並重新計算" if not show_route_preview else "指定此站並重排後續路線",
                     type="primary",
                     use_container_width=True,
                     key=f"{dispatch_prefix}::confirm_manual_station",
@@ -5887,14 +6073,15 @@
     status_cache: dict,
     shared_location: dict | None,
 ) -> None:
-    """獨立的 D2／D3 長距離 AI 路線頁。"""
+    """整合一般即時推薦與 D2／D3 長途路線的智慧調度頁。"""
     st.markdown(
         """
         <section style="padding:1rem 1.05rem;border:1px solid rgba(22,119,255,.22);border-radius:20px;
         background:linear-gradient(135deg,rgba(22,119,255,.08),rgba(16,185,129,.05));margin:.35rem 0 .9rem;">
-          <div style="font-size:1.45rem;font-weight:900;">🚚 D2／D3 AI 長途路線</div>
+          <div style="font-size:1.45rem;font-weight:900;">🚚 智慧調度</div>
           <div style="margin-top:.3rem;font-size:.86rem;opacity:.72;line-height:1.55;">
-            此頁只處理 D2、D3 長距離調度；AI 依實際可行駛道路與後續三站效益安排，不再用南北緯度判定折返。
+            一般模式只依最新定位、即時車數、道路時間與貨車載量推薦最高效率單站；
+            單趟、來回與環狀模式才會建立後續路線。
           </div>
         </section>
         """,
@@ -5902,7 +6089,7 @@
     )
 
     settings_prefix = f"long_distance_settings::{active_base['token']}::{selected_shift}"
-    long_context_key = f"D2D3長途｜{selected_shift}"
+    long_context_key = f"全區智慧調度｜{selected_shift}"
     dispatch_prefix = f"smart_dispatch::long_distance::{active_base['token']}::{long_context_key}"
     active_trip_key = f"{dispatch_prefix}::active_trip"
     loop_order_key = f"{dispatch_prefix}::loop_zone_order"
@@ -5936,78 +6123,84 @@
                 rerun_app()
 
     with st.expander(
-        "本次長途任務設定" + ("（路線執行中已鎖定）" if settings_locked else ""),
+        "本次智慧調度設定" + ("（任務執行中已鎖定）" if settings_locked else ""),
         expanded=not settings_locked,
     ):
-        setting_col_1, setting_col_2 = st.columns(2)
-        with setting_col_1:
+        trip_mode = st.radio(
+            "路線模式",
+            ["一般模式", "單趟", "來回", "環狀一圈"],
+            horizontal=True,
+            key=f"{settings_prefix}::trip_mode",
+            disabled=settings_locked,
+        )
+
+        loop_direction_preference = ""
+        start_name = ""
+        endpoint_name = ""
+
+        if trip_mode == "一般模式":
+            selected_zones = list(ALL_DISPATCH_ZONES)
+            st.caption(
+                "一般模式固定讀取 D1、D2、D3；只使用每 30 秒更新的即時定位安排最高效率場站，"
+                "不建立完整路線或後續站序。"
+            )
+        else:
             start_name = st.selectbox(
                 "出發維調",
                 list(LONG_DISTANCE_START_POINTS.keys()),
                 key=f"{settings_prefix}::start_name",
                 disabled=settings_locked,
             )
-        with setting_col_2:
-            trip_mode = st.radio(
-                "路線模式",
-                ["單趟", "來回", "環狀一圈"],
-                horizontal=True,
-                key=f"{settings_prefix}::trip_mode",
-                disabled=settings_locked,
-            )
 
-        loop_direction_preference = ""
-        if trip_mode == "環狀一圈":
-            selected_zones = list(LONG_DISTANCE_ROUTE_ZONES)
-            st.multiselect(
-                "執行範圍",
-                list(LONG_DISTANCE_ROUTE_ZONES),
-                default=list(LONG_DISTANCE_ROUTE_ZONES),
-                key=f"{settings_prefix}::loop_zones",
-                disabled=True,
-                help="環狀一圈必須同時載入 D2 與 D3，並各自從對應配置版本讀取。",
-            )
-            loop_direction_preference = st.radio(
-                "環狀方向",
-                list(LONG_DISTANCE_LOOP_DIRECTION_OPTIONS),
-                horizontal=True,
-                key=f"{settings_prefix}::loop_direction",
-                disabled=settings_locked,
-            )
-            endpoint_name = start_name
-            st.caption(
-                f"環狀規則：從 {start_name} 出發 → 先完成第一區 → 經 {LONG_DISTANCE_TRANSFER_LABEL} "
-                "跨越海岸山脈 → 完成第二區 → 返回出發維調。開始後不會在 D2、D3 之間反覆折返。"
-            )
-        else:
-            selected_zones = st.multiselect(
-                "執行範圍",
-                list(LONG_DISTANCE_ROUTE_ZONES),
-                default=list(LONG_DISTANCE_ROUTE_ZONES),
-                key=f"{settings_prefix}::zones",
-                disabled=trip_locked,
-            )
-            endpoint_name = ""
-            if trip_mode == "單趟":
-                endpoint_choice = st.selectbox(
-                    "單趟結束方向",
-                    ["最後一站結束", "台東維調", "池上維調"],
-                    key=f"{settings_prefix}::single_endpoint",
-                    disabled=trip_locked,
+            if trip_mode == "環狀一圈":
+                selected_zones = list(LONG_DISTANCE_ROUTE_ZONES)
+                st.multiselect(
+                    "執行範圍",
+                    list(LONG_DISTANCE_ROUTE_ZONES),
+                    default=list(LONG_DISTANCE_ROUTE_ZONES),
+                    key=f"{settings_prefix}::loop_zones",
+                    disabled=True,
+                    help="環狀一圈固定同時載入 D2 與 D3。",
+                )
+                loop_direction_preference = st.radio(
+                    "環狀方向",
+                    list(LONG_DISTANCE_LOOP_DIRECTION_OPTIONS),
+                    horizontal=True,
+                    key=f"{settings_prefix}::loop_direction",
+                    disabled=settings_locked,
                 )
-                if endpoint_choice != "最後一站結束":
-                    endpoint_name = endpoint_choice
-            else:
                 endpoint_name = start_name
+                st.caption(
+                    f"環狀規則：從 {start_name} 出發 → 先完成第一區 → 經 {LONG_DISTANCE_TRANSFER_LABEL} "
+                    "跨越海岸山脈 → 完成第二區 → 返回出發維調。開始後不會在 D2、D3 之間反覆折返。"
+                )
+            else:
+                selected_zones = st.multiselect(
+                    "執行範圍",
+                    list(LONG_DISTANCE_ROUTE_ZONES),
+                    default=list(LONG_DISTANCE_ROUTE_ZONES),
+                    key=f"{settings_prefix}::zones",
+                    disabled=trip_locked,
+                )
+                if trip_mode == "單趟":
+                    endpoint_choice = st.selectbox(
+                        "單趟結束方向",
+                        ["最後一站結束", "台東維調", "池上維調"],
+                        key=f"{settings_prefix}::single_endpoint",
+                        disabled=trip_locked,
+                    )
+                    if endpoint_choice != "最後一站結束":
+                        endpoint_name = endpoint_choice
+                else:
+                    endpoint_name = start_name
 
-        start_data = LONG_DISTANCE_START_POINTS[start_name]
-        st.caption(f"{start_name}｜{start_data['description']}（固定位置為約略點，實際計算優先採用最新 GPS）")
+            start_data = LONG_DISTANCE_START_POINTS[start_name]
+            st.caption(f"{start_name}｜{start_data['description']}（固定位置為約略點，實際計算優先採用最新 GPS）")
 
     if not selected_zones:
-        st.warning("請至少選擇 D2 或 D3。")
+        st.warning("請至少選擇一個執行區域。")
         return
 
-    # 離開環狀模式時移除舊的環狀鎖定，避免下次一般單趟／來回沿用舊階段。
     if trip_mode != "環狀一圈" and not trip_locked:
         st.session_state.pop(loop_order_key, None)
         st.session_state.pop(loop_phase_key, None)
@@ -6030,11 +6223,12 @@
             if isinstance(meta, dict) and meta.get("fetched_at")
         ]
         if latest_times:
-            st.caption(f"即時車數最近同步：{max(latest_times)}｜資料變動後會重新計算未鎖定路線")
+            suffix = "重新計算未鎖定推薦" if trip_mode == "一般模式" else "重新計算未鎖定路線"
+            st.caption(f"即時車數最近同步：{max(latest_times)}｜資料變動後會{suffix}")
         else:
             st.caption("即時車數尚未完成第一次同步；上方同步元件完成後會自動帶入。")
 
-    long_status_df, station_locations = build_long_distance_status_dataframe(
+    dispatch_status_df, station_locations = build_long_distance_status_dataframe(
         active_base=active_base,
         options=options,
         selected_sheet=selected_sheet,
@@ -6042,28 +6236,20 @@
         selected_zones=selected_zones,
         status_cache=status_cache,
     )
-    if long_status_df.empty:
-        st.warning("目前配置中找不到可用的 D2／D3 場站。")
+    if dispatch_status_df.empty:
+        st.warning("目前配置中找不到可用的場站。")
         return
 
-    if "配置來源" in long_status_df.columns:
-        source_parts = []
-        for zone in LONG_DISTANCE_ROUTE_ZONES:
-            zone_sources = long_status_df.loc[
-                long_status_df["路線區域"].astype(str).eq(zone), "配置來源"
-            ].astype(str).drop_duplicates().tolist()
-            if zone_sources:
-                source_parts.append(f"{zone}：{zone_sources[0]}")
-        if source_parts:
-            st.caption("配置來源｜" + "｜".join(source_parts))
-
-    start_location = {
-        "latitude": float(LONG_DISTANCE_START_POINTS[start_name]["latitude"]),
-        "longitude": float(LONG_DISTANCE_START_POINTS[start_name]["longitude"]),
-        "accuracy": 0.0,
-        "updated_at": 0.0,
-        "source": "dispatch_start",
-    }
+    start_location = None
+    if start_name:
+        start_location = {
+            "latitude": float(LONG_DISTANCE_START_POINTS[start_name]["latitude"]),
+            "longitude": float(LONG_DISTANCE_START_POINTS[start_name]["longitude"]),
+            "accuracy": 0.0,
+            "updated_at": 0.0,
+            "source": "dispatch_start",
+        }
+
     endpoint_location = None
     endpoint_label = ""
     if endpoint_name:
@@ -6076,22 +6262,32 @@
         }
         endpoint_label = f"{endpoint_name}（{endpoint_data['description']}）"
 
-    zone_text = "＋".join(selected_zones)
+    if trip_mode == "一般模式":
+        page_title = "全區最高效益場站"
+        page_caption = (
+            "系統只評估目前 GPS 位置到各場站的可行駛道路時間、可調度台數與貨車剩餘載量，"
+            "每次只推薦一站，不產生完整路線。即時車數或定位更新後，未鎖定推薦會立即重算。"
+        )
+    else:
+        zone_text = "＋".join(selected_zones)
+        page_title = f"{zone_text} AI 動態路線"
+        page_caption = (
+            "AI 會依可行駛道路時間預看後續三站，再決定真正的下一站；每完成一站、即時數據更新或手動改選後，"
+            "都會重排未鎖定路線。市區短折返會自然保留，偏遠場站與繞山成本則會完整計入。"
+        )
+
     render_smart_dispatch(
-        full_status_df=long_status_df,
+        full_status_df=dispatch_status_df,
         selected_region="全部",
         status_cache=status_cache,
         current_context_key=long_context_key,
         active_base=active_base,
-        page_title=f"{zone_text} AI 動態路線",
-        page_caption=(
-            "AI 會依可行駛道路時間預看後續三站，再決定真正的下一站；每完成一站、即時數據更新或手動改選後，"
-            "都會重排未鎖定路線。市區短折返會自然保留，偏遠場站與繞山成本則會完整計入。"
-        ),
+        page_title=page_title,
+        page_caption=page_caption,
         dispatch_scope="long_distance",
         external_location=shared_location,
-        fallback_location=start_location,
-        location_label=start_name,
+        fallback_location=None if trip_mode == "一般模式" else start_location,
+        location_label="即時 GPS" if trip_mode == "一般模式" else start_name,
         trip_mode=trip_mode,
         endpoint_location=endpoint_location,
         endpoint_label=endpoint_label,
@@ -6099,7 +6295,8 @@
         loop_start_name=start_name,
         station_locations_override=station_locations,
         allow_manual_station_choice=True,
-        show_route_preview=True,
+        show_route_preview=trip_mode != "一般模式",
+        require_external_location=trip_mode == "一般模式",
     )
 
 
@@ -6258,54 +6455,58 @@
             st.session_state["full_reset_notice"] = True
             rerun_app()
 
-option_labels = [f"{route}｜{sheet_name}" for sheet_name, route in options]
-with st.container(border=True):
-    config_col, shift_col = st.columns([1.55, 0.75])
-    with config_col:
-        selected_label = st.selectbox(
-            "配置版本",
-            option_labels,
-            index=None,
-            placeholder="選擇配置版本",
-            key=f"config_version::{active_base['token']}",
-        )
-    with shift_col:
-        selected_shift = st.selectbox(
-            "班別",
-            list(SHIFT_COLUMNS.keys()),
-            key=f"shift::{active_base['token']}",
-        )
+selected_sheet = preferred_configuration_sheet(options)
+available_zone_set = {
+    zone
+    for _sheet_name, route in options
+    if (zone := normalize_dispatch_zone(route)) is not None
+}
+missing_zones = [zone for zone in ALL_DISPATCH_ZONES if zone not in available_zone_set]
+if missing_zones:
+    st.error(f"配置表缺少以下區域：{'、'.join(missing_zones)}。系統必須同時讀取 D1、D2、D3。")
+    st.stop()
 
+with st.container(border=True):
+    selected_shift = st.selectbox(
+        "班別",
+        list(SHIFT_COLUMNS.keys()),
+        key=f"shift::{active_base['token']}",
+    )
     page_mode = st.radio(
         "工作模式",
-        ["一般分析", "D2／D3 AI 路線"],
+        ["一般分析", "智慧調度"],
         horizontal=True,
         key=f"page_mode::{active_base['token']}::{selected_shift}",
     )
 
-# 未主動選擇配置前，不解析、不恢復現況，也不執行任何同步或辨識。
-if selected_label is None:
-    st.info("選擇配置版本後，系統會立即載入場站並開始即時同步。")
-    st.stop()
-
-selected_index = option_labels.index(selected_label)
-selected_sheet, selected_route = options[selected_index]
-base_df = cached_parse_route(active_base["bytes"], selected_sheet, selected_route, selected_shift)
-
+selected_route = "D1＋D2＋D3"
+current_context_key = f"D1D2D3整合｜{selected_shift}"
 status_cache = load_cached_status(active_base["token"], active_base["expires_at"])
-current_context_key = status_context_key(selected_sheet, selected_route, selected_shift)
-saved_current_status = status_cache["contexts"].get(current_context_key)
-base_df = blank_current_status(base_df)
-if saved_current_status is not None:
-    base_df = restore_current_status(base_df, saved_current_status)
+base_df, combined_station_locations = build_long_distance_status_dataframe(
+    active_base=active_base,
+    options=options,
+    selected_sheet=selected_sheet,
+    selected_shift=selected_shift,
+    selected_zones=list(ALL_DISPATCH_ZONES),
+    status_cache=status_cache,
+)
 
 if base_df.empty:
-    st.warning("此配置沒有可用場站。")
+    st.warning("D1、D2、D3 沒有可用場站。")
     st.stop()
 
+aggregate_meta = status_cache.setdefault("metadata", {}).setdefault(current_context_key, {})
+if combined_station_locations:
+    previous_locations = aggregate_meta.get("station_locations", {})
+    if not isinstance(previous_locations, dict):
+        previous_locations = {}
+    merged_locations = dict(previous_locations)
+    merged_locations.update(combined_station_locations)
+    aggregate_meta["station_locations"] = merged_locations
+
 previous_live_meta = status_cache.get("metadata", {}).get(current_context_key, {})
 render_context_strip(
-    route=selected_route,
+    route="D1／D2／D3",
     shift=selected_shift,
     station_count=len(base_df),
     page_mode=page_mode,
@@ -6390,8 +6591,8 @@
                         st.error("沒有任何場站通過安全配對，因此未修改現況資料。")
                     else:
                         base_df = live_updated_df
-                        status_cache["contexts"][current_context_key] = dataframe_to_status_records(base_df)
-                        status_cache.setdefault("metadata", {})[current_context_key] = {
+                        live_event_id = str(live_payload.get("event_id") or browser_event_id or "")
+                        common_live_meta = {
                             "source": live_payload.get("source", "YouBike 官網公開接口（瀏覽器直連，免 TDX）"),
                             "fetched_at": live_payload["fetched_at"],
                             "latest_source_time": live_payload.get("latest_source_time", ""),
@@ -6405,12 +6606,17 @@
                             "batch_round_count": live_payload.get("batch_round_count", 0),
                             "single_round_count": live_payload.get("single_round_count", 0),
                             "station_locations": previous_location_map,
-                            "last_live_event_id": str(live_payload.get("event_id") or browser_event_id or ""),
+                            "last_live_event_id": live_event_id,
                         }
-                        save_cached_status(
-                            active_base["token"],
-                            active_base["expires_at"],
-                            status_cache,
+                        status_cache.setdefault("metadata", {})[current_context_key] = dict(common_live_meta)
+                        if "_狀態內容鍵" in base_df.columns:
+                            for source_context_key in base_df["_狀態內容鍵"].dropna().astype(str).unique():
+                                status_cache.setdefault("metadata", {})[source_context_key] = dict(common_live_meta)
+                        save_dispatch_dataframe_contexts(
+                            base_df,
+                            status_cache=status_cache,
+                            active_base=active_base,
+                            default_context_key=current_context_key,
                         )
                         clear_editor_session_state(active_base["token"])
                         official_time = str(live_payload.get("latest_source_time") or "").strip()
@@ -6467,7 +6673,7 @@
                 st.error(f"YouBike 官網同步發生未預期錯誤：{exc}")
 
 
-if page_mode == "D2／D3 AI 路線":
+if page_mode == "智慧調度":
     render_long_distance_route_page(
         active_base=active_base,
         options=options,
@@ -6476,15 +6682,13 @@
         status_cache=status_cache,
         shared_location=shared_location,
     )
-    st.caption("即時車數來源：YouBike 官網；長途路線與距離為動態估算，仍請以現場與實際路況為準。")
+    st.caption("即時車數來源：YouBike 官網；智慧推薦與道路時間為動態估算，仍請以現場與實際路況為準。")
     st.stop()
 
 
 # 多張平板翻拍照片辨識：只讀取「單位」後方的在站 2.0／2.0E。
 # 照片先累積在 session_state，因此手機即使每次只能加入一張，也能分批加入後一次辨識。
-photo_context_id = (
-    f"{active_base['token']}::{selected_sheet}::{selected_route}::{selected_shift}"
-)
+photo_context_id = f"{active_base['token']}::{current_context_key}"
 photo_pool_key = f"station_photo_pool::{photo_context_id}"
 photo_uploader_version_key = f"station_photo_uploader_version::{photo_context_id}"
 
@@ -6569,8 +6773,12 @@
 
                     if recognized_updates:
                         updated_full_df = apply_ocr_updates_to_dataframe(base_df, recognized_updates)
-                        status_cache["contexts"][current_context_key] = dataframe_to_status_records(updated_full_df)
-                        save_cached_status(active_base["token"], active_base["expires_at"], status_cache)
+                        save_dispatch_dataframe_contexts(
+                            updated_full_df,
+                            status_cache=status_cache,
+                            active_base=active_base,
+                            default_context_key=current_context_key,
+                        )
                         clear_editor_session_state(active_base["token"])
                         base_df = updated_full_df
 
@@ -6612,8 +6820,7 @@
 st.markdown('<div id="current-status-input-anchor"></div>', unsafe_allow_html=True)
 
 editor_key = (
-    f"editor::{active_base['token']}::{selected_sheet}::{selected_route}::"
-    f"{selected_shift}::{selected_region}"
+    f"editor::{active_base['token']}::{current_context_key}::{selected_region}"
 )
 
 edited_df = working_df.copy()
@@ -6726,8 +6933,12 @@
     current_records = dataframe_to_status_records(full_status_df)
     previous_records = status_cache["contexts"].get(current_context_key)
     if current_records != previous_records:
-        status_cache["contexts"][current_context_key] = current_records
-        save_cached_status(active_base["token"], active_base["expires_at"], status_cache)
+        save_dispatch_dataframe_contexts(
+            full_status_df,
+            status_cache=status_cache,
+            active_base=active_base,
+            default_context_key=current_context_key,
+        )
         base_df = full_status_df
     st.success("✅ 現況已套用並儲存，分析結果已更新。")
 
@@ -6792,11 +7003,26 @@
 else:
     render_dispatch_legend()
 
-    # 使用固定寬度的響應式表格：完整向下展開，不在表格內產生左右或上下捲動。
-    for region, region_df in result_df.groupby("行政區", sort=False):
-        st.markdown(f"#### {selected_route}｜{region}")
-        render_analysis_result_table(region_df)
+# 不論該區是否有缺／多車，都保留行政區區塊並顯示區域總覽。
+for region in edited_df["行政區"].astype(str).drop_duplicates():
+    region_source_df = edited_df[edited_df["行政區"].astype(str).eq(region)].copy()
+    region_result_df = result_df[result_df["行政區"].astype(str).eq(region)].copy()
+    zone_values = []
+    if "路線區域" in region_source_df.columns:
+        zone_values = [
+            zone for zone in region_source_df["路線區域"].astype(str).drop_duplicates().tolist()
+            if zone in ALL_DISPATCH_ZONES
+        ]
+    zone_prefix = "／".join(zone_values)
+    heading = f"{zone_prefix}｜{region}" if zone_prefix else region
+    st.markdown(f"#### {heading}")
+    render_region_inventory_overview(region, region_source_df)
+    if region_result_df.empty:
+        st.success("此行政區目前全部符合配置。")
+    else:
+        render_analysis_result_table(region_result_df)
 
+if not result_df.empty:
     export_tools_open = on_demand_toggle(
         "⬇️ 開啟報表下載",
         key=f"export_tools_open::{current_context_key}::{selected_region}",
@@ -6807,24 +7033,12 @@
             export_df = make_colored_export_df(result_df)
             csv_data = export_df.to_csv(index=False).encode("utf-8-sig")
             excel_data = build_colored_excel(export_df)
-
-            download_col_1, download_col_2 = st.columns(2)
-            with download_col_1:
-                st.download_button(
-                    "下載 CSV",
-                    data=csv_data,
-                    file_name=f"{selected_route}_{selected_shift}_調度分析_彩色標記.csv",
-                    mime="text/csv",
-                    use_container_width=True,
-                )
-            with download_col_2:
-                st.download_button(
-                    "下載 Excel",
-                    data=excel_data,
-                    file_name=f"{selected_route}_{selected_shift}_調度分析_彩色.xlsx",
-                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
-                    use_container_width=True,
-                )
+            render_new_window_download_panel(
+                csv_data=csv_data,
+                csv_filename=f"D1_D2_D3_{selected_shift}_調度分析_彩色標記.csv",
+                excel_data=excel_data,
+                excel_filename=f"D1_D2_D3_{selected_shift}_調度分析_彩色.xlsx",
+            )
 
 st.caption("即時車數來源：YouBike 官網；以現場狀況為準。")
 
