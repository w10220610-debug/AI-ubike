from __future__ import annotations

V29_COMPATIBILITY_NOTES = """V29 old-UI compatibility entrypoint.

The original legacy UI is kept byte-for-byte in ``legacy_ui.py``. Before
executing it, this entrypoint applies focused V29 compatibility fixes:

1. battery ranges accept any non-empty Excel zone instead of D1/D2/D3 only;
2. an uploaded workbook never gets merged with the built-in Taitung fallback;
3. Taitung fallback remains available only in the existing no-workbook path;
4. live station sync uses the V29 Python Server service instead of requiring a
   hidden browser Streamlit component on mobile;
5. the floating refresh button requests a fresh server sync by reloading with a
   one-time refresh token when no browser component exists;
6. the floating battery query uses the V29 Fast Client battery engine and a
   mobile-safe one-way HTML UI, avoiding custom-component readiness failures;
7. the V29 battery entry occupies the exact legacy battery-button slot so the
   new engine replaces the old entry instead of appearing as a second control;
8. AI learning guard separates natural demand, confirmed manual intervention
   and suspected intervention before future model training;
9. geolocation uses a visible direct user-triggered control before background
   refresh, improving iPhone/in-app-browser permission reliability.
"""

import time
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from ai_learning_guard import (
    MAX_MANUAL_EVENTS,
    build_manual_intervention_event,
    classify_live_transition,
    trim_learning_records,
)
from battery_icon_data import BATTERY_ICON_DATA_URI
from battery_upgrade import render_floating_server_battery as _render_floating_server_battery


LEGACY_APP = Path(__file__).with_name("legacy_ui.py")
source = LEGACY_APP.read_text(encoding="utf-8")


AI_TIMEZONE = ZoneInfo("Asia/Taipei")
AI_NIGHT_SHIFT_END_HOUR = 7
AI_NIGHT_SHIFT_END_MINUTE = 30
AI_LEARNING_META_KEY = "__ai_learning__"


def resolve_ai_shift_context(shift: str, now: datetime | None = None) -> dict[str, str]:
    """Resolve the main-page shift into the AI learning day context.

    Early/late shifts follow the normal calendar day. The legacy ``夜班配置``
    is the user's 大夜 shift: before/through 07:30 it belongs to the calendar
    day already reached after midnight; after 07:30 it belongs to the next
    operating day. This keeps one 21:30-07:30 shift on one learning date.
    """
    local_now = now.astimezone(AI_TIMEZONE) if now is not None else datetime.now(AI_TIMEZONE)
    raw_shift = str(shift or "").strip()

    if "大夜" in raw_shift or ("夜班" in raw_shift and "晚班" not in raw_shift):
        shift_label = "大夜"
    elif "早班" in raw_shift:
        shift_label = "早班"
    elif "晚班" in raw_shift:
        shift_label = "晚班"
    else:
        shift_label = raw_shift.replace("配置", "") or "未設定"

    operating_date = local_now.date()
    if shift_label == "大夜":
        current_hm = (local_now.hour, local_now.minute)
        night_end_hm = (AI_NIGHT_SHIFT_END_HOUR, AI_NIGHT_SHIFT_END_MINUTE)
        if current_hm > night_end_hm:
            operating_date += timedelta(days=1)

    day_type = "假日" if operating_date.weekday() >= 5 else "平日"
    return {
        "source_shift": raw_shift,
        "shift": shift_label,
        "day_type": day_type,
        "operating_date": operating_date.isoformat(),
        "actual_datetime": local_now.isoformat(),
    }


def _ai_learning_meta(status_cache: dict) -> dict:
    metadata = status_cache.setdefault("metadata", {})
    learning = metadata.setdefault(AI_LEARNING_META_KEY, {})
    if not isinstance(learning, dict):
        learning = {}
        metadata[AI_LEARNING_META_KEY] = learning
    return learning


def render_ai_learning_guard_controls(
    base_df,
    *,
    active_base: dict,
    status_cache: dict,
) -> None:
    """Compact manual-intervention recorder shared by analysis/dispatch pages."""
    if base_df is None or getattr(base_df, "empty", True) or "場站名稱" not in base_df.columns:
        return

    station_names = [
        name
        for name in dict.fromkeys(str(value or "").strip() for value in base_df["場站名稱"].tolist())
        if name
    ]
    if not station_names:
        return

    learning = _ai_learning_meta(status_cache)
    manual_events = learning.setdefault("manual_events", [])
    if not isinstance(manual_events, list):
        manual_events = []
        learning["manual_events"] = manual_events
    summary = learning.get("last_summary", {})
    if not isinstance(summary, dict):
        summary = {}

    token = str(active_base.get("token") or "default")
    ai_context = st.session_state.get("ai_shift_context", {})
    with st.expander("🛠️ AI 人工調度紀錄", expanded=False):
        st.caption(
            "有自行調度時記一筆即可。正數＝補進場站，負數＝從場站載走；"
            "兩欄都填 0 也可只標記『此站有人工作業』。"
        )
        if summary:
            st.caption(
                "最近一次同步分類｜"
                f"自然 {int(summary.get('natural', 0))}｜"
                f"人工 {int(summary.get('manual_intervention', 0))}｜"
                f"疑似人工 {int(summary.get('suspected_intervention', 0))}"
            )

        station_name = st.selectbox(
            "場站",
            station_names,
            key=f"ai_manual_station::{token}",
        )
        c1, c2 = st.columns(2)
        with c1:
            bike_delta = int(
                st.number_input(
                    "2.0 變化",
                    min_value=-30,
                    max_value=30,
                    value=0,
                    step=1,
                    key=f"ai_manual_bike_delta::{token}",
                )
            )
        with c2:
            ebike_delta = int(
                st.number_input(
                    "2.0E 變化",
                    min_value=-30,
                    max_value=30,
                    value=0,
                    step=1,
                    key=f"ai_manual_ebike_delta::{token}",
                )
            )

        if st.button(
            "記錄人工調度",
            use_container_width=True,
            key=f"ai_manual_save::{token}",
        ):
            event = build_manual_intervention_event(
                station_name=station_name,
                bike_delta=bike_delta,
                ebike_delta=ebike_delta,
                ai_context=ai_context,
            )
            manual_events.append(event)
            learning["manual_events"] = manual_events[-MAX_MANUAL_EVENTS:]
            save_cached_status(
                active_base["token"],
                active_base.get("expires_at"),
                status_cache,
            )
            st.success(f"已標記人工調度：{station_name}")

        recent = [item for item in manual_events if isinstance(item, dict)][-3:]
        if recent:
            st.caption("最近人工紀錄")
            for event in reversed(recent):
                try:
                    when = datetime.fromtimestamp(
                        float(event.get("recorded_at_epoch") or 0),
                        AI_TIMEZONE,
                    ).strftime("%H:%M:%S")
                except (TypeError, ValueError, OSError):
                    when = "—"
                bike = int(event.get("bike_delta") or 0)
                ebike = int(event.get("ebike_delta") or 0)
                used = "｜已套用" if event.get("consumed") else "｜待下次同步"
                st.caption(
                    f"{when}｜{event.get('station_name', '')}｜"
                    f"2.0 {bike:+d}｜2.0E {ebike:+d}{used}"
                )


def render_floating_battery_query(
    route_station_map: dict[str, list[dict]],
    mobile_mode: bool,
) -> None:
    """Render V29 Server battery UI in the legacy floating-button position."""
    _render_floating_server_battery(route_station_map, mobile_mode)

    # battery_upgrade owns the data/UI engine. This tiny one-way HTML shim only
    # preserves the legacy physical slot and removes any hot-reload leftovers.
    # ``components`` is imported by legacy_ui.py before any call reaches here.
    icon_html = r'''
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
              padding: 0 !important;
              overflow: hidden !important;
              border: 1px solid rgba(96, 232, 255, .65) !important;
              background: #07101c !important;
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
            fab.innerHTML = '<img alt="" draggable="false" src="__BATTERY_ICON_DATA_URI__" style="width:100%;height:100%;display:block;object-fit:cover;border-radius:15px;pointer-events:none;user-select:none;-webkit-user-drag:none;">';
            fab.title = '查詢 YouBike 2.0E 電量';
            fab.setAttribute('aria-label', '電量查詢');
          };

          const pillarSortKey = (row) => {
            const label = row.querySelector('span')?.textContent || '';
            const match = label.match(/\d+/);
            return match ? Number(match[0]) : Number.MAX_SAFE_INTEGER;
          };

          const sortBatteryRowsByPillar = () => {
            doc.querySelectorAll('#ubike-battery-v29-upgrade .bike-list').forEach(list => {
              const current = Array.from(list.children).filter(node => node.classList?.contains('bike'));
              if (current.length < 2) return;
              const sorted = [...current].sort((a, b) => {
                const diff = pillarSortKey(a) - pillarSortKey(b);
                if (diff) return diff;
                const aText = a.querySelector('span')?.textContent || '';
                const bText = b.querySelector('span')?.textContent || '';
                return aText.localeCompare(bText, 'zh-Hant', { numeric: true, sensitivity: 'base' });
              });
              const changed = current.some((node, index) => node !== sorted[index]);
              if (!changed) return;
              const firstNonBike = Array.from(list.children).find(node => !node.classList?.contains('bike')) || null;
              sorted.forEach(node => list.insertBefore(node, firstNonBike));
            });
          };

          applyLabel();
          sortBatteryRowsByPillar();
          window.setTimeout(applyLabel, 80);
          window.setTimeout(applyLabel, 300);
          window.setTimeout(sortBatteryRowsByPillar, 100);
          window.setTimeout(sortBatteryRowsByPillar, 500);

          const batteryRoot = doc.body;
          if (batteryRoot && !window.parent.__ubikePillarSortObserver) {
            let sortTimer = null;
            const observer = new MutationObserver(() => {
              window.clearTimeout(sortTimer);
              sortTimer = window.setTimeout(sortBatteryRowsByPillar, 30);
            });
            observer.observe(batteryRoot, { childList: true, subtree: true });
            window.parent.__ubikePillarSortObserver = observer;
          }
        })();
        </script>
        '''
    components.html(
        icon_html.replace("__BATTERY_ICON_DATA_URI__", BATTERY_ICON_DATA_URI),
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
    '''battery_route_map = merge_battery_route_station_maps(
    build_battery_route_station_map(base_df),
    DEFAULT_BATTERY_ROUTE_STATION_MAP,
)''',
    '''battery_route_map = build_battery_route_station_map(base_df)''',
    label="uploaded-workbook battery range",
)

# General-analysis rows reserve a compact AI prediction line now. Until the
# learning database/model is connected, show an explicit learning state rather
# than inventing a forecast. The same slot can later render 30m/risk output.
replace_exact(
    """            f'<small>{"｜".join(station_meta)}</small></td>'""",
    """            f'<small>{"｜".join(station_meta)}</small>'
            f'<small class="analysis-ai-prediction" style="color:#55f6ff;opacity:.92;font-weight:850;">🔮 AI 預測：學習中</small></td>'""",
    label="general analysis AI prediction slot",
)

replace_exact(
    '''st.set_page_config(
    page_title=f"臺東 YouBike 智慧調度｜{APP_VERSION_NAME}",
    page_icon="🚚",
    layout="wide",
)''',
    '''st.set_page_config(
    page_title=f"臺東 YouBike 智慧調度｜{APP_VERSION_NAME}",
    page_icon="🚚",
    layout="wide",
)

_UPDATE_CONTENT_MD = """
#### V29 更新內容
- 電池查詢範圍支援 Excel 任意區域，不再限制 D1／D2／D3。
- 上傳外縣市 Excel 時，不會混入台東內建備援場站。
- 未上傳配置表時，仍保留台東備援電量查詢。
- 場站即時車數使用 V29 同步架構，不再依賴手機隱藏同步元件。
- 右側更新按鈕可重新取得即時場站資料。
- 電池查詢已升級為 V29 Fast Client：並行查詢、逐站回填，不阻塞主畫面。
- 電池場站展開後，低電車明細依柱號由小到大排列。
- AI 班別直接跟隨主頁班別；早班／晚班用當天，大夜用跨日後的營運日判斷平日／假日。
- 一般分析的場站列已加入 AI 預測位置；模型尚未接入時明確顯示「學習中」。
- AI 學習防污染：自然流量、人工調度、疑似人工調度分開標記；人工資料不進自然需求訓練。
- iPhone 定位改為可見的直接定位按鈕；第一次由使用者點擊授權，成功後再進行背景更新。
- 新版電池入口沿用舊按鈕位置，並保留新版電池圖示。
"""
if hasattr(st, "popover"):
    with st.popover("更新內容"):
        st.markdown(_UPDATE_CONTENT_MD)
else:
    with st.expander("更新內容", expanded=False):
        st.markdown(_UPDATE_CONTENT_MD)''',
    label="update content popover",
)

replace_exact(
    '''    selected_shift = st.selectbox(
        "班別",
        list(SHIFT_COLUMNS.keys()),
        key=f"shift::{active_base['token']}",
    )
    page_mode = st.radio(''',
    '''    selected_shift = st.selectbox(
        "班別",
        list(SHIFT_COLUMNS.keys()),
        key=f"shift::{active_base['token']}",
    )
    _ai_shift_context = resolve_ai_shift_context(selected_shift)
    st.session_state["ai_shift_context"] = _ai_shift_context
    st.caption(
        f"🤖 AI 模式：{_ai_shift_context['day_type']}・{_ai_shift_context['shift']}"
    )
    page_mode = st.radio(''',
    label="AI shift day context",
)

replace_exact(
    '''render_context_strip(
    route=f"{selected_configuration_type}｜D1／D2／D3",
    shift=selected_shift,
    station_count=len(base_df),
    page_mode=page_mode,
    live_meta=previous_live_meta,
)
render_binding_vehicle_requirements(base_df, selected_shift=selected_shift)''',
    '''render_context_strip(
    route=f"{selected_configuration_type}｜D1／D2／D3",
    shift=selected_shift,
    station_count=len(base_df),
    page_mode=page_mode,
    live_meta=previous_live_meta,
)
render_ai_learning_guard_controls(
    base_df,
    active_base=active_base,
    status_cache=status_cache,
)
render_binding_vehicle_requirements(base_df, selected_shift=selected_shift)''',
    label="AI manual intervention recorder",
)

replace_exact(
    '''                    else:
                        base_df = live_updated_df
                        live_event_id = str(live_payload.get("event_id") or browser_event_id or "")
                        common_live_meta = {''',
    '''                    else:
                        previous_ai_df = base_df.copy(deep=True)
                        base_df = live_updated_df
                        live_event_id = str(live_payload.get("event_id") or browser_event_id or "")

                        ai_learning_meta = _ai_learning_meta(status_cache)
                        ai_transition = classify_live_transition(
                            previous_ai_df,
                            base_df,
                            manual_events=ai_learning_meta.get("manual_events", []),
                            ai_context=st.session_state.get("ai_shift_context", {}),
                            observed_at_epoch=time.time(),
                            source_event_id=live_event_id,
                        )
                        ai_learning_meta["manual_events"] = ai_transition["manual_events"]
                        changed_ai_records = [
                            record
                            for record in ai_transition["records"]
                            if (
                                record.get("classification") in {
                                    "manual_intervention",
                                    "suspected_intervention",
                                }
                                or record.get("bike_delta") not in (None, 0)
                                or record.get("ebike_delta") not in (None, 0)
                            )
                        ]
                        existing_ai_records = ai_learning_meta.get("transitions", [])
                        if not isinstance(existing_ai_records, list):
                            existing_ai_records = []
                        existing_ai_records.extend(changed_ai_records)
                        ai_learning_meta["transitions"] = trim_learning_records(existing_ai_records)
                        ai_learning_meta["last_summary"] = ai_transition["summary"]
                        ai_learning_meta["last_observed_at_epoch"] = ai_transition["observed_at_epoch"]

                        common_live_meta = {''',
    label="AI live transition classification",
)

# iPhone/in-app-browser geolocation fix. Do not schedule background geolocation
# before the user has explicitly interacted with the visible location control.
replace_exact(
    '''  function scheduleAutoLocate() {
    clearAutoTimer();
    if (!args.auto_refresh) return;
    const seconds = Math.max(10, Math.min(300, Number(args.auto_refresh_seconds || 30)));
    autoTimer = window.setTimeout(() => {
      autoTimer = null;
      if (busy) scheduleAutoLocate();
      else runLocate({ automatic: true });
    }, seconds * 1000);
  }''',
    '''  function scheduleAutoLocate() {
    clearAutoTimer();
    if (!args.auto_refresh || !autoStarted) return;
    const seconds = Math.max(10, Math.min(300, Number(args.auto_refresh_seconds || 30)));
    autoTimer = window.setTimeout(() => {
      autoTimer = null;
      if (busy) scheduleAutoLocate();
      else runLocate({ automatic: true });
    }, seconds * 1000);
  }''',
    label="geolocation wait for direct user gesture",
)

replace_exact(
    '''      error => {
        const message = error && error.message ? error.message : "定位失敗";
        setValue({
          ok: false,
          event_id: eventId(),
          request_token: String(args.request_token || ""),
          error: message,
        });
        setStatus(`定位失敗：${message}`, true);
        busy = false;
        button.disabled = false;
        button.textContent = "📍 再試一次";
        scheduleAutoLocate();
      },''',
    '''      error => {
        const code = Number(error?.code || 0);
        let message = error && error.message ? error.message : "定位失敗";
        if (code === 1) {
          message = "位置權限被拒絕；請到 iPhone 設定開啟此瀏覽器／App 的位置權限後再試一次";
        } else if (code === 2) {
          message = "目前無法取得位置；請確認定位服務已開啟並稍後再試";
        } else if (code === 3) {
          message = "定位逾時；請到室外或靠近窗邊後再試一次";
        }
        setValue({
          ok: false,
          event_id: eventId(),
          request_token: String(args.request_token || ""),
          error: message,
        });
        setStatus(`定位失敗：${message}`, true);
        busy = false;
        button.disabled = false;
        button.textContent = "📍 再試一次";
        scheduleAutoLocate();
      },''',
    label="geolocation readable mobile errors",
)

replace_exact(
    '  button.addEventListener("click", () => runLocate());',
    '''  button.addEventListener("click", () => {
    autoStarted = true;
    runLocate({ automatic: false, forceDelivery: true });
  });''',
    label="geolocation direct click",
)

replace_exact(
    '''            auto_start=True,
            auto_refresh=True,
            auto_refresh_seconds=SHARED_GEOLOCATION_REFRESH_SECONDS,
            compact=True,''',
    '''            auto_start=bool(st.session_state.get(state_key)),
            auto_refresh=True,
            auto_refresh_seconds=SHARED_GEOLOCATION_REFRESH_SECONDS,
            compact=False,''',
    label="visible geolocation component",
)

replace_exact(
    '''# 配置表一載入就開始定位，之後每 30 秒在背景更新一次。
shared_location = render_shared_geolocation(active_base)
with st.sidebar:
    render_shared_location_summary(active_base, shared_location)
    if st.button(
        "立即更新定位",
        use_container_width=True,
        key=f"sidebar_refresh_location::{active_base['token']}",
    ):
        request_shared_geolocation_refresh(active_base)
        rerun_app()''',
    '''# 第一次定位必須由使用者直接點擊；授權成功後保留背景更新能力。
with st.sidebar:
    st.caption("📍 目前位置")
    st.caption("第一次請直接按下方定位按鈕；成功後系統會自動更新距離與智慧調度路線。")
    shared_location = render_shared_geolocation(active_base)
    render_shared_location_summary(active_base, shared_location)''',
    label="sidebar direct geolocation control",
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
    '''def get_youbike_browser_sync_component():
    """V29 compatibility: obtain live station data from the Python Server."""
    from live_status_service import LiveStatusServiceError, get_live_status_for_stations

    def _server_sync_component(**_kwargs):
        stations = globals().get("_V29_SERVER_LIVE_STATIONS", [])
        if not stations:
            return {
                "ok": False,
                "event_id": uuid.uuid4().hex,
                "error": "目前配置沒有可供同步的場站。",
            }

        refresh_token = ""
        try:
            refresh_token = str(st.query_params.get("live_refresh", "") or "").strip()
        except Exception:
            refresh_token = ""
        refresh_state_key = "v29_server_live_refresh_token"
        force_refresh = bool(
            refresh_token
            and st.session_state.get(refresh_state_key) != refresh_token
        )
        if force_refresh:
            st.session_state[refresh_state_key] = refresh_token

        try:
            return get_live_status_for_stations(stations, force=force_refresh)
        except LiveStatusServiceError as exc:
            return {
                "ok": False,
                "event_id": uuid.uuid4().hex,
                "error": str(exc),
            }
        except Exception as exc:
            return {
                "ok": False,
                "event_id": uuid.uuid4().hex,
                "error": f"Server 即時車數同步失敗：{exc}",
            }

    return _server_sync_component


def normalize_browser_live_payload(payload) -> dict:''',
    label="server live component adapter",
)

# Build the exact station scope from the currently selected/uploaded workbook.
# This also makes the server sync automatically follow cross-county Excel data.
replace_exact(
    '    browser_payload = None',
    '''    _V29_SERVER_LIVE_STATIONS = [
        {
            "name": str(row.get("場站名稱") or "").strip(),
            "district": str(row.get("行政區") or "").strip(),
        }
        for _, row in base_df.iterrows()
        if str(row.get("場站名稱") or "").strip()
    ]
    browser_payload = None''',
    label="server live station scope",
)

# The legacy floating refresh button used to require discovery of a hidden
# Streamlit iframe. With the V29 server adapter there is intentionally no iframe.
# A refresh token causes one forced server fetch on the next Streamlit run.
replace_exact(
    '''                if (!postedCount) {{
                    showToast("同步元件尚未準備完成，請稍後再按一次");
                    return;
                }}
                setRefreshButtonState(true);''',
    '''                if (!postedCount) {{
                    setRefreshButtonState(true);
                    showToast("正在重新同步 YouBike 即時資料…");
                    try {{
                        const refreshUrl = new URL(win.location.href);
                        refreshUrl.searchParams.set("live_refresh", String(Date.now()));
                        win.location.href = refreshUrl.toString();
                    }} catch (_refreshError) {{
                        win.location.reload();
                    }}
                    return;
                }}
                setRefreshButtonState(true);''',
    label="server live floating refresh",
)

exec(compile(source, str(LEGACY_APP), "exec"), globals(), globals())
