from __future__ import annotations

V29_COMPATIBILITY_NOTES = """V29 old-UI compatibility entrypoint.

The original legacy UI is kept byte-for-byte in ``legacy_ui.py``. Before
executing it, this entrypoint applies focused V29 compatibility fixes:

1. battery ranges accept any non-empty Excel zone instead of D1/D2/D3 only;
2. an uploaded workbook never gets merged with the built-in Taitung fallback;
3. Taitung fallback remains available only in the existing no-workbook path;
4. live station sync uses the V29 Python Server service instead of requiring a
   hidden browser Streamlit component on mobile;
5. the floating refresh button requests a fresh server sync through the existing
   Streamlit geolocation bridge, without reloading the whole browser page;
6. the floating battery query uses the V29 Fast Client battery engine and a
   mobile-safe one-way HTML UI, avoiding custom-component readiness failures;
7. the V29 battery entry occupies the exact legacy battery-button slot so the
   new engine replaces the old entry instead of appearing as a second control;
8. AI learning guard separates natural demand, confirmed manual intervention
   and suspected intervention before future model training;
9. geolocation uses a visible direct user-triggered control before background
   refresh, improving iPhone/in-app-browser permission reliability.
"""

import json
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
          const win = window.parent;
          const doc = win.document;
          const repairBatteryFab = () => {
            try { win.__ubikeV29FastBattery?.repairFloatingButton?.(); } catch (_) {}
          };
          repairBatteryFab();
          win.setTimeout(repairBatteryFab, 60);
          win.setTimeout(repairBatteryFab, 250);
          win.setTimeout(repairBatteryFab, 900);
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


PRIORITY_STATION_META_KEY = "__priority_stations_v1__"
PRIORITY_COMPLETED_LIMIT = 80


def _priority_station_key(value) -> str:
    """建立優先場站比對鍵；保留原始顯示名稱，只用於安全比對。"""
    text = str(value or "").strip().casefold().replace("臺", "台")
    return "".join(text.split())


def _priority_scope_id(selected_shift: str = "") -> str:
    """優先清單以營運日＋班別分開，避免隔天仍把昨天已完成任務算進本班。"""
    context = st.session_state.get("ai_shift_context", {})
    if not isinstance(context, dict):
        context = {}
    operating_date = str(context.get("operating_date") or "").strip()
    if not operating_date:
        operating_date = datetime.now(AI_TIMEZONE).date().isoformat()
    shift_label = str(
        selected_shift
        or context.get("source_shift")
        or context.get("shift")
        or "未設定班別"
    ).strip()
    return f"{operating_date}｜{shift_label}"


def _priority_store(status_cache: dict) -> dict:
    metadata = status_cache.setdefault("metadata", {})
    store = metadata.get(PRIORITY_STATION_META_KEY)
    if not isinstance(store, dict):
        store = {}
        metadata[PRIORITY_STATION_META_KEY] = store
    return store


def _priority_bucket(
    status_cache: dict,
    selected_shift: str = "",
    *,
    create: bool = True,
) -> dict:
    store = _priority_store(status_cache)
    scope_id = _priority_scope_id(selected_shift)
    bucket = store.get(scope_id)
    if not isinstance(bucket, dict):
        if not create:
            return {"pending": [], "completed": []}
        bucket = {"pending": [], "completed": [], "updated_at_epoch": time.time()}
        store[scope_id] = bucket

    pending_raw = bucket.get("pending", [])
    completed_raw = bucket.get("completed", [])
    if not isinstance(pending_raw, list):
        pending_raw = []
    if not isinstance(completed_raw, list):
        completed_raw = []

    pending: list[dict] = []
    seen_pending: set[str] = set()
    for raw in pending_raw:
        item = {"station_name": raw} if isinstance(raw, str) else (dict(raw) if isinstance(raw, dict) else {})
        station_name = str(item.get("station_name") or "").strip()
        station_key = _priority_station_key(station_name)
        if not station_name or not station_key or station_key in seen_pending:
            continue
        seen_pending.add(station_key)
        item["station_name"] = station_name
        item.setdefault("note", "")
        item.setdefault("source", "manual")
        pending.append(item)

    completed: list[dict] = []
    for raw in completed_raw:
        item = {"station_name": raw} if isinstance(raw, str) else (dict(raw) if isinstance(raw, dict) else {})
        station_name = str(item.get("station_name") or "").strip()
        if not station_name:
            continue
        item["station_name"] = station_name
        item.setdefault("note", "")
        completed.append(item)

    bucket["pending"] = pending
    bucket["completed"] = completed[-PRIORITY_COMPLETED_LIMIT:]
    return bucket


def _save_priority_state(
    *,
    status_cache: dict,
    active_base_token: str,
    selected_shift: str = "",
) -> None:
    bucket = _priority_bucket(status_cache, selected_shift)
    bucket["updated_at_epoch"] = time.time()
    save_cached_status(active_base_token, None, status_cache)


def _priority_pending_names(status_cache: dict, selected_shift: str = "") -> list[str]:
    bucket = _priority_bucket(status_cache, selected_shift, create=False)
    return [
        str(item.get("station_name") or "").strip()
        for item in bucket.get("pending", [])
        if isinstance(item, dict) and str(item.get("station_name") or "").strip()
    ]


def priority_station_rank(
    station_name: str,
    status_cache: dict,
    selected_shift: str = "",
) -> int | None:
    wanted = _priority_station_key(station_name)
    if not wanted:
        return None
    for index, name in enumerate(_priority_pending_names(status_cache, selected_shift), start=1):
        if _priority_station_key(name) == wanted:
            return index
    return None


def apply_priority_station_order(
    candidates: list[dict],
    status_cache: dict,
    selected_shift: str = "",
) -> list[dict]:
    """優先清單只標記必處理站，不採用人工清單順序決定路線。

    智慧調度仍以原本 AI 候選排序為基礎：車上 2.0／2.0E、剩餘載量、
    場站缺多、道路時間與路線預看都先算完；可執行的優先站會集中在
    一般站之前，但多個優先站彼此之間完全沿用 AI 算出的效率順序。
    """
    if not candidates:
        return candidates

    priority_keys = {
        _priority_station_key(name)
        for name in _priority_pending_names(status_cache, selected_shift)
    }
    prepared: list[dict] = []
    for ai_rank, candidate in enumerate(candidates, start=1):
        item = dict(candidate)
        is_priority = (
            _priority_station_key(item.get("station_name"))
            in priority_keys
        )
        item["_ai_rank"] = ai_rank
        item["_is_priority_task"] = bool(is_priority)
        # 保留欄位給既有 UI 相容，但不再使用人工清單名次排序。
        item["_priority_rank"] = 1 if is_priority else 0
        prepared.append(item)

    # 關鍵：優先站之間只依 AI 原始排名，不依使用者在待辦清單的上下順序。
    prepared.sort(
        key=lambda item: (
            0 if bool(item.get("_is_priority_task")) else 1,
            int(item.get("_ai_rank") or 999999),
        )
    )
    return prepared


def priority_mandatory_route_state(
    candidates: list[dict],
    status_cache: dict,
    selected_shift: str = "",
) -> dict:
    """回報本班必完成任務是否已有可立即執行的場站。"""
    pending_names = _priority_pending_names(status_cache, selected_shift)
    pending_keys = {
        _priority_station_key(name)
        for name in pending_names
    }
    executable = [
        candidate
        for candidate in candidates
        if _priority_station_key(candidate.get("station_name")) in pending_keys
    ]
    return {
        "pending_count": len(pending_names),
        "executable_count": len(executable),
        "has_pending": bool(pending_names),
        "has_executable": bool(executable),
    }


def _priority_general_preference_keys(
    active_base_token: str,
    selected_shift: str,
) -> tuple[str, str]:
    scope_id = _priority_scope_id(selected_shift)
    return (
        f"priority_pin::{active_base_token}::{scope_id}",
        f"priority_only::{active_base_token}::{scope_id}",
    )


def _priority_general_preferences(
    active_base_token: str,
    selected_shift: str,
) -> tuple[bool, bool]:
    pin_key, only_key = _priority_general_preference_keys(
        active_base_token,
        selected_shift,
    )
    return (
        bool(st.session_state.get(pin_key, True)),
        bool(st.session_state.get(only_key, False)),
    )


def _priority_status_lookup(status_df) -> dict[str, object]:
    lookup: dict[str, object] = {}
    if status_df is None or getattr(status_df, "empty", True):
        return lookup
    if "場站名稱" not in status_df.columns:
        return lookup
    for _, row in status_df.iterrows():
        station_name = str(row.get("場站名稱") or "").strip()
        station_key = _priority_station_key(station_name)
        if station_name and station_key:
            lookup[station_key] = row
    return lookup


def _priority_status_badge_html(row, label: str, current_column: str, standard_column: str) -> str:
    """把缺／多狀態壓成高辨識度小標籤，保留資料未取得狀態。"""
    if row is None:
        status_text = "資料未取得"
    else:
        status_text = format_dispatch_status(
            row.get(current_column),
            row.get(standard_column),
        )

    status_text = str(status_text or "資料未取得")
    if "缺" in status_text or "少" in status_text:
        background = "#fee2e2"
        foreground = "#b91c1c"
        border = "#fecaca"
        icon = "▼"
    elif "多" in status_text:
        background = "#dcfce7"
        foreground = "#166534"
        border = "#bbf7d0"
        icon = "▲"
    elif "符合" in status_text:
        background = "#eef2f7"
        foreground = "#475569"
        border = "#d7dee8"
        icon = "●"
    else:
        background = "#fff7ed"
        foreground = "#b45309"
        border = "#fed7aa"
        icon = "?"
    return (
        '<span style="display:inline-flex;align-items:center;gap:.24rem;'
        'padding:.16rem .46rem;border-radius:999px;'
        f'background:{background};color:{foreground};border:1px solid {border};'
        'font-size:.74rem;font-weight:950;line-height:1.2;white-space:nowrap;">'
        f'{icon} {html.escape(label)} {html.escape(status_text.replace(" 台", ""))}</span>'
    )


def _priority_status_badges_html(row) -> str:
    bike_badge = _priority_status_badge_html(
        row,
        "2.0",
        "2.0 現況",
        "2.0 標準",
    )
    ebike_badge = _priority_status_badge_html(
        row,
        "2.0E",
        "2.0E 現況",
        "2.0E 標準",
    )
    location_text = ""
    if row is not None:
        route_zone = str(row.get("路線區域") or "").strip()
        district = str(row.get("行政區") or "").strip()
        location_text = "｜".join(part for part in (route_zone, district) if part)
    location_html = (
        f'<span style="font-size:.67rem;opacity:.58;white-space:nowrap;">'
        f'{html.escape(location_text)}</span>'
        if location_text else ""
    )
    return (
        '<div style="display:flex;align-items:center;gap:.3rem;'
        'flex-wrap:wrap;margin-top:.24rem;">'
        f'{bike_badge}{ebike_badge}{location_html}</div>'
    )


def render_priority_station_manager(
    status_df,
    *,
    status_cache: dict,
    active_base_token: str,
    selected_shift: str,
    page_mode: str,
) -> dict:
    """智慧調度／一般分析共用的人工派工待辦（精簡介面）。"""
    bucket = _priority_bucket(status_cache, selected_shift)
    pending = bucket["pending"]
    completed = bucket["completed"]
    scope_id = _priority_scope_id(selected_shift)
    row_lookup = _priority_status_lookup(status_df)

    station_names: list[str] = []
    if (
        status_df is not None
        and not getattr(status_df, "empty", True)
        and "場站名稱" in status_df.columns
    ):
        station_names = [
            name
            for name in dict.fromkeys(
                str(value or "").strip()
                for value in status_df["場站名稱"].tolist()
            )
            if name
        ]

    priority_pin = True
    priority_only = False
    if page_mode == "一般分析":
        pin_key, only_key = _priority_general_preference_keys(
            active_base_token,
            selected_shift,
        )
        toggle = getattr(st, "toggle", st.checkbox)
        pref_col_1, pref_col_2 = st.columns(2)
        with pref_col_1:
            priority_pin = bool(
                toggle(
                    "🚨 優先置頂",
                    value=True,
                    key=pin_key,
                    help="優先場站只顯示在上方待辦區，不在一般場站重複出現。",
                )
            )
        with pref_col_2:
            priority_only = bool(
                toggle(
                    "只看優先",
                    value=False,
                    key=only_key,
                    help="只顯示本班尚未完成的人工派工場站。",
                )
            )

    with st.expander(
        f"🚨 優先場站｜{len(pending)} 站未完成",
        expanded=False,
    ):
        if pending:
            st.caption("優先場站＝本班下班前要完成｜紅＝缺車｜綠＝多車｜灰＝符合｜↑↓只調整待辦顯示，不影響智慧調度路線")
            for index, item in enumerate(list(pending)):
                station_name = str(item.get("station_name") or "").strip()
                station_key = _priority_station_key(station_name)
                row = row_lookup.get(station_key)
                note = str(item.get("note") or "").strip()

                info_col, up_col, down_col, done_col = st.columns(
                    [6.4, .72, .72, 1.15],
                    gap="small",
                )
                with info_col:
                    st.markdown(
                        (
                            '<div style="padding:.38rem .52rem;'
                            'border:1px solid rgba(148,163,184,.20);'
                            'border-radius:10px;background:rgba(255,255,255,.72);'
                            'margin:0 0 .12rem;">'
                            f'<div style="font-size:.88rem;font-weight:950;'
                            f'line-height:1.2;">{index + 1}. 🚨 {html.escape(station_name)}</div>'
                            f'{_priority_status_badges_html(row)}'
                            '</div>'
                        ),
                        unsafe_allow_html=True,
                    )
                    if note:
                        st.caption(f"📝 {note}")

                with up_col:
                    move_up = st.button(
                        "↑",
                        use_container_width=True,
                        disabled=index == 0,
                        key=f"priority_up::{active_base_token}::{scope_id}::{station_key}",
                        help="提高優先順序",
                    )
                with down_col:
                    move_down = st.button(
                        "↓",
                        use_container_width=True,
                        disabled=index >= len(pending) - 1,
                        key=f"priority_down::{active_base_token}::{scope_id}::{station_key}",
                        help="降低優先順序",
                    )
                with done_col:
                    mark_done = st.button(
                        "✓",
                        type="primary",
                        use_container_width=True,
                        key=f"priority_done::{active_base_token}::{scope_id}::{station_key}",
                        help="完成此優先場站",
                    )

                if move_up and index > 0:
                    pending[index - 1], pending[index] = pending[index], pending[index - 1]
                    _save_priority_state(
                        status_cache=status_cache,
                        active_base_token=active_base_token,
                        selected_shift=selected_shift,
                    )
                    rerun_app()
                if move_down and index < len(pending) - 1:
                    pending[index + 1], pending[index] = pending[index], pending[index + 1]
                    _save_priority_state(
                        status_cache=status_cache,
                        active_base_token=active_base_token,
                        selected_shift=selected_shift,
                    )
                    rerun_app()
                if mark_done:
                    completed_item = dict(pending.pop(index))
                    completed_item["completed_at_epoch"] = time.time()
                    completed_item["completed_shift"] = selected_shift
                    completed.append(completed_item)
                    bucket["completed"] = completed[-PRIORITY_COMPLETED_LIMIT:]
                    _save_priority_state(
                        status_cache=status_cache,
                        active_base_token=active_base_token,
                        selected_shift=selected_shift,
                    )
                    rerun_app()
        else:
            st.caption("本班目前沒有未完成的優先場站。")

        pending_keys = {
            _priority_station_key(item.get("station_name"))
            for item in pending
            if isinstance(item, dict)
        }
        available_names = [
            name
            for name in station_names
            if _priority_station_key(name) not in pending_keys
        ]

        selected_names: list[str] = []
        priority_note = ""
        add_submitted = False
        if hasattr(st, "popover"):
            with st.popover("＋ 新增優先場站", use_container_width=True):
                with st.form(
                    key=f"priority_add_form::{active_base_token}::{scope_id}::{page_mode}",
                    clear_on_submit=True,
                ):
                    selected_names = st.multiselect(
                        "選擇場站",
                        available_names,
                        placeholder="搜尋並勾選場站",
                    )
                    priority_note = st.text_input(
                        "備註（可留空）",
                        placeholder="例如：下班前處理",
                    )
                    add_submitted = st.form_submit_button(
                        "加入優先清單",
                        type="primary",
                        use_container_width=True,
                    )
        else:
            with st.form(
                key=f"priority_add_form::{active_base_token}::{scope_id}::{page_mode}",
                clear_on_submit=True,
            ):
                selected_names = st.multiselect(
                    "＋ 新增優先場站",
                    available_names,
                    placeholder="搜尋並勾選場站",
                )
                priority_note = st.text_input(
                    "備註（可留空）",
                    placeholder="例如：下班前處理",
                )
                add_submitted = st.form_submit_button(
                    "加入優先清單",
                    type="primary",
                    use_container_width=True,
                )

        if add_submitted:
            selected_names = [
                str(name).strip()
                for name in selected_names
                if str(name).strip()
            ]
            if not selected_names:
                st.warning("請至少選擇一個場站。")
            else:
                created_at = time.time()
                for station_name in selected_names:
                    station_key = _priority_station_key(station_name)
                    if any(
                        _priority_station_key(item.get("station_name")) == station_key
                        for item in pending
                        if isinstance(item, dict)
                    ):
                        continue
                    pending.append(
                        {
                            "station_name": station_name,
                            "note": str(priority_note or "").strip(),
                            "created_at_epoch": created_at,
                            "created_shift": selected_shift,
                            "source": "manual",
                        }
                    )
                _save_priority_state(
                    status_cache=status_cache,
                    active_base_token=active_base_token,
                    selected_shift=selected_shift,
                )
                rerun_app()

    if completed:
        with st.expander(f"✅ 本班已完成｜{len(completed)}", expanded=False):
            for reverse_index, item in enumerate(reversed(completed[-10:]), start=1):
                station_name = str(item.get("station_name") or "").strip()
                completed_epoch = float(item.get("completed_at_epoch") or 0)
                try:
                    completed_time = datetime.fromtimestamp(
                        completed_epoch,
                        AI_TIMEZONE,
                    ).strftime("%H:%M")
                except (TypeError, ValueError, OSError):
                    completed_time = "—"

                row_col, restore_col = st.columns([5, 1], gap="small")
                with row_col:
                    st.markdown(f"✓ **{station_name}**　（{completed_time}）")
                with restore_col:
                    restore_clicked = st.button(
                        "↩",
                        use_container_width=True,
                        key=(
                            f"priority_restore::{active_base_token}::{scope_id}::"
                            f"{reverse_index}::{_priority_station_key(station_name)}"
                        ),
                        help="復原到未完成",
                    )

                if restore_clicked:
                    restored = dict(item)
                    restored.pop("completed_at_epoch", None)
                    restored.pop("completed_shift", None)
                    if not any(
                        _priority_station_key(pending_item.get("station_name"))
                        == _priority_station_key(station_name)
                        for pending_item in pending
                        if isinstance(pending_item, dict)
                    ):
                        pending.append(restored)
                    completed.remove(item)
                    _save_priority_state(
                        status_cache=status_cache,
                        active_base_token=active_base_token,
                        selected_shift=selected_shift,
                    )
                    rerun_app()

    if pending:
        st.caption(f"🚨 本班尚有 {len(pending)} 個優先場站待完成")

    return {
        "pending_names": _priority_pending_names(status_cache, selected_shift),
        "priority_pin": priority_pin,
        "priority_only": priority_only,
    }


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

**2026/09/22｜優先場站＋智慧調度整合**
- 新增「優先場站」共用待辦：智慧調度與一般分析同步使用同一份清單，可多選新增、手動排序、完成、復原與收合；其用途仍是本班下班前要完成的場站。
- 一般分析支援「優先置頂」與「只看優先」，並把優先場站與一般場站分開顯示。
- 優先場站介面改成精簡模式：預設收合、單列操作，缺車用紅色高對比標籤、多車用綠色標籤；新增場站移到彈出視窗。
- 「優先場站」代表本班下班前一定要處理完成，不代表人工指定行車順序；↑↓ 只調整待辦顯示。
- 智慧調度會依車上 2.0／2.0E、剩餘載量、場站缺／多車、道路距離與效率，自動決定優先場站的最佳先後順序。
- 只要還有可立即執行的優先場站，就優先從優先任務中挑選下一站；若目前全部無法執行，才允許先安排一般準備站調整車上車種或載量，再回到優先場站任務。
- 優先場站只有手動按「完成」才會移出待辦，並保留本班完成時間與復原功能。

**既有 V29 更新**
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
    '''else:
    with st.expander("更新內容", expanded=False):
        st.markdown(_UPDATE_CONTENT_MD)''',
    '''else:
    with st.expander("更新內容", expanded=False):
        st.markdown(_UPDATE_CONTENT_MD)

_BUG_FIX_CONTENT_MD = """
#### 🐞 BUG 修復紀錄
> 僅記錄已實際完成的修復；單純新增功能不列入。

**2026/09/20｜自動檢修：V29 啟動修補目標失效**
- 自動逐一驗證 25 個 `replace_exact()` 修補，發現定位錯誤訊息修補的舊目標被誤改，可能導致重新部署／重啟時直接啟動失敗。
- 已修正修補目標並重新驗證，目前 25/25 全部可依序套用。

**2026/09/20｜更新＋定位事件競態**
- 修復懸浮更新先觸發 Streamlit rerun、再等待 GPS 的競態問題。
- 改為 GPS 成功或失敗回傳時，同一事件一起要求 YouBike 場站資料強制更新，避免定位結果在 iframe 重建時遺失。

**2026/09/20｜電池全螢幕懸浮工具遮擋**
- 修復新版電池 `open()` 結構變更後，相容層找不到舊字串，造成進入電池全螢幕時右側其他懸浮工具未隱藏。
- 相容層現在同時支援新版與舊版電池開啟函式。

**2026/09/19｜電池查詢懸浮按鈕在更新後消失**
- 修復 Streamlit rerun 後電池查詢 FAB 可能被 DOM 重建移除的問題。
- 電池 FAB 改由 parent window 永久守護，並加入 MutationObserver 與 750ms 自我修復。
- 主頁每次重跑也會主動要求電池模組立即重建／顯示按鈕。

**2026/09/19｜懸浮更新造成整頁重載、定位中斷**
- 修復右側懸浮「更新」會重新載入整個瀏覽器頁面的問題。
- 移除 `window.location.reload()` 與 `live_refresh` 網址跳轉刷新。
- 懸浮「更新」現在會同時刷新 YouBike 即時場站資料與 GPS 定位，但不重建整個頁面。

**2026/09/10｜電池懸浮按鈕偶發消失**
- 修復電池懸浮按鈕在手機／頁面重跑後偶發不見。
- 按鈕、頁面或主容器缺失時會自動重建，並強化顯示層級。
- 保留定位、距離排序與行政區需換電池統整。

**2026/09/10｜電池頁定位與排序顯示修復**
- 調整持續定位與依距離排序流程。
- 修正行政區需換電池統整顯示，避免更新後排序／統計不同步。

**2026/09/08｜V29 部署依賴失敗**
- 修復 Streamlit 部署時 Debian `bullseye-security` Release file 過期造成安裝失敗。
- 移除造成部署阻塞的 `packages.txt` apt 依賴設定，恢復 V29 部署。

**2026/09/05｜跨縣市 Excel 場站範圍錯誤**
- 修復電池查詢被硬限制只能使用 D1／D2／D3。
- 改為依 Excel 實際區域讀取場站。
- 上傳外縣市 Excel 時不再混入台東內建備援場站；只有未上傳 Excel 時才使用台東備援。

**2026/09/05｜手機「同步元件尚未準備完成」**
- 修復手機端依賴隱藏 iframe 同步元件造成無法更新即時車數。
- 場站即時車數改由 V29 Python Server 同步服務處理，後續不再依賴隱藏同步 iframe。

**2026/09/05｜手機電池查詢元件／重複入口**
- 修復手機電池查詢 custom component readiness 問題。
- 改用 V29 Server 電池引擎與 mobile-safe UI。
- 新版電池入口直接取代舊按鈕位置，避免畫面出現兩個電池入口。

**2026/08/30｜執行環境 PyArrow 穩定性**
- 調整 PyArrow 相容性，避開 25.x 世代導入期間的執行風險。
- 目前 Streamlit 1.59 的相依限制會維持 PyArrow 25 以下；24.0.0 屬可用版本。
"""

with st.sidebar:
    with st.expander("🐞 BUG修復內容", expanded=False):
        st.markdown(_BUG_FIX_CONTENT_MD)''',
    label="sidebar bug fix history",
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
          manual_live_refresh: Boolean(requestLiveRefresh),
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
    '  function runLocate({ automatic = false, forceDelivery = false } = {}) {',
    '  function runLocate({ automatic = false, forceDelivery = false, requestLiveRefresh = false } = {}) {',
    label="geolocation refresh bridge signature",
)

replace_exact(
    '''          request_token: String(args.request_token || ""),
          latitude: Number(position.coords.latitude),''',
    '''          request_token: String(args.request_token || ""),
          manual_live_refresh: Boolean(requestLiveRefresh),
          latitude: Number(position.coords.latitude),''',
    label="geolocation refresh bridge success payload",
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
    '''  window.addEventListener("message", event => {
    if (!event.data || event.data.type !== "streamlit:render") return;
    args = event.data.args || {};
    document.body.classList.toggle("compact", Boolean(args.compact));''',
    '''  window.addEventListener("message", event => {
    if (!event.data) return;
    if (event.data.type === "ubike:manual-locate-and-refresh") {
      autoStarted = true;
      runLocate({ automatic: false, forceDelivery: true, requestLiveRefresh: true });
      return;
    }
    if (event.data.type !== "streamlit:render") return;
    args = event.data.args || {};
    document.body.classList.toggle("compact", Boolean(args.compact));''',
    label="geolocation floating refresh listener",
)

replace_exact(
    '''    except Exception as exc:
        st.session_state[f"{prefix}::error"] = str(exc)

    if isinstance(payload, dict):''',
    '''    except Exception as exc:
        st.session_state[f"{prefix}::error"] = str(exc)

    if isinstance(payload, dict) and payload.get("manual_live_refresh"):
        st.session_state["v29_server_live_force_refresh"] = True

    if isinstance(payload, dict):''',
    label="geolocation refresh bridge state",
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

        force_refresh = bool(
            st.session_state.pop("v29_server_live_force_refresh", False)
        )

        # Backward compatibility for an old bookmarked live_refresh URL.
        refresh_token = ""
        try:
            refresh_token = str(st.query_params.get("live_refresh", "") or "").strip()
        except Exception:
            refresh_token = ""
        refresh_state_key = "v29_server_live_refresh_token"
        if (
            refresh_token
            and st.session_state.get(refresh_state_key) != refresh_token
        ):
            st.session_state[refresh_state_key] = refresh_token
            force_refresh = True

        def _emit_sync_state(state: str, *, station_count: int = 0, message: str = "") -> None:
            try:
                event_payload = json.dumps(
                    {
                        "source": "ubike-browser-sync",
                        "type": "ubike:sync-state",
                        "state": state,
                        "station_count": max(0, int(station_count or 0)),
                        "message": str(message or ""),
                    },
                    ensure_ascii=False,
                )
                components.html(
                    f"<script>window.parent.postMessage({event_payload}, '*');</script>",
                    height=0,
                    scrolling=False,
                )
            except Exception:
                pass

        try:
            result = get_live_status_for_stations(stations, force=force_refresh)
            if force_refresh:
                _emit_sync_state(
                    "success",
                    station_count=int(result.get("station_count") or 0),
                )
            return result
        except LiveStatusServiceError as exc:
            if force_refresh:
                _emit_sync_state("error", message=str(exc))
            return {
                "ok": False,
                "event_id": uuid.uuid4().hex,
                "error": str(exc),
            }
        except Exception as exc:
            if force_refresh:
                _emit_sync_state("error", message=str(exc))
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

# The floating refresh button now also asks the visible geolocation component
# for a fresh GPS fix. That component acts as a bidirectional Streamlit bridge:
# it triggers a normal Streamlit rerun, marks one forced server fetch, and keeps
# the browser page/session alive instead of using window.location reload.
replace_exact(
    '''            function requestManualSync() {{
                let postedCount = 0;
                for (const frame of doc.querySelectorAll("iframe")) {{''',
    '''            function requestManualSync() {{
                let locationPostedCount = 0;
                for (const frame of doc.querySelectorAll("iframe")) {{
                    try {{
                        if (!frame.contentWindow) continue;
                        const frameTitle = String(frame.getAttribute("title") || "").toLowerCase();
                        const frameSource = String(frame.getAttribute("src") || "").toLowerCase();
                        let isLocationFrame = frameTitle.includes("dispatch_geolocation")
                            || frameSource.includes("dispatch_geolocation");
                        try {{
                            isLocationFrame = isLocationFrame
                                || Boolean(frame.contentDocument?.getElementById("locateButton"));
                        }} catch (_accessError) {{
                            // 跨來源時改以 title／src 判斷。
                        }}
                        if (!isLocationFrame) continue;
                        frame.contentWindow.postMessage(
                            {{ type: "ubike:manual-locate-and-refresh" }},
                            "*",
                        );
                        locationPostedCount += 1;
                    }} catch (_error) {{
                        // 略過無法存取的其他 iframe。
                    }}
                }}

                let postedCount = 0;
                for (const frame of doc.querySelectorAll("iframe")) {{''',
    label="floating refresh location bridge",
)

replace_exact(
    '''                if (!postedCount) {{
                    showToast("同步元件尚未準備完成，請稍後再按一次");
                    return;
                }}
                setRefreshButtonState(true);''',
    '''                if (!postedCount) {{
                    if (!locationPostedCount) {{
                        setRefreshButtonState(false);
                        showToast("更新元件尚未準備完成，請稍後再按一次");
                        return;
                    }}
                    setRefreshButtonState(true);
                    showToast("正在更新 YouBike 即時資料與定位…");
                    win.clearTimeout(win.__ubikeManualSyncFallbackTimer);
                    win.__ubikeManualSyncFallbackTimer = win.setTimeout(() => {{
                        setRefreshButtonState(false);
                    }}, 45000);
                    return;
                }}
                setRefreshButtonState(true);''',
    label="server live floating refresh",
)


# Shared priority-station task list: render once after live data/alerts are ready,
# then let both work modes consume the same persisted state.
replace_exact(
    '''render_station_alert_summary(
    base_df,
    station_locations=combined_station_locations,
    current_location=shared_location,
    alerts=station_alerts,
)

if page_mode == "智慧調度":''',
    '''render_station_alert_summary(
    base_df,
    station_locations=combined_station_locations,
    current_location=shared_location,
    alerts=station_alerts,
)

priority_station_ui = render_priority_station_manager(
    base_df,
    status_cache=status_cache,
    active_base_token=active_base["token"],
    selected_shift=selected_shift,
    page_mode=page_mode,
)

if page_mode == "智慧調度":''',
    label="shared priority station manager",
)

replace_exact(
    '''    selected_shift: str,
    active_base_token: str,
) -> None:
    """局部重跑排序、報表及懸浮工具，避免一般分析操作重新載入整個系統。"""''',
    '''    selected_shift: str,
    active_base_token: str,
    all_status_df: pd.DataFrame,
    status_cache: dict,
) -> None:
    """局部重跑排序、報表及懸浮工具，避免一般分析操作重新載入整個系統。"""''',
    label="general analysis priority parameters",
)

replace_exact(
    '''    result_df = analysis_result_df

    with st.expander("排序設定", expanded=False):''',
    '''    result_df = analysis_result_df

    priority_pin, priority_only = _priority_general_preferences(
        active_base_token,
        selected_shift,
    )
    priority_pending_names = _priority_pending_names(
        status_cache,
        selected_shift,
    )
    priority_pending_keys = {
        _priority_station_key(name)
        for name in priority_pending_names
    }

    if priority_only:
        st.info("🚨 目前僅顯示本班尚未完成的優先場站。")
        priority_detail_df = build_analysis_result(all_status_df)
        if priority_pending_keys:
            priority_detail_df = priority_detail_df[
                priority_detail_df["場站名稱"].astype(str).map(
                    _priority_station_key
                ).isin(priority_pending_keys)
            ].reset_index(drop=True)
        else:
            priority_detail_df = priority_detail_df.iloc[0:0].copy()

        if priority_detail_df.empty:
            st.success("本班目前沒有未完成的優先場站。")
        else:
            render_dispatch_legend()
            render_analysis_result_table(priority_detail_df)

        st.caption("即時車數來源：YouBike 官網；優先場站必須手動按完成才會移除。")
        render_floating_station_search(
            priority_detail_df,
            mobile_mode,
            page_mode="一般分析",
        )
        render_floating_battery_query(
            battery_route_map,
            mobile_mode,
        )
        return

    if priority_pin and priority_pending_keys and not result_df.empty:
        result_df = result_df[
            ~result_df["場站名稱"].astype(str).map(
                _priority_station_key
            ).isin(priority_pending_keys)
        ].reset_index(drop=True)
        st.markdown("### 一般場站")

    with st.expander("排序設定", expanded=False):''',
    label="general analysis priority filter",
)

replace_exact(
    '''    selected_shift=selected_shift,
    active_base_token=active_base["token"],
)''',
    '''    selected_shift=selected_shift,
    active_base_token=active_base["token"],
    all_status_df=base_df,
    status_cache=status_cache,
)''',
    label="general analysis priority call args",
)

replace_exact(
    '''            active_loop_phase=active_loop_phase,
        )

    manual_station_name = str(st.session_state.get(manual_station_key) or "").strip()''',
    '''            active_loop_phase=active_loop_phase,
        )

    candidates = apply_priority_station_order(
        candidates,
        status_cache,
    )
    priority_route_state = priority_mandatory_route_state(
        candidates,
        status_cache,
    )

    manual_station_name = str(st.session_state.get(manual_station_key) or "").strip()''',
    label="smart dispatch priority candidate ordering",
)

replace_exact(
    '''    recommendation_title = "使用者指定下一站" if manual_station_name else "下一站最高效益推薦"''',
    '''    recommended_priority_rank = priority_station_rank(
        str(recommended.get("station_name") or ""),
        status_cache,
    )
    if priority_route_state.get("has_pending"):
        pending_count = safe_nonnegative_int(priority_route_state.get("pending_count"))
        executable_count = safe_nonnegative_int(priority_route_state.get("executable_count"))
        if executable_count:
            st.info(
                f"🚨 本班尚有 {pending_count} 個優先場站；這些站需在下班前完成。 "
                "目前有可執行任務，下一站會從優先場站中依車上載量、缺多車與道路效率決定。"
            )
        else:
            st.warning(
                f"🚨 本班尚有 {pending_count} 個優先場站待完成，但目前沒有可立即執行的優先站。"
                "智慧調度可先安排一般站調整車上 2.0／2.0E 或載量，之後再回到優先場站任務。"
            )

    recommendation_title = (
        "使用者指定下一站"
        if manual_station_name
        else (
            "🚨 優先場站｜AI 最佳下一站"
            if recommended_priority_rank is not None
            else (
                "🔄 準備調度｜為優先場站調整車況"
                if priority_route_state.get("has_pending")
                else "下一站最高效益推薦"
            )
        )
    )''',
    label="smart dispatch priority recommendation title",
)

replace_exact(
    '''            rank_text = "🤖 AI 首選" if rank == 1 else f"第 {rank} 名"''',
    '''            priority_rank = safe_nonnegative_int(candidate.get("_priority_rank"))
            ai_rank = safe_nonnegative_int(candidate.get("_ai_rank")) or rank
            if priority_rank:
                rank_text = f"🚨 優先場站｜AI 第 {ai_rank} 名"
            elif priority_route_state.get("has_pending"):
                rank_text = f"🔄 一般準備站｜AI 第 {ai_rank} 名"
            elif ai_rank == 1:
                rank_text = "🤖 AI 首選"
            else:
                rank_text = f"AI 第 {ai_rank} 名"''',
    label="smart dispatch priority candidate badge",
)

exec(compile(source, str(LEGACY_APP), "exec"), globals(), globals())
