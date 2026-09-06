from __future__ import annotations

"""Compatibility helpers shared by the maintained V30 entrypoint and legacy UI.

This module provides three focused compatibility layers:
1. safe Streamlit v1 component declaration for legacy ``exec()`` code;
2. centralized system-version / generation metadata and changelog rendering;
3. maintained interaction upgrades that can be applied without rewriting the
   large legacy UI file on every release.
"""

import inspect
from typing import Any

import streamlit as st
import streamlit.components.v1 as components


# ---------------------------------------------------------------------------
# System version / generation metadata
# ---------------------------------------------------------------------------
# V30 is the baseline. From now on, five completed functional adjustments make
# one generation. Changes 1-4 remain on the current generation; the fifth
# adjustment promotes the visible/system version by one generation.
VERSION_BASE_GENERATION = 30
VERSION_CHANGE_THRESHOLD = 5
POST_V30_CHANGES: tuple[tuple[str, str], ...] = (
    (
        "2026-09-07",
        "右側懸浮更新改為同時強制更新 YouBike 場站即時資料與目前 GPS 定位。",
    ),
    (
        "2026-09-07",
        "電池查詢新增每站獨立反向按鈕，可改看高於低電門檻、也就是不需要換電池的車號與柱號；反向模式套用明顯不同的綠色表格樣式避免誤判。",
    ),
    (
        "2026-09-07",
        "電池查詢第二門檻（紅色緊急門檻）預設值由 40% 調整為 69%。",
    ),
    (
        "2026-09-07",
        "進入全螢幕電量查詢時暫時隱藏右側懸浮工具與賈維斯狀態浮層，返回主畫面後自動恢復，避免手機畫面被遮擋。",
    ),
)

_COMPLETED_GENERATIONS_AFTER_V30 = len(POST_V30_CHANGES) // VERSION_CHANGE_THRESHOLD
VERSION_PENDING_CHANGE_COUNT = len(POST_V30_CHANGES) % VERSION_CHANGE_THRESHOLD
SYSTEM_GENERATION = VERSION_BASE_GENERATION + _COMPLETED_GENERATIONS_AFTER_V30
SYSTEM_VERSION = f"V{SYSTEM_GENERATION}"
SYSTEM_BUILD_DATE = "2026-09-07"
SYSTEM_VERSION_LABEL = f"{SYSTEM_VERSION}｜第{SYSTEM_GENERATION}代"
NEXT_SYSTEM_VERSION = f"V{SYSTEM_GENERATION + 1}"


def _build_post_v30_changelog() -> str:
    """Build released five-change generations plus the current pending queue."""
    if not POST_V30_CHANGES:
        return (
            f"#### {SYSTEM_VERSION} 調整進度｜0/{VERSION_CHANGE_THRESHOLD}\n"
            f"滿 {VERSION_CHANGE_THRESHOLD} 項功能調整後升級為 {NEXT_SYSTEM_VERSION}。\n"
        )

    parts: list[str] = []
    full_groups = len(POST_V30_CHANGES) // VERSION_CHANGE_THRESHOLD
    for group_index in range(full_groups, 0, -1):
        generation = VERSION_BASE_GENERATION + group_index
        start = (group_index - 1) * VERSION_CHANGE_THRESHOLD
        group = POST_V30_CHANGES[start : start + VERSION_CHANGE_THRESHOLD]
        parts.append(f"### 第{generation}代｜V{generation}")
        parts.append(f"**升版條件：完成 {VERSION_CHANGE_THRESHOLD} 項調整**")
        parts.append("")
        for date_text, description in group:
            parts.append(f"- {description}（{date_text}）")
        parts.append("")
        parts.append("---")
        parts.append("")

    pending = POST_V30_CHANGES[full_groups * VERSION_CHANGE_THRESHOLD :]
    if pending:
        parts.append(
            f"#### {SYSTEM_VERSION} 調整進度｜{len(pending)}/{VERSION_CHANGE_THRESHOLD}"
        )
        parts.append(
            f"再完成 {VERSION_CHANGE_THRESHOLD - len(pending)} 項功能調整後升級為 {NEXT_SYSTEM_VERSION}。"
        )
        parts.append("")
        for date_text, description in pending:
            parts.append(f"- {description}（{date_text}）")
        parts.append("")
        parts.append("---")
        parts.append("")
    return "\n".join(parts)


SYSTEM_CHANGELOG_MD = _build_post_v30_changelog() + """
### 第30代｜V30
**更新日期：2026-09-07**

- AI 預測正式改為 **60 分鐘後**的場站車量趨勢／估計值，並顯示信心水準與樣本數。
- AI 學習紀錄上限提升為 **50,000 筆**。
- 手機、平板、電腦共用同一個 AI 學習池，學習資料預設在背景自動同步。
- 新增共用調度員定位池：兩名使用者的有效 GPS 都可參與人工調度判定。
- 場站車量發生變化時，若調度員位於該站附近，AI 會自動標記為人工調度，避免污染自然需求學習。
- 保留人工調度手動標記，並維持自然／人工／疑似人工三類防污染機制。
- 系統版本欄位改由中央版本資料管理。
- 「更新內容」入口移至左側選單，並以「第幾代｜版本」方式保留歷代更新紀錄。
- 自 V30 起採 **每累積 5 項功能調整升 1 個版本** 的版本節奏。

---

### 第29代｜V29

- 電池查詢範圍支援 Excel 任意區域，不再限制 D1／D2／D3。
- 上傳外縣市 Excel 時，不會混入台東內建備援場站；未上傳配置時仍保留台東備援電量查詢。
- 場站即時車數改採 V29 Python Server 同步架構，降低手機隱藏元件造成的同步失敗。
- 電池查詢採 Fast Client 並行查詢與逐站回填，低電車明細依柱號排序。
- AI 班別跟隨主頁班別，並加入自然流量／人工調度／疑似人工調度分類。
- iPhone 定位改為首次直接授權後持續背景更新。
"""


# ---------------------------------------------------------------------------
# Streamlit v1 component declaration compatibility
# ---------------------------------------------------------------------------
_ORIGINAL_DECLARE_COMPONENT = components.declare_component
_ORIGINAL_COMPONENT_HTML = components.html
_COMPONENT_COMPAT_INSTALLED = False
_VERSION_UI_INSTALLED = False
_LOCATION_REFRESH_QUERY_KEY = "location_refresh"
_LOCATION_REFRESH_CONSUMED_STATE_KEY = "__v30_location_refresh_consumed__"


def _query_param_text(name: str) -> str:
    try:
        value = st.query_params.get(name)
    except Exception:
        try:
            values = st.experimental_get_query_params().get(name, [])
            value = values[0] if values else ""
        except Exception:
            value = ""
    if isinstance(value, (list, tuple)):
        value = value[0] if value else ""
    return str(value or "").strip()


def _wrap_dispatch_geolocation_component(component):
    """Feed the floating-refresh GPS token into the background GPS component."""

    def dispatch_geolocation_component(*args, **kwargs):
        key = str(kwargs.get("key") or "")
        refresh_token = _query_param_text(_LOCATION_REFRESH_QUERY_KEY)
        consumed_token = str(
            st.session_state.get(_LOCATION_REFRESH_CONSUMED_STATE_KEY) or ""
        )
        if (
            refresh_token
            and refresh_token != consumed_token
            and key.startswith("dispatch_geolocation_background::")
        ):
            kwargs["request_token"] = refresh_token
            st.session_state[_LOCATION_REFRESH_CONSUMED_STATE_KEY] = refresh_token
        return component(*args, **kwargs)

    return dispatch_geolocation_component


def _module_safe_declare_component(
    name: str,
    path: str | None = None,
    url: str | None = None,
) -> Any:
    """Call Streamlit's original declaration from a normal imported module."""
    component = _ORIGINAL_DECLARE_COMPONENT(name=name, path=path, url=url)
    if name == "dispatch_geolocation_v3":
        return _wrap_dispatch_geolocation_component(component)
    return component


def _combined_refresh_component_html(body, *args, **kwargs):
    """Apply maintained browser-side compatibility patches to legacy HTML."""
    if isinstance(body, str):
        if 'refreshUrl.searchParams.set("live_refresh", String(Date.now()));' in body:
            body = body.replace(
                'refreshUrl.searchParams.set("live_refresh", String(Date.now()));',
                'const combinedRefreshToken = String(Date.now());\n'
                '                        refreshUrl.searchParams.set("live_refresh", combinedRefreshToken);\n'
                '                        refreshUrl.searchParams.set("location_refresh", combinedRefreshToken);',
                1,
            )
            body = body.replace(
                "正在重新同步 YouBike 即時資料…",
                "正在更新場站資料與定位…",
            )

        if "const ROOT='ubike-battery-v29-upgrade';" in body:
            body = body.replace(
                " const ROOT='ubike-battery-v29-upgrade';",
                " const ROOT='ubike-battery-v29-upgrade';\n"
                " const FLOAT_HIDE_STYLE_ID='ubike-battery-float-hide-style';\n"
                " function setBatteryModalFloatingHidden(hidden){\n"
                "   let style=doc.getElementById(FLOAT_HIDE_STYLE_ID);\n"
                "   if(!style){style=doc.createElement('style');style.id=FLOAT_HIDE_STYLE_ID;style.textContent='html.ubike-battery-modal-open #ubike-float-tools,html.ubike-battery-modal-open #jarvis-voice-indicator{display:none!important;}';doc.head.appendChild(style);}\n"
                "   doc.documentElement.classList.toggle('ubike-battery-modal-open',Boolean(hidden));\n"
                " }",
                1,
            )
            body = body.replace(
                "function open(){const root=ensure(),page=root.querySelector('#ub-v29-page');root.querySelector('#ub-v29-fab').style.display='none';",
                "function open(){const root=ensure(),page=root.querySelector('#ub-v29-page');setBatteryModalFloatingHidden(true);root.querySelector('#ub-v29-fab').style.display='none';",
                1,
            )
            body = body.replace(
                "doc.body.style.overflow=root._bodyOverflow||'';reverseStations.clear();}",
                "doc.body.style.overflow=root._bodyOverflow||'';setBatteryModalFloatingHidden(false);reverseStations.clear();}",
                1,
            )

    return _ORIGINAL_COMPONENT_HTML(body, *args, **kwargs)


def install_component_declare_compat() -> None:
    """Install once for all legacy v1 component declarations and refresh HTML."""
    global _COMPONENT_COMPAT_INSTALLED
    if _COMPONENT_COMPAT_INSTALLED:
        return
    components.declare_component = _module_safe_declare_component
    components.html = _combined_refresh_component_html
    _COMPONENT_COMPAT_INSTALLED = True


# ---------------------------------------------------------------------------
# Version / changelog UI compatibility
# ---------------------------------------------------------------------------
_ORIGINAL_SET_PAGE_CONFIG = st.set_page_config
_ORIGINAL_CAPTION = st.caption
_ORIGINAL_MARKDOWN = st.markdown
_ORIGINAL_POPOVER = getattr(st, "popover", None)


def _set_caller_version_globals() -> None:
    """Promote the exec() legacy globals to the current maintained version."""
    frame = inspect.currentframe()
    try:
        caller = frame.f_back.f_back if frame and frame.f_back else None
        if caller is None:
            return
        namespace = caller.f_globals
        namespace["APP_VERSION"] = SYSTEM_VERSION
        namespace["APP_VERSION_NAME"] = SYSTEM_VERSION
        namespace["APP_BUILD_DATE"] = SYSTEM_BUILD_DATE
    finally:
        del frame


def _versioned_set_page_config(*args, **kwargs):
    _set_caller_version_globals()
    page_title = kwargs.get("page_title")
    if isinstance(page_title, str) and "YouBike 智慧調度" in page_title:
        kwargs["page_title"] = f"臺東 YouBike 智慧調度｜{SYSTEM_VERSION}"
    return _ORIGINAL_SET_PAGE_CONFIG(*args, **kwargs)


def _versioned_caption(body, *args, **kwargs):
    if isinstance(body, str) and body.strip().startswith("系統版本："):
        body = (
            f"系統版本：{SYSTEM_VERSION}｜第{SYSTEM_GENERATION}代｜"
            f"調整 {VERSION_PENDING_CHANGE_COUNT}/{VERSION_CHANGE_THRESHOLD}"
        )
    return _ORIGINAL_CAPTION(body, *args, **kwargs)


def _replace_legacy_version_badge(body: str) -> str:
    marker = '<div id="jarvis-secret-trigger" class="dispatch-version-badge" title="">測試版</div>'
    if marker in body:
        body = body.replace(
            marker,
            f'<div id="jarvis-secret-trigger" class="dispatch-version-badge" title="">{SYSTEM_VERSION} · 第{SYSTEM_GENERATION}代</div>',
            1,
        )
    return body


def _versioned_markdown(body, *args, **kwargs):
    if isinstance(body, str):
        body = _replace_legacy_version_badge(body)
        stripped = body.lstrip()
        if stripped.startswith("#### V29 更新內容") or stripped.startswith("### V29 更新內容"):
            body = SYSTEM_CHANGELOG_MD
    return _ORIGINAL_MARKDOWN(body, *args, **kwargs)


def _versioned_popover(label, *args, **kwargs):
    """Move the legacy update-content popover into the left sidebar."""
    if str(label or "").strip() == "更新內容":
        sidebar_popover = getattr(st.sidebar, "popover", None)
        if callable(sidebar_popover):
            return sidebar_popover("📋 更新內容", *args, **kwargs)
        return st.sidebar.expander("📋 更新內容", expanded=False)
    if callable(_ORIGINAL_POPOVER):
        return _ORIGINAL_POPOVER(label, *args, **kwargs)
    return st.expander(str(label), expanded=False)


def install_version_ui_compat() -> None:
    """Install current version, badge, changelog and sidebar-location wrappers."""
    global _VERSION_UI_INSTALLED
    if _VERSION_UI_INSTALLED:
        return
    st.set_page_config = _versioned_set_page_config
    st.caption = _versioned_caption
    st.markdown = _versioned_markdown
    if callable(_ORIGINAL_POPOVER):
        st.popover = _versioned_popover
    _VERSION_UI_INSTALLED = True


install_version_ui_compat()
