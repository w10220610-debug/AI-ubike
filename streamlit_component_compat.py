from __future__ import annotations

"""Compatibility helpers shared by the V30 entrypoint and legacy UI.

This module still provides the original Streamlit v1 component declaration shim,
but it is also the single source of truth for the visible/system version badge
and generation changelog.  The legacy UI is executed through ``exec(compile())``,
so these lightweight wrappers let the maintained entrypoint advance versions
without rewriting the large legacy UI file on every release.
"""

import inspect
from typing import Any

import streamlit as st
import streamlit.components.v1 as components


# ---------------------------------------------------------------------------
# System version / generation metadata
# ---------------------------------------------------------------------------
SYSTEM_VERSION = "V30"
SYSTEM_GENERATION = 30
SYSTEM_BUILD_DATE = "2026-09-07"
SYSTEM_VERSION_LABEL = f"{SYSTEM_VERSION}｜第{SYSTEM_GENERATION}代"

SYSTEM_CHANGELOG_MD = """
### 第30代｜V30
**更新日期：2026-09-07**

- AI 預測正式改為 **60 分鐘後**的場站車量趨勢／估計值，並顯示信心水準與樣本數。
- AI 學習紀錄上限提升為 **50,000 筆**。
- 手機、平板、電腦共用同一個 AI 學習池，學習資料預設在背景自動同步。
- 新增共用調度員定位池：兩名使用者的有效 GPS 都可參與人工調度判定。
- 場站車量發生變化時，若調度員位於該站附近，AI 會自動標記為人工調度，避免污染自然需求學習。
- 保留人工調度手動標記，並維持自然／人工／疑似人工三類防污染機制。
- 系統版本欄位改由中央版本資料管理；之後每次大版本更新可持續往 V31、V32…提升。
- 「更新內容」入口移至左側選單，並以「第幾代｜版本」方式保留歷代更新紀錄。

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
_COMPONENT_COMPAT_INSTALLED = False


def _module_safe_declare_component(
    name: str,
    path: str | None = None,
    url: str | None = None,
) -> Any:
    """Call Streamlit's original declaration from a normal imported module."""
    return _ORIGINAL_DECLARE_COMPONENT(name=name, path=path, url=url)


def install_component_declare_compat() -> None:
    """Install once for all legacy v1 custom-component declarations."""
    global _COMPONENT_COMPAT_INSTALLED
    if _COMPONENT_COMPAT_INSTALLED:
        return
    components.declare_component = _module_safe_declare_component
    _COMPONENT_COMPAT_INSTALLED = True


# ---------------------------------------------------------------------------
# V30 version / changelog UI compatibility
# ---------------------------------------------------------------------------
_ORIGINAL_SET_PAGE_CONFIG = st.set_page_config
_ORIGINAL_CAPTION = st.caption
_ORIGINAL_MARKDOWN = st.markdown
_ORIGINAL_POPOVER = getattr(st, "popover", None)
_VERSION_UI_INSTALLED = False


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
        body = f"系統版本：{SYSTEM_VERSION}｜第{SYSTEM_GENERATION}代"
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
    """Install V30 version, badge, changelog and sidebar-location wrappers once."""
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
