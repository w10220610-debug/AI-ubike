"""Nonblocking live-status bridge for the maintained legacy UI."""
from __future__ import annotations

import hashlib
import json

import streamlit as st

from background_refresh import live_refresh
import live_status_service


def station_scope(stations) -> tuple[tuple[str, str], ...]:
    return tuple(sorted({(str(s.get("name") or s.get("station_name") or "").strip(),
                          str(s.get("district") or "").strip()) for s in stations
                         if isinstance(s, dict) and (s.get("name") or s.get("station_name"))}))


def _load(scope, *, force=False):
    result = live_status_service.get_live_status_for_stations([{"name": name, "district": district}
                                          for name, district in scope], force=force)
    if not result.get("ok") or result.get("error"):
        raise RuntimeError(result.get("error") or "即時車數查詢失敗")
    return result


def _emit_sync_state(state, *, station_count=0, message=""):
    payload = json.dumps({"source": "ubike-browser-sync", "type": "ubike:sync-state",
                          "state": state, "station_count": station_count, "message": message},
                         ensure_ascii=False).replace("</", "<\\/")
    st.iframe(f"<script>window.parent.postMessage({payload}, '*');</script>",
              height=1, tab_index=-1)


@st.fragment(run_every=2)
def _monitor(scope, state_key):
    state = live_refresh.poll(scope, lambda: _load(scope))
    awaiting_key = state_key + "::manual"
    if state.pending:
        st.caption("即時車數在背景更新中；目前仍可操作頁面。")
    if state.error:
        st.warning(f"即時車數更新未完成：{state.error} 原有車數未清除。")
    if state.value is None and not state.error:
        st.caption("尚未取得即時車數，請先確認資料更新時間。")

    awaiting = st.session_state.get(awaiting_key)
    if awaiting is not None and not state.pending and state.revision > awaiting:
        _emit_sync_state("error" if state.error else "success",
                         station_count=len((state.value or {}).get("records", [])),
                         message=state.error)
        st.session_state.pop(awaiting_key, None)

    event = str((state.value or {}).get("event_id") or "")
    if event and event != st.session_state.get(state_key):
        st.session_state[state_key] = event
        # Only a newly completed snapshot reruns data rendering; polling stays local.
        st.rerun()


def render_background_live_status(stations, *, force=False):
    scope = station_scope(stations)
    if not scope:
        st.warning("目前配置沒有可供同步的場站；原有資料保留。")
        return None
    state_key = "background_live::" + hashlib.sha256(repr(scope).encode()).hexdigest()
    state = live_refresh.poll(scope, lambda: _load(scope, force=force), force=force)
    if force:
        st.session_state[state_key + "::manual"] = state.revision
        _emit_sync_state("busy")
    st.session_state[state_key] = str((state.value or {}).get("event_id") or "")
    _monitor(scope, state_key)
    return state.value
