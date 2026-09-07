from __future__ import annotations

"""Durable remote backing store for the shared AI learning pool.

The existing ``ai_learning_guard`` local JSON pool remains the fast cache and
fallback. When Supabase credentials are present in Streamlit Secrets or the
environment, this module mirrors records to PostgREST and restores them after a
Streamlit container/redeploy loses ``.base_cache``.

Supported secrets/environment variables:
- SUPABASE_URL
- SUPABASE_SERVICE_ROLE_KEY (recommended) or SUPABASE_ANON_KEY
- SUPABASE_AI_TABLE (optional, defaults to ``ai_learning_records``)

The same values may be placed under ``[supabase]`` as ``url``,
``service_role_key``/``anon_key`` and ``ai_table``.
"""

import json
import math
import os
import threading
import time
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

import streamlit as st


REMOTE_PULL_SECONDS = 60.0
REMOTE_PAGE_SIZE = 1000
REMOTE_PUSH_BATCH_SIZE = 200
REMOTE_TIMEOUT_SECONDS = 12.0
DEFAULT_TABLE = "ai_learning_records"

_REMOTE_LOCK = threading.RLock()
_REMOTE_CACHE: list[dict] = []
_REMOTE_CACHE_IDS: set[str] = set()
_REMOTE_LAST_PULL_MONOTONIC = 0.0
_REMOTE_INITIALIZED = False
_INSTALL_LOCK = threading.RLock()


def _secret_text(name: str, nested_name: str) -> str:
    value = str(os.getenv(name, "") or "").strip()
    if value:
        return value
    try:
        direct = st.secrets.get(name, "")
        value = str(direct or "").strip()
        if value:
            return value
    except Exception:
        pass
    try:
        section = st.secrets.get("supabase", {})
        if hasattr(section, "get"):
            value = str(section.get(nested_name, "") or "").strip()
            if value:
                return value
    except Exception:
        pass
    return ""


def _config() -> tuple[str, str, str] | None:
    base_url = _secret_text("SUPABASE_URL", "url").rstrip("/")
    service_key = _secret_text("SUPABASE_SERVICE_ROLE_KEY", "service_role_key")
    anon_key = _secret_text("SUPABASE_ANON_KEY", "anon_key")
    api_key = service_key or anon_key
    table = _secret_text("SUPABASE_AI_TABLE", "ai_table") or DEFAULT_TABLE
    if not base_url or not api_key:
        return None
    return base_url, api_key, table


def _headers(api_key: str, *, write: bool = False) -> dict[str, str]:
    headers = {
        "apikey": api_key,
        "Authorization": f"Bearer {api_key}",
        "Accept": "application/json",
    }
    if write:
        headers["Content-Type"] = "application/json"
        headers["Prefer"] = "resolution=merge-duplicates,return=minimal"
    return headers


def _request_json(
    url: str,
    *,
    api_key: str,
    method: str = "GET",
    payload: object | None = None,
    write: bool = False,
) -> object:
    body = None
    if payload is not None:
        body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), allow_nan=False).encode("utf-8")
    request = Request(
        url,
        data=body,
        headers=_headers(api_key, write=write),
        method=method,
    )
    try:
        with urlopen(request, timeout=REMOTE_TIMEOUT_SECONDS) as response:
            raw = response.read()
            if not raw:
                return None
            return json.loads(raw.decode("utf-8"))
    except HTTPError as exc:
        try:
            detail = exc.read().decode("utf-8", errors="replace")[:500]
        except Exception:
            detail = ""
        raise RuntimeError(f"Supabase HTTP {exc.code}: {detail or exc.reason}") from exc
    except URLError as exc:
        raise RuntimeError(f"Supabase 連線失敗：{exc.reason}") from exc


def _finite_epoch(value: object) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError, OverflowError):
        return 0.0
    return number if math.isfinite(number) else 0.0


def _safe_json_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, dict):
        return {str(key): _safe_json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_safe_json_value(item) for item in value]
    return str(value)


def _record_identity(module, record: dict) -> str:
    return str(module._record_identity(record))


def _set_status(
    *,
    mode: str,
    local_count: int,
    remote_count: int | None = None,
    error: str = "",
) -> None:
    try:
        st.session_state["__ai_persistent_pool_status__"] = {
            "mode": mode,
            "local_count": int(local_count),
            "remote_count": None if remote_count is None else int(remote_count),
            "error": str(error or ""),
            "updated_at_epoch": time.time(),
        }
    except Exception:
        pass


def _remote_pull(module, cfg: tuple[str, str, str]) -> list[dict]:
    base_url, api_key, table = cfg
    endpoint = f"{base_url}/rest/v1/{quote(table, safe='')}"
    output: list[dict] = []
    offset = 0
    max_records = int(getattr(module, "MAX_TRANSITION_RECORDS", 50000))

    while len(output) < max_records:
        limit = min(REMOTE_PAGE_SIZE, max_records - len(output))
        query = urlencode(
            {
                "select": "record_id,observed_at_epoch,payload",
                "order": "observed_at_epoch.desc",
                "offset": offset,
                "limit": limit,
            }
        )
        raw = _request_json(f"{endpoint}?{query}", api_key=api_key)
        rows = raw if isinstance(raw, list) else []
        if not rows:
            break
        for row in rows:
            if not isinstance(row, dict):
                continue
            payload = row.get("payload")
            if not isinstance(payload, dict):
                continue
            record = dict(payload)
            if not str(record.get("record_id") or "").strip():
                record["record_id"] = str(row.get("record_id") or "").strip()
            if not record.get("observed_at_epoch"):
                record["observed_at_epoch"] = _finite_epoch(row.get("observed_at_epoch"))
            output.append(record)
            if len(output) >= max_records:
                break
        if len(rows) < limit:
            break
        offset += len(rows)

    return module._sort_trim_records(output)


def _remote_upsert(module, cfg: tuple[str, str, str], records: list[dict]) -> None:
    if not records:
        return
    base_url, api_key, table = cfg
    endpoint = f"{base_url}/rest/v1/{quote(table, safe='')}?on_conflict=record_id"
    for start in range(0, len(records), REMOTE_PUSH_BATCH_SIZE):
        batch = records[start : start + REMOTE_PUSH_BATCH_SIZE]
        rows = []
        for record in batch:
            if not isinstance(record, dict):
                continue
            rows.append(
                {
                    "record_id": _record_identity(module, record),
                    "observed_at_epoch": _finite_epoch(record.get("observed_at_epoch")),
                    "payload": _safe_json_value(record),
                }
            )
        if rows:
            _request_json(
                endpoint,
                api_key=api_key,
                method="POST",
                payload=rows,
                write=True,
            )


def install_persistent_learning_pool(module) -> bool:
    """Patch ``ai_learning_guard.sync_shared_learning_pool`` once.

    Returns True when the patch is installed. Remote durability activates only
    after Supabase credentials exist; otherwise the original local shared pool
    continues unchanged.
    """
    global _REMOTE_CACHE, _REMOTE_CACHE_IDS, _REMOTE_LAST_PULL_MONOTONIC, _REMOTE_INITIALIZED

    with _INSTALL_LOCK:
        current = getattr(module, "sync_shared_learning_pool", None)
        if not callable(current):
            return False
        if getattr(current, "__durable_remote_pool__", False):
            return True

        original_sync = current

        def durable_sync(
            incoming_records: list[dict] | None = None,
            *,
            force_read: bool = False,
        ) -> list[dict]:
            global _REMOTE_CACHE, _REMOTE_CACHE_IDS, _REMOTE_LAST_PULL_MONOTONIC, _REMOTE_INITIALIZED

            local = original_sync(incoming_records, force_read=force_read)
            cfg = _config()
            if cfg is None:
                _set_status(mode="local_only", local_count=len(local))
                return local

            with _REMOTE_LOCK:
                now_mono = time.monotonic()
                need_pull = bool(
                    force_read
                    or not _REMOTE_INITIALIZED
                    or now_mono - _REMOTE_LAST_PULL_MONOTONIC >= REMOTE_PULL_SECONDS
                )

                if need_pull:
                    try:
                        pulled = _remote_pull(module, cfg)
                        _REMOTE_CACHE = [dict(item) for item in pulled]
                        _REMOTE_CACHE_IDS = {
                            _record_identity(module, item) for item in _REMOTE_CACHE
                        }
                        _REMOTE_LAST_PULL_MONOTONIC = now_mono
                        _REMOTE_INITIALIZED = True
                    except Exception as exc:
                        _set_status(
                            mode="remote_error",
                            local_count=len(local),
                            remote_count=len(_REMOTE_CACHE) if _REMOTE_INITIALIZED else None,
                            error=str(exc),
                        )
                        return local if not _REMOTE_CACHE else module._sort_trim_records([*_REMOTE_CACHE, *local])

                merged = module._sort_trim_records([*_REMOTE_CACHE, *local])
                local_ids = {_record_identity(module, item) for item in local}
                merged_ids = {_record_identity(module, item) for item in merged}
                if merged_ids != local_ids:
                    try:
                        module._write_shared_pool(merged)
                    except Exception:
                        pass

                missing_remote = [
                    item
                    for item in merged
                    if _record_identity(module, item) not in _REMOTE_CACHE_IDS
                ]
                if missing_remote:
                    try:
                        _remote_upsert(module, cfg, missing_remote)
                        _REMOTE_CACHE = module._sort_trim_records([*_REMOTE_CACHE, *missing_remote])
                        _REMOTE_CACHE_IDS = {
                            _record_identity(module, item) for item in _REMOTE_CACHE
                        }
                    except Exception as exc:
                        _set_status(
                            mode="remote_error",
                            local_count=len(merged),
                            remote_count=len(_REMOTE_CACHE),
                            error=str(exc),
                        )
                        return merged

                try:
                    module._LATEST_LEARNING_RECORDS = [dict(item) for item in merged]
                    st.session_state["__ai_prediction_records__"] = [dict(item) for item in merged]
                    st.session_state["__ai_shared_pool_count__"] = len(merged)
                    st.session_state["__ai_shared_pool_synced_at__"] = time.time()
                except Exception:
                    pass

                _set_status(
                    mode="supabase",
                    local_count=len(merged),
                    remote_count=len(_REMOTE_CACHE),
                )
                return [dict(item) for item in merged]

        durable_sync.__durable_remote_pool__ = True
        durable_sync.__wrapped__ = original_sync
        module.sync_shared_learning_pool = durable_sync
        return True
