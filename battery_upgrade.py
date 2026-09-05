from __future__ import annotations

import hashlib
import json
from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st
import streamlit.components.v1 as components

from battery_ui_adapter import query_stations_for_ui


TAIPEI = ZoneInfo("Asia/Taipei")


def _fingerprint(value: object) -> str:
    raw = json.dumps(value, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
    return hashlib.sha1(raw).hexdigest()[:16]


def _query_param(name: str, default: str = "") -> str:
    try:
        value = st.query_params.get(name, default)
    except Exception:
        return default
    if isinstance(value, (list, tuple)):
        value = value[-1] if value else default
    return str(value or default)


def _parse_zones(raw: str, valid_zones: set[str]) -> list[str]:
    try:
        value = json.loads(raw) if raw else []
    except (TypeError, ValueError, json.JSONDecodeError):
        value = []
    if not isinstance(value, list):
        return []
    return [str(zone) for zone in value if str(zone) in valid_zones]


def _clean_route_map(route_station_map: dict[str, list[dict]]) -> dict[str, list[dict]]:
    output: dict[str, list[dict]] = {}
    for zone, items in (route_station_map or {}).items():
        zone_name = str(zone).strip()
        if not zone_name or not isinstance(items, list):
            continue
        seen: set[str] = set()
        cleaned: list[dict] = []
        for item in items:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name") or item.get("station_name") or "").strip()
            district = str(item.get("district") or "").strip()
            if not name or name in seen:
                continue
            seen.add(name)
            cleaned.append({"name": name, "district": district})
        if cleaned:
            output[zone_name] = cleaned
    return output


def render_floating_server_battery(
    route_station_map: dict[str, list[dict]],
    mobile_mode: bool,
    *,
    threshold: int = 89,
    priority_threshold: int = 40,
) -> None:
    """V29 Server 電池查詢頁。

    UI 使用單向 components.html 注入，不依賴 Streamlit custom-component ready
    handshake；查詢按鈕透過 URL query params 觸發下一次 Streamlit run。
    """
    clean_map = _clean_route_map(route_station_map)
    fingerprint_data = [
        (zone, [(item["name"], item["district"]) for item in items])
        for zone, items in clean_map.items()
    ]
    fp = _fingerprint(fingerprint_data)
    state_key = f"v29_battery_upgrade::{fp}"
    state = st.session_state.get(state_key)
    if not isinstance(state, dict):
        state = {
            "results": {},
            "selected_zones": [],
            "threshold": int(threshold),
            "priority_threshold": int(priority_threshold),
            "priority_enabled": True,
            "last_message": "",
            "last_updated": "",
            "last_nonce": "",
            "failed_count": 0,
        }

    valid_zones = set(clean_map)
    nonce = _query_param("battery_nonce")
    if nonce and nonce != str(state.get("last_nonce") or ""):
        selected_zones = _parse_zones(_query_param("battery_zones"), valid_zones)
        try:
            th = max(0, min(100, int(float(_query_param("battery_th", str(threshold))))))
        except (TypeError, ValueError):
            th = int(threshold)
        try:
            pr = max(0, min(th, int(float(_query_param("battery_pr", str(priority_threshold))))))
        except (TypeError, ValueError):
            pr = min(th, int(priority_threshold))
        priority_enabled = _query_param("battery_priority", "1") != "0"
        force_refresh = _query_param("battery_force", "0") == "1"

        stations: list[dict] = []
        seen: set[str] = set()
        for zone in selected_zones:
            for item in clean_map.get(zone, []):
                if item["name"] in seen:
                    continue
                seen.add(item["name"])
                stations.append(item)

        state["last_nonce"] = nonce
        state["selected_zones"] = selected_zones
        state["threshold"] = th
        state["priority_threshold"] = pr
        state["priority_enabled"] = priority_enabled

        if not selected_zones:
            state["results"] = {}
            state["failed_count"] = 0
            state["last_message"] = "請至少選擇一個範圍"
        elif not stations:
            state["results"] = {}
            state["failed_count"] = 0
            state["last_message"] = "所選範圍沒有可查詢場站"
        else:
            force_names = {item["name"] for item in stations} if force_refresh else set()
            with st.spinner("正在由 V29 Server 查詢 2.0E 電池……"):
                results = query_stations_for_ui(
                    stations,
                    threshold=th,
                    priority_threshold=pr,
                    force_names=force_names,
                )
            state["results"] = results
            failed_count = sum(1 for item in results.values() if item.get("error"))
            ok_count = len(results) - failed_count
            low_station_count = sum(
                1 for item in results.values()
                if not item.get("error") and int(item.get("low_count") or 0) > 0
            )
            low_bike_count = sum(
                int(item.get("low_count") or 0)
                for item in results.values()
                if not item.get("error")
            )
            state["failed_count"] = failed_count
            state["last_updated"] = datetime.now(TAIPEI).strftime("%Y/%m/%d %H:%M:%S")
            state["last_message"] = (
                f"完成 {ok_count} 站｜{low_station_count} 站有低電｜低電 {low_bike_count} 台"
                + (f"｜{failed_count} 站未取得" if failed_count else "")
                + ("｜強制更新" if force_refresh else "｜共用快取")
            )
        st.session_state[state_key] = state

    # 清掉不存在於新 Excel 的舊區域，但保留使用者曾選擇的範圍。
    state["selected_zones"] = [
        zone for zone in state.get("selected_zones", []) if zone in valid_zones
    ]
    st.session_state[state_key] = state

    args = {
        "route_station_map": clean_map,
        "results": state.get("results", {}),
        "selected_zones": state.get("selected_zones", []),
        "threshold": int(state.get("threshold", threshold)),
        "priority_threshold": int(state.get("priority_threshold", priority_threshold)),
        "priority_enabled": bool(state.get("priority_enabled", True)),
        "last_message": str(state.get("last_message") or ""),
        "last_updated": str(state.get("last_updated") or ""),
        "failed_count": int(state.get("failed_count") or 0),
        "open_page": _query_param("battery_open", "0") == "1",
        "mobile": bool(mobile_mode),
    }
    payload = json.dumps(args, ensure_ascii=False, default=str).replace("</", "<\\/")

    html_text = r'''<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">
<style>html,body{width:1px;height:1px;margin:0;padding:0;overflow:hidden;background:transparent}</style></head><body>
<script>
(()=>{
 const args=__ARGS__;
 const win=window.parent, doc=win.document;
 const ROOT='ubike-battery-v29-upgrade';
 function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
 function number(v,fallback=0){const n=Number(v);return Number.isFinite(n)?n:fallback;}
 function setUrlQuery(force){
   const root=doc.getElementById(ROOT); if(!root)return;
   const selected=[...root.querySelectorAll('[data-zone]:checked')].map(x=>x.value);
   const status=root.querySelector('#ub-v29-status');
   if(!selected.length){status.textContent='請至少選擇一個範圍';status.classList.add('warn');return;}
   const th=Math.max(0,Math.min(100,number(root.querySelector('#ub-v29-th').value,89)));
   const pr=Math.max(0,Math.min(th,number(root.querySelector('#ub-v29-pr').value,40)));
   const priority=root.querySelector('#ub-v29-pr-enabled').checked;
   try{localStorage.setItem('ubike-v29-battery-pref2',JSON.stringify({zones:selected,threshold:th,priority_threshold:pr,priority_enabled:priority}));}catch(_){}
   const url=new URL(win.location.href);
   url.searchParams.set('battery_open','1');
   url.searchParams.set('battery_zones',JSON.stringify(selected));
   url.searchParams.set('battery_th',String(th));
   url.searchParams.set('battery_pr',String(pr));
   url.searchParams.set('battery_priority',priority?'1':'0');
   url.searchParams.set('battery_force',force?'1':'0');
   url.searchParams.set('battery_nonce',String(Date.now()));
   status.textContent=force?'正在強制重新抓取官方電池資料…':'正在查詢，優先使用 V29 Server 共用快取…';
   status.classList.remove('warn');
   root.querySelectorAll('button').forEach(b=>b.disabled=true);
   win.location.href=url.toString();
 }
 function clearBatteryParams(){
   try{const url=new URL(win.location.href);['battery_open','battery_zones','battery_th','battery_pr','battery_priority','battery_force','battery_nonce'].forEach(k=>url.searchParams.delete(k));win.history.replaceState(win.history.state,'',url.toString());}catch(_){}
 }
 function minBattery(r){const lows=Array.isArray(r.low_bikes)?r.low_bikes:[];if(!lows.length)return 101;return Math.min(...lows.map(b=>number(b.battery_power,101)));}
 function resultRows(results,priorityEnabled){
   const rows=Object.values(results||{}).filter(r=>r&&!r.error&&number(r.low_count)>0);
   rows.sort((a,b)=>minBattery(a)-minBattery(b)||number(b.low_count)-number(a.low_count)||String(a.requested_name||'').localeCompare(String(b.requested_name||''),'zh-Hant'));
   return rows.map(r=>{
     const min=minBattery(r), urgent=priorityEnabled?number(r.priority_count):0;
     const bikes=(r.low_bikes||[]).slice().sort((a,b)=>number(a.battery_power,101)-number(b.battery_power,101)||String(a.pillar_no||'').localeCompare(String(b.pillar_no||''))).map(b=>{
       const power=number(b.battery_power,101);const isUrgent=priorityEnabled&&power<=number(r.priority_threshold,40);
       return `<div class="bike ${isUrgent?'urgent':''}"><span>柱 ${esc(b.pillar_no||'—')}</span><span>${esc(b.bike_no||'')}</span><strong>${esc(power)}%</strong></div>`;
     }).join('');
     const stale=r.source==='stale_cache'?`<div class="stale">⚠ 官方更新失敗，顯示約 ${Math.round(number(r.age_seconds))} 秒前快取</div>`:'';
     return `<details class="station" data-station="${esc(String(r.requested_name||r.station_name||'').toLowerCase())}"><summary><div><strong>${esc(r.requested_name||r.station_name||'')}</strong><small>低電 ${number(r.low_count)} 台${urgent?`｜緊急 ${urgent} 台`:''}</small></div><div class="min ${min<=number(r.priority_threshold,40)&&priorityEnabled?'hot':''}">${min}%</div></summary><div class="bike-list">${bikes}${stale}</div></details>`;
   }).join('');
 }
 function ensure(){
   let root=doc.getElementById(ROOT);
   if(root)return root;
   root=doc.createElement('div');root.id=ROOT;
   root.innerHTML=`<button id="ub-v29-fab" aria-label="電量查詢">⚡</button><section id="ub-v29-page" aria-hidden="true"><div class="shell"><header><button id="ub-v29-close">‹ 返回</button><div><h1>⚡ 電量查詢</h1><p>V29 Server 引擎｜共用快取｜手機穩定版</p></div></header><main id="ub-v29-main"></main></div></section>`;
   const style=doc.createElement('style');style.id=ROOT+'-style';style.textContent=`
#${ROOT}{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft JhengHei",sans-serif;color:#eaf7ff}#ub-v29-fab{position:fixed;right:18px;bottom:110px;z-index:2147482600;width:54px;height:54px;border:1px solid rgba(99,235,255,.55);border-radius:50%;background:linear-gradient(145deg,#081524,#101f35);color:#6cf1ff;font-size:24px;box-shadow:0 0 0 1px rgba(255,67,159,.18),0 10px 30px rgba(0,0,0,.4),0 0 22px rgba(44,215,255,.18);touch-action:manipulation}#ub-v29-page{display:none;position:fixed;inset:0;z-index:2147482601;background:radial-gradient(circle at 80% 0%,rgba(255,54,147,.16),transparent 32%),radial-gradient(circle at 0% 20%,rgba(28,224,255,.13),transparent 30%),#07101c;overflow:auto;overscroll-behavior:contain}#ub-v29-page.open{display:block}.shell{max-width:900px;margin:0 auto;padding:max(10px,env(safe-area-inset-top)) 12px max(30px,env(safe-area-inset-bottom))}header{position:sticky;top:0;z-index:5;display:flex;justify-content:space-between;align-items:center;gap:14px;padding:9px 0 12px;background:linear-gradient(#07101c 75%,transparent)}header h1{font-size:24px;margin:0;color:#f4fbff}header p{margin:3px 0 0;font-size:12px;color:#89a9bd}#ub-v29-close{min-width:82px;min-height:44px;border:1px solid #27465d;border-radius:999px;background:#0d1c2b;color:#e9fbff;font-size:16px;font-weight:800}.panel{background:rgba(12,28,43,.92);border:1px solid rgba(94,205,229,.23);border-radius:17px;padding:14px;margin:8px 0;box-shadow:0 12px 32px rgba(0,0,0,.18)}.title-row,.actions,.summary{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.title-row{justify-content:space-between}.title-row strong{font-size:15px}.mini{border:1px solid #294b63;background:#10283a;color:#bdeffc;border-radius:999px;padding:6px 10px;font-weight:750}.zones{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}.zone{border:1px solid #2a4e66;border-radius:999px;padding:7px 10px;background:#0a1c2a;color:#d8f6ff}.zone input{accent-color:#4feaff}.controls{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:12px}.field{display:flex;flex-direction:column;gap:5px;color:#9db8c8;font-size:12px}.field input[type=number],#ub-v29-search{box-sizing:border-box;width:100%;min-height:42px;border:1px solid #2c526b;border-radius:11px;background:#081825;color:#eefcff;padding:8px 10px;font-size:16px}.toggle{display:flex;align-items:center;gap:8px;margin-top:10px;color:#bed3df}.actions{margin-top:12px}.actions button{flex:1;min-width:125px;min-height:44px;border:0;border-radius:12px;font-size:15px;font-weight:900}.query{background:linear-gradient(90deg,#43e6ff,#74f3c5);color:#041018}.force{background:#38152c;color:#ffb8d6;border:1px solid #7e2d59!important}.status{padding:10px 2px;color:#9dc1d1;font-size:13px}.status.warn{color:#ffb46a}.updated{font-size:12px;color:#789aaa}.summary{margin:8px 0}.chip{border:1px solid #2b5269;border-radius:999px;padding:6px 9px;background:#0b2030;color:#bfefff;font-size:12px}.chip.hot{border-color:#8b2d5b;color:#ff9bc7}.station{background:rgba(11,26,40,.96);border:1px solid #24495e;border-radius:15px;margin:9px 0;overflow:hidden}.station summary{list-style:none;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:13px 14px;cursor:pointer}.station summary::-webkit-details-marker{display:none}.station summary div:first-child{display:flex;flex-direction:column;gap:4px}.station summary strong{color:#f2fbff}.station summary small{color:#8db0c1}.min{min-width:54px;text-align:center;border-radius:999px;padding:7px 9px;background:#103447;color:#6cecff;font-weight:900}.min.hot{background:#4d1734;color:#ff8cbd}.bike-list{border-top:1px solid #1f3b4d;padding:8px 13px 12px}.bike{display:grid;grid-template-columns:74px 1fr 52px;gap:8px;align-items:center;padding:7px 3px;border-bottom:1px dashed rgba(110,160,180,.18);color:#cce7f2}.bike strong{text-align:right;color:#66edff}.bike.urgent,.bike.urgent strong{color:#ff83b7;font-weight:850}.stale{padding:8px 0 0;color:#ffb968;font-size:12px}.empty{padding:30px 12px;text-align:center;color:#7fa4b6}.search-wrap{margin-top:10px}@media(max-width:560px){.controls{grid-template-columns:1fr}.shell{padding-left:10px;padding-right:10px}header h1{font-size:21px}.bike{grid-template-columns:65px 1fr 48px}}
`;
   doc.head.appendChild(style);doc.body.appendChild(root);
   root.querySelector('#ub-v29-fab').onclick=()=>open();root.querySelector('#ub-v29-close').onclick=()=>close();
   return root;
 }
 function open(){const root=ensure(),page=root.querySelector('#ub-v29-page');root.querySelector('#ub-v29-fab').style.display='none';page.classList.add('open');page.setAttribute('aria-hidden','false');root._htmlOverflow=doc.documentElement.style.overflow;root._bodyOverflow=doc.body.style.overflow;doc.documentElement.style.overflow='hidden';doc.body.style.overflow='hidden';render();}
 function close(){const root=ensure(),page=root.querySelector('#ub-v29-page');page.classList.remove('open');page.setAttribute('aria-hidden','true');root.querySelector('#ub-v29-fab').style.display='';doc.documentElement.style.overflow=root._htmlOverflow||'';doc.body.style.overflow=root._bodyOverflow||'';clearBatteryParams();}
 function render(){
   const root=ensure(),main=root.querySelector('#ub-v29-main'),map=args.route_station_map||{},zones=Object.keys(map),results=args.results||{};
   let pref={};try{pref=JSON.parse(localStorage.getItem('ubike-v29-battery-pref2')||'{}')||{};}catch(_){}
   const selected=(Array.isArray(args.selected_zones)&&args.selected_zones.length?args.selected_zones:(Array.isArray(pref.zones)?pref.zones:[])).filter(z=>zones.includes(z));
   const th=number(args.threshold,89),pr=number(args.priority_threshold,40),priorityEnabled=args.priority_enabled!==false;
   const rows=Object.values(results).filter(r=>r&&!r.error&&number(r.low_count)>0);const lowBikes=rows.reduce((s,r)=>s+number(r.low_count),0);const urgent=priorityEnabled?rows.reduce((s,r)=>s+number(r.priority_count),0):0;
   main.innerHTML=`<div class="panel"><div class="title-row"><strong>選擇範圍</strong><div><button class="mini" id="ub-v29-all">全選</button> <button class="mini" id="ub-v29-none">取消</button></div></div><div class="zones">${zones.map(z=>`<label class="zone"><input data-zone type="checkbox" value="${esc(z)}" ${selected.includes(z)?'checked':''}> ${esc(z)} <span style="opacity:.55">(${(map[z]||[]).length})</span></label>`).join('')}</div><div class="controls"><label class="field">低電門檻<input id="ub-v29-th" type="number" min="0" max="100" value="${th}"></label><label class="field">緊急門檻<input id="ub-v29-pr" type="number" min="0" max="100" value="${pr}" ${priorityEnabled?'':'disabled'}></label></div><label class="toggle"><input id="ub-v29-pr-enabled" type="checkbox" ${priorityEnabled?'checked':''}> 啟用第二門檻／紅色緊急提示</label><div class="actions"><button class="query" id="ub-v29-query">查詢</button><button class="force" id="ub-v29-force">強制更新官方資料</button></div></div><div class="status" id="ub-v29-status">${esc(args.last_message||'選好範圍後按「查詢」')}</div><div class="updated">${args.last_updated?`最後更新：${esc(args.last_updated)}`:'尚未查詢'}</div><div class="summary"><span class="chip">低電場站 ${rows.length}</span><span class="chip">低電車 ${lowBikes}</span>${priorityEnabled?`<span class="chip hot">緊急 ${urgent}</span>`:''}${number(args.failed_count)>0?`<span class="chip hot">未取得 ${number(args.failed_count)} 站</span>`:''}</div><div class="search-wrap"><input id="ub-v29-search" type="search" placeholder="搜尋場站名稱"></div><div id="ub-v29-results">${rows.length?resultRows(results,priorityEnabled):'<div class="empty">目前沒有低於門檻的 2.0E 場站</div>'}</div>`;
   main.querySelector('#ub-v29-all').onclick=()=>main.querySelectorAll('[data-zone]').forEach(x=>x.checked=true);main.querySelector('#ub-v29-none').onclick=()=>main.querySelectorAll('[data-zone]').forEach(x=>x.checked=false);main.querySelector('#ub-v29-pr-enabled').onchange=e=>{main.querySelector('#ub-v29-pr').disabled=!e.target.checked;};main.querySelector('#ub-v29-query').onclick=()=>setUrlQuery(false);main.querySelector('#ub-v29-force').onclick=()=>setUrlQuery(true);main.querySelector('#ub-v29-search').oninput=e=>{const q=String(e.target.value||'').trim().toLowerCase();main.querySelectorAll('.station').forEach(card=>{card.style.display=!q||String(card.dataset.station||'').includes(q)?'':'none';});};
 }
 ensure();render();if(args.open_page)open();
})();
</script></body></html>'''.replace('__ARGS__', payload)

    components.html(html_text, height=0, scrolling=False)
