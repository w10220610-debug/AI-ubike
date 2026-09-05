from __future__ import annotations

import hashlib
import json
import tempfile
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from battery_ui_adapter import query_station_for_ui, query_stations_for_ui


_COMPONENT_CACHE: dict[str, object] = {}


def _component(name: str, html_text: str):
    cached = _COMPONENT_CACHE.get(name)
    if cached is not None:
        return cached
    component_dir = Path(tempfile.gettempdir()) / name
    component_dir.mkdir(parents=True, exist_ok=True)
    index_path = component_dir / "index.html"
    try:
        if not index_path.exists() or index_path.read_text(encoding="utf-8") != html_text:
            index_path.write_text(html_text, encoding="utf-8")
    except OSError:
        pass
    declared = components.declare_component(name, path=str(component_dir))
    _COMPONENT_CACHE[name] = declared
    return declared


INLINE_COMPONENT_HTML = r'''<!doctype html><html><head><meta charset="utf-8"><style>html,body{margin:0;padding:0;background:transparent}</style></head><body>
<script>
(()=>{
 const API=1;
 function send(type,data={}){window.parent.postMessage({isStreamlitMessage:true,type,...data},"*");}
 function setValue(value){send("streamlit:setComponentValue",{value});}
 function setHeight(){send("streamlit:setFrameHeight",{height:1});}
 function esc(s){return String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
 function clearOld(doc){doc.querySelectorAll('.ubike-server-inline-battery').forEach(n=>n.remove());}
 function render(args){
   const doc=window.parent.document; clearOld(doc);
   const specs=Array.isArray(args.specs)?args.specs:[]; const results=args.results||{};
   for(const spec of specs){
     const host=doc.getElementById(String(spec.target||"")); if(!host) continue;
     const name=String(spec.name||""); const result=results[name];
     const wrap=doc.createElement('div'); wrap.className='ubike-server-inline-battery';
     wrap.style.cssText='margin-top:6px;padding:7px 9px;border-radius:10px;background:rgba(15,23,42,.06);font-size:12px;line-height:1.55;';
     if(result && !result.error){
       const low=Array.isArray(result.low_bikes)?result.low_bikes:[]; const pri=Array.isArray(result.priority_bikes)?result.priority_bikes:[];
       if(!low.length){wrap.innerHTML='<span style="opacity:.72">⚡ 無低電車</span>';}
       else {
         let rows=low.map(b=>`<span style="display:inline-block;margin:2px 7px 2px 0;${Number(b.battery_power)<=Number(args.priority_threshold)?'color:#dc2626;font-weight:800;':''}">柱 ${esc(b.pillar_no||'—')}｜${esc(b.battery_power)}%</span>`).join('');
         const stale=result.source==='stale_cache'?`<div style="color:#d97706">⚠ 顯示 ${Math.round(Number(result.age_seconds||0))} 秒前資料</div>`:'';
         wrap.innerHTML=`<div style="font-weight:800">⚡ 低電 ${low.length} 台${pri.length?`｜🔴 ${pri.length} 台`:''}</div>${rows}${stale}`;
       }
     } else if(result && result.error){
       wrap.innerHTML='<span style="color:#b91c1c">⚠ 柱號暫時無法取得</span>';
     } else if(!args.auto_query){
       const btn=doc.createElement('button'); btn.type='button'; btn.textContent='⚡ 查柱號';
       btn.style.cssText='border:0;border-radius:999px;padding:5px 10px;background:#111827;color:white;font-size:12px;font-weight:700;cursor:pointer;';
       btn.onclick=()=>{btn.disabled=true;btn.textContent='查詢中…';setValue({action:'query_station',station_name:name,district:String(spec.district||''),nonce:Date.now()});};
       wrap.appendChild(btn);
     } else {wrap.innerHTML='<span style="opacity:.65">⚡ 電池資料準備中</span>';}
     host.appendChild(wrap);
   }
 }
 window.addEventListener('message',e=>{if(!e.data||e.data.type!=='streamlit:render')return;render(e.data.args||{});setHeight();});
 send('streamlit:componentReady',{apiVersion:API});setHeight();
})();
</script></body></html>'''


FLOATING_COMPONENT_HTML = r'''<!doctype html><html><head><meta charset="utf-8"><style>html,body{margin:0;padding:0;background:transparent}</style></head><body>
<script>
(()=>{
 const API=1; const HISTORY_KEY='ubikeBatteryOverlay'; let args={};
 function send(type,data={}){window.parent.postMessage({isStreamlitMessage:true,type,...data},"*");}
 function setValue(value){send('streamlit:setComponentValue',{value});}
 function setHeight(){send('streamlit:setFrameHeight',{height:1});}
 function esc(s){return String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
 function ensure(){
  const doc=window.parent.document;
  let fab=doc.getElementById('ubike-server-battery-fab');
  if(!fab){fab=doc.createElement('button');fab.id='ubike-server-battery-fab';fab.type='button';fab.textContent='⚡';fab.style.cssText='position:fixed;right:18px;bottom:110px;z-index:999990;width:52px;height:52px;border:0;border-radius:50%;background:#111827;color:white;font-size:23px;box-shadow:0 8px 28px rgba(0,0,0,.28);cursor:pointer;';doc.body.appendChild(fab);}
  let page=doc.getElementById('ubike-server-battery-page');
  if(!page){page=doc.createElement('div');page.id='ubike-server-battery-page';page.style.cssText='display:none;position:fixed;inset:0;z-index:999991;width:100%;height:100vh;height:100dvh;box-sizing:border-box;background:#f7f8fb;color:#111827;overflow:auto;overscroll-behavior:contain;padding:max(env(safe-area-inset-top),8px) 12px max(env(safe-area-inset-bottom),8px);';doc.body.appendChild(page);}
  fab.onclick=openPage;
  const win=doc.defaultView;
  if(page._ubikePopHandler)win.removeEventListener('popstate',page._ubikePopHandler);
  page._ubikePopHandler=()=>{page._ubikeHistory=false;hidePage();};
  win.addEventListener('popstate',page._ubikePopHandler);
  return {doc,fab,page};
 }
 function hidePage(){
   const {doc,fab,page}=ensure();
   page.style.display='none';fab.style.display='';
   doc.documentElement.style.overflow=page._ubikeHtmlOverflow||'';
   doc.body.style.overflow=page._ubikeBodyOverflow||'';
   const y=Number(page._ubikeScrollY||0);doc.defaultView.requestAnimationFrame(()=>doc.defaultView.scrollTo(0,y));
 }
 function openPage(){
   const {doc,fab,page}=ensure();
   if(page.style.display==='block'){renderPage();return;}
   page._ubikeScrollY=doc.defaultView.scrollY||0;
   page._ubikeHtmlOverflow=doc.documentElement.style.overflow;
   page._ubikeBodyOverflow=doc.body.style.overflow;
   page.style.display='block';fab.style.display='none';
   doc.documentElement.style.overflow='hidden';doc.body.style.overflow='hidden';
   renderPage();
   try{const state={...(doc.defaultView.history.state||{})};state[HISTORY_KEY]=true;doc.defaultView.history.pushState(state,'',doc.defaultView.location.href);page._ubikeHistory=true;}catch(_){page._ubikeHistory=false;}
 }
 function closePage(){
   const {doc,page}=ensure();
   if(page._ubikeHistory&&doc.defaultView.history.state&&doc.defaultView.history.state[HISTORY_KEY]){
     doc.defaultView.history.back();
     doc.defaultView.setTimeout(()=>{if(page.style.display==='block')hidePage();},350);
   }else{hidePage();}
 }
 function resultCards(results){
   const rows=Object.values(results||{}).filter(x=>x&&(!x.error)&&Number(x.low_count||0)>0);
   rows.sort((a,b)=>Number(b.priority_count||0)-Number(a.priority_count||0)||Number(b.low_count||0)-Number(a.low_count||0)||String(a.requested_name||'').localeCompare(String(b.requested_name||''),'zh-Hant'));
   if(!rows.length)return '<div style="padding:28px;text-align:center;opacity:.65">目前沒有低於門檻的 2.0E</div>';
   return rows.map(r=>{const bikes=(r.low_bikes||[]).map(b=>`<div style="padding:4px 0;${Number(b.battery_power)<=Number(r.priority_threshold)?'color:#dc2626;font-weight:800;':''}">柱號 ${esc(b.pillar_no||'—')}｜${esc(b.bike_no||'')}｜${esc(b.battery_power)}%</div>`).join('');const stale=r.source==='stale_cache'?`<div style="color:#d97706;font-size:12px">⚠ 即時更新失敗，顯示 ${Math.round(Number(r.age_seconds||0))} 秒前資料</div>`:'';return `<details style="background:white;border:1px solid #e5e7eb;border-radius:14px;padding:12px 14px;margin:8px 0"><summary style="font-weight:850;cursor:pointer">${esc(r.requested_name||r.station_name)}｜低電 ${r.low_count} 台${Number(r.priority_count||0)?`｜🔴 ${r.priority_count}`:''}</summary><div style="padding-top:8px">${bikes}${stale}</div></details>`;}).join('');
 }
 function renderPage(){
   const {page}=ensure(); const map=args.route_station_map||{}; const zones=Object.keys(map); const results=args.results||{};
   let pref={};try{pref=JSON.parse(localStorage.getItem('ubike-v30-battery-pref')||'{}')||{};}catch(_){pref={};}
   const selected=Array.isArray(pref.zones)?pref.zones.filter(z=>zones.includes(z)):[];
   const threshold=Number.isFinite(Number(pref.threshold))?Number(pref.threshold):Number(args.threshold||89);
   const priority=Number.isFinite(Number(pref.priority_threshold))?Number(pref.priority_threshold):Number(args.priority_threshold||40);
   page.innerHTML=`<div style="max-width:820px;margin:0 auto;padding:0 4px 70px"><div style="position:sticky;top:0;z-index:2;display:flex;align-items:center;justify-content:space-between;gap:12px;background:#f7f8fb;padding:10px 0 12px"><button id="ub30-close" aria-label="返回原本頁面" style="min-width:76px;min-height:44px;border:0;background:#e5e7eb;border-radius:999px;padding:9px 14px;font-size:16px;font-weight:800;touch-action:manipulation">‹ 返回</button><div style="min-width:0;text-align:right"><div style="font-size:23px;font-weight:900">⚡ 電量查詢</div><div style="font-size:12px;opacity:.65">由伺服器查詢並共用快取</div></div></div><div style="background:white;border:1px solid #e5e7eb;border-radius:16px;padding:14px;margin-top:8px"><div style="font-weight:800;margin-bottom:8px">選擇範圍</div><div id="ub30-zones" style="display:flex;flex-wrap:wrap;gap:8px">${zones.map(z=>`<label style="padding:7px 10px;border:1px solid #d1d5db;border-radius:999px"><input type="checkbox" value="${esc(z)}" ${selected.includes(z)?'checked':''}> ${esc(z)} <span style="opacity:.55">(${(map[z]||[]).length})</span></label>`).join('')}</div><div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:12px"><label>低電門檻 <input id="ub30-th" type="number" min="0" max="100" value="${threshold}" style="width:68px"></label><label>紅色門檻 <input id="ub30-pr" type="number" min="0" max="100" value="${priority}" style="width:68px"></label><button id="ub30-query" style="border:0;border-radius:10px;background:#111827;color:white;padding:8px 16px;font-weight:800">查詢</button></div></div><div id="ub30-status" style="padding:10px 2px;font-size:13px;opacity:.7">${args.querying?'正在由 Server 查詢…':(args.last_message||'')}</div><div>${resultCards(results)}</div></div>`;
   page.querySelector('#ub30-close').onclick=closePage;
   page.querySelector('#ub30-query').onclick=()=>{const zonesSelected=[...page.querySelectorAll('#ub30-zones input:checked')].map(x=>x.value);const th=Math.max(0,Math.min(100,Number(page.querySelector('#ub30-th').value||89)));const pr=Math.max(0,Math.min(th,Number(page.querySelector('#ub30-pr').value||40)));try{localStorage.setItem('ubike-v30-battery-pref',JSON.stringify({zones:zonesSelected,threshold:th,priority_threshold:pr}));}catch(_){};if(!zonesSelected.length){page.querySelector('#ub30-status').textContent='請至少選擇一個範圍';return;}page.querySelector('#ub30-query').disabled=true;page.querySelector('#ub30-status').textContent='正在由 Server 查詢…';setValue({action:'query_zones',zones:zonesSelected,threshold:th,priority_threshold:pr,nonce:Date.now()});};
 }
 function render(newArgs){args=newArgs||{};const {page}=ensure();if(page.style.display==='block')renderPage();}
 window.addEventListener('message',e=>{if(!e.data||e.data.type!=='streamlit:render')return;render(e.data.args||{});setHeight();});
 send('streamlit:componentReady',{apiVersion:API});ensure();setHeight();
})();
</script></body></html>'''


def _fingerprint(value: object) -> str:
    payload = json.dumps(value, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
    return hashlib.sha1(payload).hexdigest()[:14]


def render_inline_server_battery(
    station_specs: list[tuple[str, str, str]],
    *,
    threshold: int,
    priority_threshold: int,
    mobile_mode: bool,
    auto_query: bool,
    force_station: str = "",
) -> None:
    specs = [
        {"name": str(name).strip(), "kind": str(kind), "target": str(target), "district": ""}
        for name, kind, target in station_specs
        if str(name).strip()
    ]
    if not specs:
        return
    threshold = max(0, min(100, int(threshold)))
    priority_threshold = max(0, min(threshold, int(priority_threshold)))
    fp = _fingerprint([(item["name"], item["target"]) for item in specs])
    state_key = f"server_inline_battery_results::{fp}"
    results = st.session_state.get(state_key)
    if not isinstance(results, dict):
        results = {}

    if auto_query:
        # 智慧調度通常只傳前 10 個候選站；這裡只查真正要顯示的站。
        to_query = [{"name": item["name"], "district": item.get("district", "")} for item in specs[:10]]
        force_names = {str(force_station).strip()} if str(force_station).strip() else set()
        queried = query_stations_for_ui(
            to_query,
            threshold=threshold,
            priority_threshold=priority_threshold,
            force_names=force_names,
        )
        results.update(queried)
        st.session_state[state_key] = results

    component = _component("ubike_server_inline_battery_v30", INLINE_COMPONENT_HTML)
    event = component(
        key=f"ubike_server_inline::{fp}",
        default=None,
        specs=specs,
        results=results,
        threshold=threshold,
        priority_threshold=priority_threshold,
        auto_query=bool(auto_query),
        mobile=bool(mobile_mode),
    )
    if isinstance(event, dict) and event.get("action") == "query_station":
        nonce_key = f"ubike_server_inline_nonce::{fp}"
        nonce = str(event.get("nonce") or "")
        if nonce and nonce != st.session_state.get(nonce_key):
            st.session_state[nonce_key] = nonce
            name = str(event.get("station_name") or "").strip()
            district = str(event.get("district") or "").strip()
            if name:
                try:
                    results[name] = query_station_for_ui(
                        name,
                        district=district,
                        threshold=threshold,
                        priority_threshold=priority_threshold,
                        force=True,
                    )
                except Exception as exc:
                    results[name] = {"requested_name": name, "error": str(exc), "low_bikes": [], "priority_bikes": []}
                st.session_state[state_key] = results
                st.rerun()


def render_floating_server_battery(
    route_station_map: dict[str, list[dict]],
    mobile_mode: bool,
    *,
    threshold: int = 89,
    priority_threshold: int = 40,
) -> None:
    clean_map = {
        str(zone): [
            {"name": str(item.get("name") or "").strip(), "district": str(item.get("district") or "").strip()}
            for item in items
            if str(item.get("name") or "").strip()
        ]
        for zone, items in (route_station_map or {}).items()
        if isinstance(items, list)
    }
    fp = _fingerprint(list(clean_map.keys()))
    state_key = f"server_floating_battery::{fp}"
    state = st.session_state.get(state_key)
    if not isinstance(state, dict):
        state = {"results": {}, "last_message": "", "querying": False}

    component = _component("ubike_server_floating_battery_v30", FLOATING_COMPONENT_HTML)
    event = component(
        key=f"ubike_server_floating::{fp}",
        default=None,
        route_station_map=clean_map,
        results=state.get("results", {}),
        last_message=state.get("last_message", ""),
        querying=bool(state.get("querying")),
        threshold=int(threshold),
        priority_threshold=int(priority_threshold),
        mobile=bool(mobile_mode),
    )
    if isinstance(event, dict) and event.get("action") == "query_zones":
        nonce_key = f"ubike_server_floating_nonce::{fp}"
        nonce = str(event.get("nonce") or "")
        if nonce and nonce != st.session_state.get(nonce_key):
            st.session_state[nonce_key] = nonce
            selected_zones = [str(zone) for zone in event.get("zones", []) if str(zone) in clean_map]
            th = max(0, min(100, int(event.get("threshold", threshold))))
            pr = max(0, min(th, int(event.get("priority_threshold", priority_threshold))))
            stations: list[dict] = []
            seen = set()
            for zone in selected_zones:
                for item in clean_map.get(zone, []):
                    name = item["name"]
                    if name in seen:
                        continue
                    seen.add(name)
                    stations.append(item)
            state["querying"] = True
            st.session_state[state_key] = state
            results = query_stations_for_ui(stations, threshold=th, priority_threshold=pr)
            state["results"] = results
            state["querying"] = False
            ok_count = sum(1 for item in results.values() if not item.get("error"))
            failed_count = len(results) - ok_count
            state["last_message"] = f"完成 {ok_count} 站" + (f"｜{failed_count} 站未取得" if failed_count else "")
            st.session_state[state_key] = state
            st.rerun()

