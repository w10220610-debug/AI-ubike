from __future__ import annotations

import json

import streamlit.components.v1 as components

from station_service import StationServiceError, get_station_catalog, match_station


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


def _resolve_station_numbers(
    route_station_map: dict[str, list[dict]],
) -> tuple[dict[str, list[dict]], str]:
    """Resolve Excel station names on the Streamlit server."""
    clean_map = _clean_route_map(route_station_map)
    if not clean_map:
        return {}, ""

    try:
        catalog = get_station_catalog()
    except StationServiceError as exc:
        failed = {
            zone: [
                {
                    **item,
                    "station_no": "",
                    "official_name": "",
                    "match_error": f"伺服器場站清單讀取失敗：{exc}",
                }
                for item in items
            ]
            for zone, items in clean_map.items()
        }
        return failed, str(exc)

    resolved: dict[str, list[dict]] = {}
    for zone, items in clean_map.items():
        zone_items: list[dict] = []
        for item in items:
            matched = match_station(
                item["name"],
                district=item.get("district", ""),
                catalog=catalog,
            )
            if matched is None:
                zone_items.append(
                    {
                        **item,
                        "station_no": "",
                        "official_name": "",
                        "match_error": "找不到可安全配對的 YouBike 場站",
                    }
                )
                continue
            zone_items.append(
                {
                    **item,
                    "station_no": str(matched.get("station_no") or "").strip(),
                    "official_name": str(matched.get("station_name") or item["name"]).strip(),
                    "match_error": "",
                }
            )
        if zone_items:
            resolved[zone] = zone_items
    return resolved, ""


def render_floating_server_battery(
    route_station_map: dict[str, list[dict]],
    mobile_mode: bool,
    *,
    threshold: int = 89,
    priority_threshold: int = 69,
) -> None:
    """V30 Hybrid battery page with per-station reverse display."""
    resolved_map, catalog_error = _resolve_station_numbers(route_station_map)
    args = {
        "route_station_map": resolved_map,
        "catalog_error": catalog_error,
        "threshold": max(0, min(100, int(threshold))),
        "priority_threshold": max(0, min(int(threshold), int(priority_threshold))),
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
 const BATTERY_URL='https://apis.youbike.com.tw/api/front/bike/lists';
 const FRESH_MS=30000, STALE_MS=300000, CONCURRENCY=8, REQUEST_TIMEOUT_MS=8500;
 const PREF_VERSION=2;
 const runtime=win.__ubikeV29FastBattery||(win.__ubikeV29FastBattery={cache:new Map(),run:0});
 if(!(runtime.cache instanceof Map)) runtime.cache=new Map();
 let currentResults={};
 let running=false;
 const reverseStations=new Set();

 function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
 function num(v,f=0){const n=Number(v);return Number.isFinite(n)?n:f;}
 function norm(v){return String(v??'').normalize?.('NFKC').toLowerCase().replace(/臺/g,'台').replace(/^(?:youbike|ubike)\s*2\s*[.．]?\s*0\s*e?\s*[_\-－—:：]*\s*/i,'').replace(/公共自行車租賃站/g,'').replace(/[^0-9a-z\u3400-\u9fff]/g,'');}
 function stationKey(r){return String(r?.requested_name||r?.station_name||'');}
 function encodedStationKey(r){return encodeURIComponent(stationKey(r));}
 function extractList(payload){
   if(Array.isArray(payload))return payload.filter(x=>x&&typeof x==='object');
   if(!payload||typeof payload!=='object')return [];
   for(const key of ['retVal','data','items','result','results','stations']){
     const v=payload[key]; if(Array.isArray(v))return v.filter(x=>x&&typeof x==='object');
     if(v&&typeof v==='object')for(const nk of ['data','items','list','results','stations'])if(Array.isArray(v[nk]))return v[nk].filter(x=>x&&typeof x==='object');
   }
   return [];
 }
 async function fetchJson(url,timeout=REQUEST_TIMEOUT_MS){
   const controller=new AbortController(); const timer=setTimeout(()=>controller.abort(),timeout);
   try{
     const response=await fetch(url,{method:'GET',cache:'no-store',credentials:'omit',headers:{'Accept':'application/json, text/plain, */*'},signal:controller.signal});
     if(!response.ok)throw new Error(`HTTP ${response.status}`);
     return await response.json();
   }finally{clearTimeout(timer);}
 }
 function normalizeBattery(payload){
   const bikes=[]; for(const x of extractList(payload)){
     const raw=Number(x.battery_power); if(!Number.isFinite(raw))continue;
     const power=Math.max(0,Math.min(100,Math.trunc(raw)));
     const bike=String(x.bike_no||'').trim(); if(!bike)continue;
     bikes.push({bike_no:bike,pillar_no:String(x.pillar_no||'').trim(),battery_power:power});
   } return bikes;
 }
 function pillarKey(value){const m=String(value??'').match(/\d+/);return m?Number(m[0]):Number.MAX_SAFE_INTEGER;}
 function pillarSort(a,b){const d=pillarKey(a.pillar_no)-pillarKey(b.pillar_no);if(d)return d;return String(a.pillar_no||'').localeCompare(String(b.pillar_no||''),'zh-Hant',{numeric:true,sensitivity:'base'});}
 async function queryOne(spec,threshold,priority,force){
   const stationNo=String(spec.station_no||'').trim();
   if(!stationNo)throw new Error(spec.match_error||'YouBike 站號尚未配對');
   const cacheKey=stationNo, now=Date.now(), cached=runtime.cache.get(cacheKey);
   if(cached&&!force&&now-cached.at<FRESH_MS)return {...cached.value,source:'memory_cache',age_seconds:(now-cached.at)/1000};
   let lastError=null;
   for(let attempt=0;attempt<2;attempt++){
     try{
       const payload=await fetchJson(`${BATTERY_URL}?station_no=${encodeURIComponent(stationNo)}`);
       const bikes=normalizeBattery(payload);
       const low=bikes.filter(b=>b.battery_power<=threshold).sort(pillarSort);
       const pri=low.filter(b=>b.battery_power<=priority);
       const value={requested_name:spec.name,requested_district:spec.district,station_name:spec.official_name||spec.name,station_no:stationNo,bikes,low_bikes:low,priority_bikes:pri,low_count:low.length,priority_count:pri.length,threshold,priority_threshold:priority,source:'live',age_seconds:0};
       runtime.cache.set(cacheKey,{at:Date.now(),value}); return value;
     }catch(e){lastError=e;if(attempt===0)await new Promise(r=>setTimeout(r,180));}
   }
   if(cached&&now-cached.at<STALE_MS)return {...cached.value,source:'stale_cache',age_seconds:(now-cached.at)/1000,error:String(lastError?.message||lastError)};
   throw lastError||new Error('電池資料查詢失敗');
 }
 function lowMin(r){const lows=Array.isArray(r.low_bikes)?r.low_bikes:[];return lows.length?Math.min(...lows.map(b=>num(b.battery_power,101))):101;}
 function safeBikes(r){return (Array.isArray(r.bikes)?r.bikes:[]).filter(b=>num(b.battery_power,-1)>num(r.threshold,89)).slice().sort(pillarSort);}
 function resultRows(results,priorityEnabled,filter=''){
   const q=norm(filter);
   const rows=Object.values(results||{}).filter(r=>r&&!r.error&&num(r.low_count)>0&&(!q||norm(r.requested_name||r.station_name).includes(q)));
   rows.sort((a,b)=>lowMin(a)-lowMin(b)||num(b.low_count)-num(a.low_count)||String(a.requested_name||'').localeCompare(String(b.requested_name||''),'zh-Hant'));
   if(!rows.length)return '<div class="empty">目前沒有符合條件的低電 2.0E</div>';
   return rows.map(r=>{
     const key=stationKey(r),encoded=encodedStationKey(r),reverse=reverseStations.has(key);
     const min=lowMin(r),urgent=priorityEnabled?num(r.priority_count):0;
     const safe=safeBikes(r);
     const shown=reverse?safe:(r.low_bikes||[]).slice().sort(pillarSort);
     const bikes=shown.map(b=>{
       const p=num(b.battery_power,101);
       if(reverse)return `<div class="bike safe"><span>柱 ${esc(b.pillar_no||'—')}</span><span>${esc(b.bike_no||'')}</span><strong>${p}%</strong></div>`;
       const hot=priorityEnabled&&p<=num(r.priority_threshold,69);
       return `<div class="bike ${hot?'urgent':''}"><span>柱 ${esc(b.pillar_no||'—')}</span><span>${esc(b.bike_no||'')}</span><strong>${p}%</strong></div>`;
     }).join('');
     const stale=r.source==='stale_cache'?`<div class="stale">⚠ 顯示約 ${Math.round(num(r.age_seconds))} 秒前快取</div>`:'';
     const modeNote=reverse?'<div class="reverse-banner">✅ 反向模式｜以下車輛「不用換電池」</div>':'';
     const emptyReverse=reverse&&!shown.length?'<div class="reverse-empty">本場站沒有高於目前低電門檻的車輛</div>':'';
     const safeMin=safe.length?Math.min(...safe.map(b=>num(b.battery_power,101))):null;
     const subtitle=reverse?`✅ 不用換 ${safe.length} 台｜反向顯示`:`低電 ${num(r.low_count)} 台${urgent?`｜緊急 ${urgent} 台`:''}`;
     const badge=reverse?(safeMin===null?'反向':`最低 ${safeMin}%`):`${min}%`;
     const badgeClass=reverse?'min reverse-min':`min ${priorityEnabled&&min<=num(r.priority_threshold,69)?'hot':''}`;
     return `<details class="station ${reverse?'reverse':''}" data-station-key="${esc(encoded)}"><summary><div class="station-main"><strong>${esc(r.requested_name||r.station_name||'')}</strong><small>${subtitle}</small></div><button type="button" class="reverse-toggle ${reverse?'active':''}" data-reverse-key="${esc(encoded)}" aria-label="${reverse?'恢復顯示需要換電池':'反向顯示不需要換電池'}" title="${reverse?'恢復顯示需要換電池':'反向顯示不需要換電池'}">⇄</button><div class="${badgeClass}">${badge}</div></summary><div class="bike-list">${modeNote}${bikes}${emptyReverse}${stale}</div></details>`;
   }).join('');
 }
 function styleText(){return `
#${ROOT}{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft JhengHei",sans-serif;color:#eaf7ff}
#ub-v29-fab{position:fixed;right:18px;bottom:110px;z-index:2147482600;width:54px;height:54px;border:1px solid rgba(99,235,255,.55);border-radius:50%;background:linear-gradient(145deg,#081524,#101f35);color:#6cf1ff;font-size:24px;box-shadow:0 10px 30px rgba(0,0,0,.4),0 0 22px rgba(44,215,255,.18)}
#ub-v29-page{display:none;position:fixed;inset:0;z-index:2147482601;background:radial-gradient(circle at 80% 0%,rgba(255,54,147,.16),transparent 32%),radial-gradient(circle at 0% 20%,rgba(28,224,255,.13),transparent 30%),#07101c;overflow:auto;overscroll-behavior:contain}
#ub-v29-page.open{display:block}.shell{max-width:900px;margin:0 auto;padding:max(10px,env(safe-area-inset-top)) 12px max(30px,env(safe-area-inset-bottom))}
header{position:sticky;top:0;z-index:5;display:flex;justify-content:space-between;align-items:center;gap:14px;padding:9px 0 12px;background:linear-gradient(#07101c 75%,transparent)}header h1{font-size:24px;margin:0}header p{margin:3px 0 0;font-size:12px;color:#89a9bd}
#ub-v29-close{min-width:82px;min-height:44px;border:1px solid #27465d;border-radius:999px;background:#0d1c2b;color:#e9fbff;font-size:16px;font-weight:800}
.panel{background:rgba(12,28,43,.92);border:1px solid rgba(94,205,229,.23);border-radius:17px;padding:14px;margin:8px 0}.title-row,.actions,.summary{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.title-row{justify-content:space-between}.mini{border:1px solid #294b63;background:#10283a;color:#bdeffc;border-radius:999px;padding:6px 10px;font-weight:750}.zones{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}.zone{border:1px solid #2a4e66;border-radius:999px;padding:7px 10px;background:#0a1c2a;color:#d8f6ff}.zone input{accent-color:#4feaff}.controls{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:12px}.field{display:flex;flex-direction:column;gap:5px;color:#9db8c8;font-size:12px}.field input[type=number],#ub-v29-search{box-sizing:border-box;width:100%;min-height:42px;border:1px solid #2c526b;border-radius:11px;background:#081825;color:#eefcff;padding:8px 10px;font-size:16px}.toggle{display:flex;align-items:center;gap:8px;margin-top:10px}.actions{margin-top:12px}.actions button{flex:1;min-width:125px;min-height:44px;border:0;border-radius:12px;font-size:15px;font-weight:900}.query{background:linear-gradient(90deg,#43e6ff,#74f3c5);color:#041018}.force{background:#38152c;color:#ffb8d6;border:1px solid #7e2d59!important}.status{padding:10px 2px;color:#9dc1d1;font-size:13px}.updated{font-size:12px;color:#789aaa}.summary{margin:8px 0}.chip{border:1px solid #2b5269;border-radius:999px;padding:6px 9px;background:#0b2030;color:#bfefff;font-size:12px}.chip.hot{border-color:#8b2d5b;color:#ff9bc7}
.station{background:rgba(11,26,40,.96);border:1px solid #24495e;border-radius:15px;margin:9px 0;overflow:hidden;transition:border-color .16s,box-shadow .16s,background .16s}.station summary{list-style:none;display:grid;grid-template-columns:minmax(0,1fr) 42px auto;align-items:center;gap:10px;padding:13px 14px}.station summary::-webkit-details-marker{display:none}.station-main{display:flex;min-width:0;flex-direction:column;gap:4px}.station-main strong{overflow-wrap:anywhere}.station summary small{color:#8db0c1}.reverse-toggle{width:40px;height:36px;padding:0;border:1px solid #315c72;border-radius:10px;background:#0d2637;color:#81eaff;font-size:21px;font-weight:900;line-height:1;box-shadow:inset 0 0 12px rgba(75,222,255,.06)}.reverse-toggle.active{border-color:#55efb0;background:#0e4939;color:#b8ffe3;box-shadow:0 0 18px rgba(61,238,166,.22)}.min{min-width:54px;text-align:center;border-radius:999px;padding:7px 9px;background:#103447;color:#6cecff;font-weight:900;white-space:nowrap}.min.hot{background:#4d1734;color:#ff8cbd}.bike-list{border-top:1px solid #1f3b4d;padding:8px 13px 12px}.bike{display:grid;grid-template-columns:74px 1fr 52px;gap:8px;align-items:center;padding:7px 3px;border-bottom:1px dashed rgba(110,160,180,.18);color:#cce7f2}.bike strong{text-align:right;color:#66edff}.bike.urgent,.bike.urgent strong{color:#ff83b7;font-weight:850}.stale{padding:8px 0 0;color:#ffb968;font-size:12px}.empty{padding:30px 12px;text-align:center;color:#7fa4b6}.progress{height:7px;border-radius:999px;background:#102838;overflow:hidden;margin:7px 0}.progress>i{display:block;height:100%;background:linear-gradient(90deg,#42e8ff,#ff52a8);width:0%;transition:width .2s}
.station.reverse{background:linear-gradient(145deg,rgba(8,48,43,.98),rgba(8,35,38,.98));border:2px solid rgba(76,239,172,.85);box-shadow:0 0 0 1px rgba(65,255,181,.10),0 0 22px rgba(34,224,150,.14)}.station.reverse summary{background:linear-gradient(90deg,rgba(38,172,116,.12),rgba(47,235,164,.04))}.station.reverse summary small{color:#8ff0c6}.station.reverse .bike-list{border-top-color:rgba(76,239,172,.34)}.reverse-min{min-width:72px;background:#155540;color:#baffdf;border:1px solid rgba(85,239,177,.35);font-size:12px}.reverse-banner{margin:1px 0 8px;padding:9px 10px;border:1px solid rgba(81,241,177,.45);border-radius:10px;background:rgba(32,151,104,.14);color:#b9ffe0;font-weight:900;font-size:13px}.bike.safe,.bike.safe strong{color:#83f2c2;font-weight:850}.reverse-empty{padding:15px 4px;text-align:center;color:#91d9bc;font-weight:750}
@media(max-width:560px){.controls{grid-template-columns:1fr}.shell{padding-left:10px;padding-right:10px}.bike{grid-template-columns:65px 1fr 48px}.station summary{grid-template-columns:minmax(0,1fr) 38px auto;gap:7px;padding:12px 10px}.reverse-toggle{width:36px;height:34px}.min{padding:7px 7px}.reverse-min{min-width:65px;font-size:11px}}
`;}
 function ensure(){
   let root=doc.getElementById(ROOT);
   if(!root){
     root=doc.createElement('div');root.id=ROOT;
     root.innerHTML=`<button id="ub-v29-fab" aria-label="電量查詢">⚡</button><section id="ub-v29-page" aria-hidden="true"><div class="shell"><header><button id="ub-v29-close">‹ 返回</button><div><h1>⚡ 電量查詢</h1><p>V30 Hybrid｜站號 Server 配對｜逐站回填</p></div></header><main id="ub-v29-main"></main></div></section>`;
     doc.body.appendChild(root);
   }
   const fab=root.querySelector('#ub-v29-fab');
   const page=root.querySelector('#ub-v29-page');
   const closeButton=root.querySelector('#ub-v29-close');
   if(fab)fab.onclick=open;
   if(closeButton)closeButton.onclick=close;
   if(fab&&page&&!page.classList.contains('open'))fab.style.display='';
   let style=doc.getElementById(ROOT+'-style');if(!style){style=doc.createElement('style');style.id=ROOT+'-style';doc.head.appendChild(style);}style.textContent=styleText();
   const subtitle=root.querySelector('header p');if(subtitle)subtitle.textContent='V30 Hybrid｜站號 Server 配對｜逐站回填';
   return root;
 }
 function open(){const root=ensure(),page=root.querySelector('#ub-v29-page');root.querySelector('#ub-v29-fab').style.display='none';page.classList.add('open');page.setAttribute('aria-hidden','false');root._htmlOverflow=doc.documentElement.style.overflow;root._bodyOverflow=doc.body.style.overflow;doc.documentElement.style.overflow='hidden';doc.body.style.overflow='hidden';render();}
 function close(){const root=ensure(),page=root.querySelector('#ub-v29-page');page.classList.remove('open');page.setAttribute('aria-hidden','true');root.querySelector('#ub-v29-fab').style.display='';doc.documentElement.style.overflow=root._htmlOverflow||'';doc.body.style.overflow=root._bodyOverflow||'';reverseStations.clear();}
 function prefs(){try{return JSON.parse(localStorage.getItem('ubike-v29-fast-battery-pref')||'{}')||{};}catch(_){return {};}}
 function savePrefs(v){try{localStorage.setItem('ubike-v29-fast-battery-pref',JSON.stringify(v));}catch(_){}}
 function rerenderResults(filter=''){
   const root=ensure(),box=root.querySelector('#ub-results');if(!box)return;
   const openKeys=new Set([...box.querySelectorAll('details.station[open]')].map(x=>x.getAttribute('data-station-key')||''));
   box.innerHTML=resultRows(currentResults,root.querySelector('#ub-pe')?.checked!==false,filter);
   for(const item of box.querySelectorAll('details.station'))if(openKeys.has(item.getAttribute('data-station-key')||''))item.open=true;
 }
 function bindResultControls(root){
   const box=root.querySelector('#ub-results');if(!box)return;
   box.onclick=e=>{
     const btn=e.target.closest?.('.reverse-toggle');if(!btn)return;
     e.preventDefault();e.stopPropagation();
     let key='';try{key=decodeURIComponent(btn.getAttribute('data-reverse-key')||'');}catch(_){key='';}
     if(!key)return;
     if(reverseStations.has(key))reverseStations.delete(key);else reverseStations.add(key);
     rerenderResults(root.querySelector('#ub-v29-search')?.value||'');
   };
 }
 function render(){
   const root=ensure(),main=root.querySelector('#ub-v29-main'),map=args.route_station_map||{},zones=Object.keys(map),p=prefs();
   const selected=Array.isArray(p.zones)?p.zones.filter(z=>zones.includes(z)):[];
   const th=Number.isFinite(Number(p.threshold))?Number(p.threshold):num(args.threshold,89);
   const prefVersion=num(p.pref_version,0);
   const pr=(prefVersion>=PREF_VERSION&&Number.isFinite(Number(p.priority_threshold)))?Number(p.priority_threshold):num(args.priority_threshold,69);
   const pe=p.priority_enabled!==false;
   main.innerHTML=`<section class="panel"><div class="title-row"><strong>選擇範圍</strong><div><button class="mini" id="ub-all">全選</button> <button class="mini" id="ub-none">取消</button></div></div><div class="zones">${zones.map(z=>`<label class="zone"><input data-zone type="checkbox" value="${esc(z)}" ${selected.includes(z)?'checked':''}> ${esc(z)} <span style="opacity:.55">(${(map[z]||[]).length})</span></label>`).join('')}</div><div class="controls"><label class="field">低電門檻<input id="ub-th" type="number" min="0" max="100" value="${th}"></label><label class="field">紅色門檻<input id="ub-pr" type="number" min="0" max="100" value="${pr}"></label></div><label class="toggle"><input id="ub-pe" type="checkbox" ${pe?'checked':''}>啟用紅色緊急門檻</label><div class="actions"><button class="query" id="ub-query">查詢</button><button class="force" id="ub-force">強制更新</button></div></section><div class="status" id="ub-status">${args.catalog_error?`⚠ 站號服務：${esc(args.catalog_error)}`:(running?'正在查詢…':'選擇範圍後開始查詢')}</div><div class="progress"><i id="ub-progress"></i></div><div class="summary" id="ub-summary"></div><div class="updated" id="ub-updated"></div><div style="margin-top:10px"><input id="ub-v29-search" placeholder="搜尋場站"></div><div id="ub-results">${resultRows(currentResults,pe,'')}</div>`;
   root.querySelector('#ub-all').onclick=()=>root.querySelectorAll('[data-zone]').forEach(x=>x.checked=true);
   root.querySelector('#ub-none').onclick=()=>root.querySelectorAll('[data-zone]').forEach(x=>x.checked=false);
   root.querySelector('#ub-query').onclick=()=>runQuery(false);
   root.querySelector('#ub-force').onclick=()=>runQuery(true);
   root.querySelector('#ub-v29-search').oninput=e=>rerenderResults(e.target.value);
   bindResultControls(root);
 }
 function summarize(done,total,failed){
   const vals=Object.values(currentResults).filter(r=>r&&!r.error);const lowStations=vals.filter(r=>num(r.low_count)>0).length;const lowBikes=vals.reduce((s,r)=>s+num(r.low_count),0);const urgent=vals.reduce((s,r)=>s+num(r.priority_count),0);const root=ensure();
   root.querySelector('#ub-summary').innerHTML=`<span class="chip">完成 ${done}/${total}</span><span class="chip">低電場站 ${lowStations}</span><span class="chip">低電 ${lowBikes} 台</span>${urgent?`<span class="chip hot">緊急 ${urgent} 台</span>`:''}${failed?`<span class="chip hot">失敗 ${failed}</span>`:''}`;
   rerenderResults(root.querySelector('#ub-v29-search').value);
   root.querySelector('#ub-progress').style.width=`${total?Math.round(done/total*100):0}%`;
 }
 async function runQuery(force){
   if(running)return;
   const root=ensure(),map=args.route_station_map||{};const selected=[...root.querySelectorAll('[data-zone]:checked')].map(x=>x.value);const status=root.querySelector('#ub-status');if(!selected.length){status.textContent='請至少選擇一個範圍';return;}
   const th=Math.max(0,Math.min(100,num(root.querySelector('#ub-th').value,89)));
   const pr=Math.max(0,Math.min(th,num(root.querySelector('#ub-pr').value,69)));
   const pe=root.querySelector('#ub-pe').checked;
   savePrefs({zones:selected,threshold:th,priority_threshold:pr,priority_enabled:pe,pref_version:PREF_VERSION});
   const specs=[],seen=new Set();for(const z of selected)for(const item of (map[z]||[])){if(!item?.name||seen.has(item.name))continue;seen.add(item.name);specs.push(item);}
   currentResults={};running=true;runtime.run+=1;const runId=runtime.run;let index=0,done=0,failed=0;root.querySelectorAll('#ub-query,#ub-force').forEach(b=>b.disabled=true);status.textContent=`正在查詢 ${specs.length} 個場站…`;summarize(0,specs.length,0);
   async function worker(){while(index<specs.length&&runId===runtime.run){const spec=specs[index++];try{currentResults[spec.name]=await queryOne(spec,th,pr,force);}catch(e){failed++;currentResults[spec.name]={requested_name:spec.name,error:String(e?.message||e),low_count:0,priority_count:0};}finally{done++;status.textContent=`查詢中：${done}/${specs.length}｜失敗 ${failed}`;summarize(done,specs.length,failed);}}}
   await Promise.all(Array.from({length:Math.min(CONCURRENCY,specs.length)},worker));
   if(runId!==runtime.run)return;
   running=false;status.textContent=`查詢完成：${done-failed} 站成功${failed?`｜${failed} 站未取得`:''}`;root.querySelector('#ub-updated').textContent=`最後更新：${new Date().toLocaleString('zh-TW',{hour12:false})}`;root.querySelectorAll('#ub-query,#ub-force').forEach(b=>b.disabled=false);summarize(done,specs.length,failed);
 }
 ensure();render();
})();
</script></body></html>'''.replace('__ARGS__', payload)
    components.html(html_text, height=0, scrolling=False)
