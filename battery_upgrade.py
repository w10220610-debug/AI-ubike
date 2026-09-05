from __future__ import annotations

import json

import streamlit as st
import streamlit.components.v1 as components


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
    """V29 Fast Client 電池查詢頁。

    保留新版 UI，但把大量電池請求移回使用者瀏覽器直接執行：
    - 不阻塞 Streamlit Python rerun
    - 8 路並行、逐站更新畫面
    - 30 秒 fresh cache / 5 分鐘 stale cache
    - 單站失敗不拖住其他場站
    """
    clean_map = _clean_route_map(route_station_map)
    args = {
        "route_station_map": clean_map,
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
 const CATALOG_URL='https://apis.youbike.com.tw/json/station-min-yb2.json';
 const BATTERY_URL='https://apis.youbike.com.tw/api/front/bike/lists';
 const FRESH_MS=30000, STALE_MS=300000, CONCURRENCY=8, REQUEST_TIMEOUT_MS=8500;
 const runtime=win.__ubikeV29FastBattery||(win.__ubikeV29FastBattery={catalog:null,catalogAt:0,cache:new Map(),run:0});
 let currentResults={};
 let running=false;

 function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
 function num(v,f=0){const n=Number(v);return Number.isFinite(n)?n:f;}
 function norm(v){return String(v??'').normalize?.('NFKC').toLowerCase().replace(/臺/g,'台').replace(/^(?:youbike|ubike)\s*2\s*[.．]?\s*0\s*e?\s*[_\-－—:：]*\s*/i,'').replace(/公共自行車租賃站/g,'').replace(/[^0-9a-z\u3400-\u9fff]/g,'');}
 function first(...values){for(const v of values){if(v!==undefined&&v!==null&&String(v).trim()!=='')return v;}return null;}
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
 async function catalog(){
   const now=Date.now(); if(runtime.catalog&&now-runtime.catalogAt<21600000)return runtime.catalog;
   const payload=await fetchJson(CATALOG_URL,12000); const rows=extractList(payload); const out=[];
   for(const x of rows){
     const no=String(first(x.station_no,x.sno,x.station_id)||'').trim();
     const name=String(first(x.name_tw,x.sna,x.station_name)||'').trim();
     if(!no||!name)continue;
     out.push({no,name,key:norm(name),district:String(first(x.district_tw,x.sarea)||'').trim(),districtKey:norm(first(x.district_tw,x.sarea)||'')});
   }
   if(!out.length)throw new Error('YouBike 場站清單沒有可辨識資料');
   runtime.catalog=out; runtime.catalogAt=now; return out;
 }
 function matchStation(spec,rows){
   const key=norm(spec.name), dk=norm(spec.district); if(!key)return null;
   let candidates=rows.filter(r=>r.key===key);
   if(dk&&candidates.length>1){const area=candidates.filter(r=>r.districtKey===dk||r.districtKey.includes(dk)||dk.includes(r.districtKey));if(area.length)candidates=area;}
   if(candidates.length)return candidates[0];
   candidates=rows.filter(r=>key.length>=4&&(r.key.includes(key)||key.includes(r.key)));
   if(dk&&candidates.length>1){const area=candidates.filter(r=>r.districtKey===dk||r.districtKey.includes(dk)||dk.includes(r.districtKey));if(area.length)candidates=area;}
   return candidates.length===1?candidates[0]:null;
 }
 function normalizeBattery(payload){
   const bikes=[]; for(const x of extractList(payload)){
     const power=Math.max(0,Math.min(100,Math.trunc(num(x.battery_power,-1))));
     const bike=String(x.bike_no||'').trim(); if(!bike||power<0)continue;
     bikes.push({bike_no:bike,pillar_no:String(x.pillar_no||'').trim(),battery_power:power});
   } return bikes;
 }
 async function queryOne(spec,threshold,priority,force){
   const rows=await catalog(); const station=matchStation(spec,rows);
   if(!station)throw new Error('找不到可安全配對的 YouBike 場站');
   const cacheKey=station.no; const now=Date.now(); const cached=runtime.cache.get(cacheKey);
   if(cached&&!force&&now-cached.at<FRESH_MS)return {...cached.value,source:'memory_cache',age_seconds:(now-cached.at)/1000};
   let lastError=null;
   for(let attempt=0;attempt<2;attempt++){
     try{
       const payload=await fetchJson(`${BATTERY_URL}?station_no=${encodeURIComponent(station.no)}`);
       const bikes=normalizeBattery(payload);
       const low=bikes.filter(b=>b.battery_power<=threshold).sort((a,b)=>a.battery_power-b.battery_power||String(a.pillar_no).localeCompare(String(b.pillar_no)));
       const pri=low.filter(b=>b.battery_power<=priority);
       const value={requested_name:spec.name,requested_district:spec.district,station_name:station.name,station_no:station.no,bikes,low_bikes:low,priority_bikes:pri,low_count:low.length,priority_count:pri.length,threshold,priority_threshold:priority,source:'live',age_seconds:0};
       runtime.cache.set(cacheKey,{at:Date.now(),value}); return value;
     }catch(e){lastError=e;if(attempt===0)await new Promise(r=>setTimeout(r,180));}
   }
   if(cached&&now-cached.at<STALE_MS)return {...cached.value,source:'stale_cache',age_seconds:(now-cached.at)/1000,error:String(lastError?.message||lastError)};
   throw lastError||new Error('電池資料查詢失敗');
 }
 function resultRows(results,priorityEnabled,filter=''){
   const q=norm(filter); const rows=Object.values(results||{}).filter(r=>r&&!r.error&&num(r.low_count)>0&&(!q||norm(r.requested_name||r.station_name).includes(q)));
   const minBattery=r=>{const lows=Array.isArray(r.low_bikes)?r.low_bikes:[];return lows.length?Math.min(...lows.map(b=>num(b.battery_power,101))):101;};
   rows.sort((a,b)=>minBattery(a)-minBattery(b)||num(b.low_count)-num(a.low_count)||String(a.requested_name||'').localeCompare(String(b.requested_name||''),'zh-Hant'));
   if(!rows.length)return '<div class="empty">目前沒有符合條件的低電 2.0E</div>';
   return rows.map(r=>{const min=minBattery(r),urgent=priorityEnabled?num(r.priority_count):0;const bikes=(r.low_bikes||[]).slice().sort((a,b)=>num(a.battery_power,101)-num(b.battery_power,101)||String(a.pillar_no||'').localeCompare(String(b.pillar_no||''))).map(b=>{const p=num(b.battery_power,101),hot=priorityEnabled&&p<=num(r.priority_threshold,40);return `<div class="bike ${hot?'urgent':''}"><span>柱 ${esc(b.pillar_no||'—')}</span><span>${esc(b.bike_no||'')}</span><strong>${p}%</strong></div>`;}).join('');const stale=r.source==='stale_cache'?`<div class="stale">⚠ 顯示約 ${Math.round(num(r.age_seconds))} 秒前快取</div>`:'';return `<details class="station"><summary><div><strong>${esc(r.requested_name||r.station_name||'')}</strong><small>低電 ${num(r.low_count)} 台${urgent?`｜緊急 ${urgent} 台`:''}</small></div><div class="min ${priorityEnabled&&min<=num(r.priority_threshold,40)?'hot':''}">${min}%</div></summary><div class="bike-list">${bikes}${stale}</div></details>`;}).join('');
 }
 function ensure(){
   let root=doc.getElementById(ROOT); if(root)return root;
   root=doc.createElement('div'); root.id=ROOT;
   root.innerHTML=`<button id="ub-v29-fab" aria-label="電量查詢">⚡</button><section id="ub-v29-page" aria-hidden="true"><div class="shell"><header><button id="ub-v29-close">‹ 返回</button><div><h1>⚡ 電量查詢</h1><p>V29 Fast Client｜逐站回填｜不阻塞主畫面</p></div></header><main id="ub-v29-main"></main></div></section>`;
   const style=doc.createElement('style'); style.id=ROOT+'-style'; style.textContent=`
#${ROOT}{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft JhengHei",sans-serif;color:#eaf7ff}#ub-v29-fab{position:fixed;right:18px;bottom:110px;z-index:2147482600;width:54px;height:54px;border:1px solid rgba(99,235,255,.55);border-radius:50%;background:linear-gradient(145deg,#081524,#101f35);color:#6cf1ff;font-size:24px;box-shadow:0 10px 30px rgba(0,0,0,.4),0 0 22px rgba(44,215,255,.18)}#ub-v29-page{display:none;position:fixed;inset:0;z-index:2147482601;background:radial-gradient(circle at 80% 0%,rgba(255,54,147,.16),transparent 32%),radial-gradient(circle at 0% 20%,rgba(28,224,255,.13),transparent 30%),#07101c;overflow:auto;overscroll-behavior:contain}#ub-v29-page.open{display:block}.shell{max-width:900px;margin:0 auto;padding:max(10px,env(safe-area-inset-top)) 12px max(30px,env(safe-area-inset-bottom))}header{position:sticky;top:0;z-index:5;display:flex;justify-content:space-between;align-items:center;gap:14px;padding:9px 0 12px;background:linear-gradient(#07101c 75%,transparent)}header h1{font-size:24px;margin:0}header p{margin:3px 0 0;font-size:12px;color:#89a9bd}#ub-v29-close{min-width:82px;min-height:44px;border:1px solid #27465d;border-radius:999px;background:#0d1c2b;color:#e9fbff;font-size:16px;font-weight:800}.panel{background:rgba(12,28,43,.92);border:1px solid rgba(94,205,229,.23);border-radius:17px;padding:14px;margin:8px 0}.title-row,.actions,.summary{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.title-row{justify-content:space-between}.mini{border:1px solid #294b63;background:#10283a;color:#bdeffc;border-radius:999px;padding:6px 10px;font-weight:750}.zones{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}.zone{border:1px solid #2a4e66;border-radius:999px;padding:7px 10px;background:#0a1c2a;color:#d8f6ff}.zone input{accent-color:#4feaff}.controls{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:12px}.field{display:flex;flex-direction:column;gap:5px;color:#9db8c8;font-size:12px}.field input[type=number],#ub-v29-search{box-sizing:border-box;width:100%;min-height:42px;border:1px solid #2c526b;border-radius:11px;background:#081825;color:#eefcff;padding:8px 10px;font-size:16px}.toggle{display:flex;align-items:center;gap:8px;margin-top:10px}.actions{margin-top:12px}.actions button{flex:1;min-width:125px;min-height:44px;border:0;border-radius:12px;font-size:15px;font-weight:900}.query{background:linear-gradient(90deg,#43e6ff,#74f3c5);color:#041018}.force{background:#38152c;color:#ffb8d6;border:1px solid #7e2d59!important}.status{padding:10px 2px;color:#9dc1d1;font-size:13px}.updated{font-size:12px;color:#789aaa}.summary{margin:8px 0}.chip{border:1px solid #2b5269;border-radius:999px;padding:6px 9px;background:#0b2030;color:#bfefff;font-size:12px}.chip.hot{border-color:#8b2d5b;color:#ff9bc7}.station{background:rgba(11,26,40,.96);border:1px solid #24495e;border-radius:15px;margin:9px 0;overflow:hidden}.station summary{list-style:none;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:13px 14px}.station summary div:first-child{display:flex;flex-direction:column;gap:4px}.station summary small{color:#8db0c1}.min{min-width:54px;text-align:center;border-radius:999px;padding:7px 9px;background:#103447;color:#6cecff;font-weight:900}.min.hot{background:#4d1734;color:#ff8cbd}.bike-list{border-top:1px solid #1f3b4d;padding:8px 13px 12px}.bike{display:grid;grid-template-columns:74px 1fr 52px;gap:8px;align-items:center;padding:7px 3px;border-bottom:1px dashed rgba(110,160,180,.18);color:#cce7f2}.bike strong{text-align:right;color:#66edff}.bike.urgent,.bike.urgent strong{color:#ff83b7;font-weight:850}.stale{padding:8px 0 0;color:#ffb968;font-size:12px}.empty{padding:30px 12px;text-align:center;color:#7fa4b6}.progress{height:7px;border-radius:999px;background:#102838;overflow:hidden;margin:7px 0}.progress>i{display:block;height:100%;background:linear-gradient(90deg,#42e8ff,#ff52a8);width:0%;transition:width .2s}@media(max-width:560px){.controls{grid-template-columns:1fr}.shell{padding-left:10px;padding-right:10px}.bike{grid-template-columns:65px 1fr 48px}}
`;
   doc.head.appendChild(style);doc.body.appendChild(root);root.querySelector('#ub-v29-fab').onclick=open;root.querySelector('#ub-v29-close').onclick=close;return root;
 }
 function open(){const root=ensure(),page=root.querySelector('#ub-v29-page');root.querySelector('#ub-v29-fab').style.display='none';page.classList.add('open');page.setAttribute('aria-hidden','false');root._htmlOverflow=doc.documentElement.style.overflow;root._bodyOverflow=doc.body.style.overflow;doc.documentElement.style.overflow='hidden';doc.body.style.overflow='hidden';render();}
 function close(){const root=ensure(),page=root.querySelector('#ub-v29-page');page.classList.remove('open');page.setAttribute('aria-hidden','true');root.querySelector('#ub-v29-fab').style.display='';doc.documentElement.style.overflow=root._htmlOverflow||'';doc.body.style.overflow=root._bodyOverflow||'';}
 function prefs(){try{return JSON.parse(localStorage.getItem('ubike-v29-fast-battery-pref')||'{}')||{};}catch(_){return {};}}
 function savePrefs(v){try{localStorage.setItem('ubike-v29-fast-battery-pref',JSON.stringify(v));}catch(_){}}
 function render(){
   const root=ensure(),main=root.querySelector('#ub-v29-main'),map=args.route_station_map||{},zones=Object.keys(map),p=prefs();
   const selected=Array.isArray(p.zones)?p.zones.filter(z=>zones.includes(z)):[];const th=Number.isFinite(Number(p.threshold))?Number(p.threshold):num(args.threshold,89);const pr=Number.isFinite(Number(p.priority_threshold))?Number(p.priority_threshold):num(args.priority_threshold,40);const pe=p.priority_enabled!==false;
   main.innerHTML=`<section class="panel"><div class="title-row"><strong>選擇範圍</strong><div><button class="mini" id="ub-all">全選</button> <button class="mini" id="ub-none">取消</button></div></div><div class="zones">${zones.map(z=>`<label class="zone"><input data-zone type="checkbox" value="${esc(z)}" ${selected.includes(z)?'checked':''}> ${esc(z)} <span style="opacity:.55">(${(map[z]||[]).length})</span></label>`).join('')}</div><div class="controls"><label class="field">低電門檻<input id="ub-th" type="number" min="0" max="100" value="${th}"></label><label class="field">紅色門檻<input id="ub-pr" type="number" min="0" max="100" value="${pr}"></label></div><label class="toggle"><input id="ub-pe" type="checkbox" ${pe?'checked':''}>啟用紅色緊急門檻</label><div class="actions"><button class="query" id="ub-query">查詢</button><button class="force" id="ub-force">強制更新</button></div></section><div class="status" id="ub-status">${running?'正在查詢…':'選擇範圍後開始查詢'}</div><div class="progress"><i id="ub-progress"></i></div><div class="summary" id="ub-summary"></div><div class="updated" id="ub-updated"></div><div style="margin-top:10px"><input id="ub-v29-search" placeholder="搜尋場站"></div><div id="ub-results">${resultRows(currentResults,pe,'')}</div>`;
   root.querySelector('#ub-all').onclick=()=>root.querySelectorAll('[data-zone]').forEach(x=>x.checked=true);root.querySelector('#ub-none').onclick=()=>root.querySelectorAll('[data-zone]').forEach(x=>x.checked=false);root.querySelector('#ub-query').onclick=()=>runQuery(false);root.querySelector('#ub-force').onclick=()=>runQuery(true);root.querySelector('#ub-v29-search').oninput=e=>{root.querySelector('#ub-results').innerHTML=resultRows(currentResults,root.querySelector('#ub-pe').checked,e.target.value)};
 }
 function summarize(done,total,failed){
   const vals=Object.values(currentResults).filter(r=>r&&!r.error);const lowStations=vals.filter(r=>num(r.low_count)>0).length;const lowBikes=vals.reduce((s,r)=>s+num(r.low_count),0);const urgent=vals.reduce((s,r)=>s+num(r.priority_count),0);const root=ensure();root.querySelector('#ub-summary').innerHTML=`<span class="chip">完成 ${done}/${total}</span><span class="chip">低電場站 ${lowStations}</span><span class="chip">低電 ${lowBikes} 台</span>${urgent?`<span class="chip hot">緊急 ${urgent} 台</span>`:''}${failed?`<span class="chip hot">失敗 ${failed}</span>`:''}`;root.querySelector('#ub-results').innerHTML=resultRows(currentResults,root.querySelector('#ub-pe').checked,root.querySelector('#ub-v29-search').value);root.querySelector('#ub-progress').style.width=`${total?Math.round(done/total*100):0}%`;
 }
 async function runQuery(force){
   if(running)return;const root=ensure(),map=args.route_station_map||{};const selected=[...root.querySelectorAll('[data-zone]:checked')].map(x=>x.value);const status=root.querySelector('#ub-status');if(!selected.length){status.textContent='請至少選擇一個範圍';return;}
   const th=Math.max(0,Math.min(100,num(root.querySelector('#ub-th').value,89)));const pr=Math.max(0,Math.min(th,num(root.querySelector('#ub-pr').value,40)));const pe=root.querySelector('#ub-pe').checked;savePrefs({zones:selected,threshold:th,priority_threshold:pr,priority_enabled:pe});
   const specs=[],seen=new Set();for(const z of selected)for(const item of (map[z]||[])){if(!item?.name||seen.has(item.name))continue;seen.add(item.name);specs.push(item);}currentResults={};running=true;runtime.run+=1;const runId=runtime.run;let index=0,done=0,failed=0;root.querySelectorAll('#ub-query,#ub-force').forEach(b=>b.disabled=true);status.textContent=`正在準備 ${specs.length} 個場站…`;summarize(0,specs.length,0);
   try{await catalog();}catch(e){status.textContent=`場站清單讀取失敗：${e?.message||e}`;running=false;root.querySelectorAll('#ub-query,#ub-force').forEach(b=>b.disabled=false);return;}
   async function worker(){while(index<specs.length&&runId===runtime.run){const spec=specs[index++];try{currentResults[spec.name]=await queryOne(spec,th,pr,force);}catch(e){failed++;currentResults[spec.name]={requested_name:spec.name,error:String(e?.message||e),low_count:0,priority_count:0};}finally{done++;status.textContent=`查詢中：${done}/${specs.length}｜失敗 ${failed}`;summarize(done,specs.length,failed);}}}
   await Promise.all(Array.from({length:Math.min(CONCURRENCY,specs.length)},worker));if(runId!==runtime.run)return;running=false;status.textContent=`查詢完成：${done-failed} 站成功${failed?`｜${failed} 站未取得`:''}`;root.querySelector('#ub-updated').textContent=`最後更新：${new Date().toLocaleString('zh-TW',{hour12:false})}`;root.querySelectorAll('#ub-query,#ub-force').forEach(b=>b.disabled=false);summarize(done,specs.length,failed);
 }
 ensure(); render();
})();
</script></body></html>'''.replace('__ARGS__', payload)
    components.html(html_text, height=0, scrolling=False)
