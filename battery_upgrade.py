from __future__ import annotations

import json

import streamlit.components.v1 as components

import ai_learning_guard as _ai_learning_guard_module
from persistent_learning_pool import install_persistent_learning_pool
from station_service import StationServiceError, get_station_catalog, match_station


install_persistent_learning_pool(_ai_learning_guard_module)


def _clean_route_map(route_station_map: dict[str, list[dict]]) -> dict[str, list[dict]]:
    output: dict[str, list[dict]] = {}
    for zone, items in (route_station_map or {}).items():
        zone_name = str(zone or "").strip()
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
    """Resolve station number, district and coordinates on the Streamlit server."""
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
                    "official_district": item.get("district", ""),
                    "latitude": None,
                    "longitude": None,
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
                        "official_district": item.get("district", ""),
                        "latitude": None,
                        "longitude": None,
                        "match_error": "找不到可安全配對的 YouBike 場站",
                    }
                )
                continue
            zone_items.append(
                {
                    **item,
                    "station_no": str(matched.get("station_no") or "").strip(),
                    "official_name": str(matched.get("station_name") or item["name"]).strip(),
                    "official_district": str(
                        matched.get("district") or item.get("district") or ""
                    ).strip(),
                    "latitude": matched.get("latitude"),
                    "longitude": matched.get("longitude"),
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
    """Reliable V29-old battery page with live location and district summary."""
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
 const PREF_VERSION=4;
 const runtime=win.__ubikeV29FastBattery||(win.__ubikeV29FastBattery={cache:new Map(),run:0});
 if(!(runtime.cache instanceof Map))runtime.cache=new Map();
 if(!runtime.locationState)runtime.locationState={lat:null,lon:null,accuracy:null,updatedAt:0,error:'',watchId:null};
 let currentResults={};
 let running=false;
 const reverseStations=new Set();

 function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
 function num(v,f=0){const n=Number(v);return Number.isFinite(n)?n:f;}
 function norm(v){return String(v??'').normalize?.('NFKC').toLowerCase().replace(/臺/g,'台').replace(/^(?:youbike|ubike)\s*2\s*[.．]?\s*0\s*e?\s*[_\-－—:：]*\s*/i,'').replace(/公共自行車租賃站/g,'').replace(/[^0-9a-z\u3400-\u9fff]/g,'');}
 function validCoord(v,min,max){const n=Number(v);return Number.isFinite(n)&&n>=min&&n<=max?n:null;}
 function stationKey(r){return String(r?.requested_name||r?.station_name||'');}
 function encodedStationKey(r){return encodeURIComponent(stationKey(r));}
 function districtOf(r){return String(r?.requested_district||r?.official_district||'未分類').trim()||'未分類';}
 function hasLocation(){return validCoord(runtime.locationState.lat,-90,90)!==null&&validCoord(runtime.locationState.lon,-180,180)!==null;}
 function haversineMeters(lat1,lon1,lat2,lon2){
   const a1=validCoord(lat1,-90,90),o1=validCoord(lon1,-180,180),a2=validCoord(lat2,-90,90),o2=validCoord(lon2,-180,180);
   if([a1,o1,a2,o2].some(v=>v===null))return Infinity;
   const R=6371000,toRad=v=>v*Math.PI/180,dLat=toRad(a2-a1),dLon=toRad(o2-o1);
   const s=Math.sin(dLat/2)**2+Math.cos(toRad(a1))*Math.cos(toRad(a2))*Math.sin(dLon/2)**2;
   return 2*R*Math.asin(Math.min(1,Math.sqrt(s)));
 }
 function distanceFor(r){return hasLocation()?haversineMeters(runtime.locationState.lat,runtime.locationState.lon,r?.latitude,r?.longitude):Infinity;}
 function distanceLabel(m){if(!Number.isFinite(m))return '';if(m<1000)return `約 ${Math.round(m)} 公尺`;const km=m/1000;return `約 ${km<10?km.toFixed(1):Math.round(km)} 公里`;}
 function extractList(payload){
   if(Array.isArray(payload))return payload.filter(x=>x&&typeof x==='object');
   if(!payload||typeof payload!=='object')return [];
   for(const key of ['retVal','data','items','result','results','stations']){
     const v=payload[key];
     if(Array.isArray(v))return v.filter(x=>x&&typeof x==='object');
     if(v&&typeof v==='object')for(const nk of ['data','items','list','results','stations'])if(Array.isArray(v[nk]))return v[nk].filter(x=>x&&typeof x==='object');
   }
   return [];
 }
 async function fetchJson(url,timeout=REQUEST_TIMEOUT_MS){
   const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),timeout);
   try{
     const response=await fetch(url,{method:'GET',cache:'no-store',credentials:'omit',headers:{'Accept':'application/json, text/plain, */*'},signal:controller.signal});
     if(!response.ok)throw new Error(`HTTP ${response.status}`);
     return await response.json();
   }finally{clearTimeout(timer);}
 }
 function normalizeBattery(payload){
   const bikes=[];
   for(const x of extractList(payload)){
     const raw=Number(x.battery_power);if(!Number.isFinite(raw))continue;
     const power=Math.max(0,Math.min(100,Math.trunc(raw))),bike=String(x.bike_no||'').trim();if(!bike)continue;
     bikes.push({bike_no:bike,pillar_no:String(x.pillar_no||'').trim(),battery_power:power});
   }
   return bikes;
 }
 function pillarKey(v){const m=String(v??'').match(/\d+/);return m?Number(m[0]):Number.MAX_SAFE_INTEGER;}
 function pillarSort(a,b){const d=pillarKey(a.pillar_no)-pillarKey(b.pillar_no);return d||String(a.pillar_no||'').localeCompare(String(b.pillar_no||''),'zh-Hant',{numeric:true,sensitivity:'base'});}
 async function queryOne(spec,threshold,priority,force){
   const stationNo=String(spec.station_no||'').trim();if(!stationNo)throw new Error(spec.match_error||'YouBike 站號尚未配對');
   const cacheKey=stationNo,now=Date.now(),cached=runtime.cache.get(cacheKey);
   const enrich=value=>({...value,latitude:validCoord(spec.latitude,-90,90),longitude:validCoord(spec.longitude,-180,180),requested_district:spec.district||spec.official_district||value.requested_district||'',official_district:spec.official_district||spec.district||value.official_district||''});
   if(cached&&!force&&now-cached.at<FRESH_MS)return enrich({...cached.value,source:'memory_cache',age_seconds:(now-cached.at)/1000});
   let lastError=null;
   for(let attempt=0;attempt<2;attempt++){
     try{
       const data=await fetchJson(`${BATTERY_URL}?station_no=${encodeURIComponent(stationNo)}`),bikes=normalizeBattery(data);
       const low=bikes.filter(b=>b.battery_power<=threshold).sort(pillarSort),pri=low.filter(b=>b.battery_power<=priority);
       const value=enrich({requested_name:spec.name,station_name:spec.official_name||spec.name,station_no:stationNo,bikes,low_bikes:low,priority_bikes:pri,low_count:low.length,priority_count:pri.length,threshold,priority_threshold:priority,source:'live',age_seconds:0});
       runtime.cache.set(cacheKey,{at:Date.now(),value});return value;
     }catch(e){lastError=e;if(attempt===0)await new Promise(r=>setTimeout(r,180));}
   }
   if(cached&&now-cached.at<STALE_MS)return enrich({...cached.value,source:'stale_cache',age_seconds:(now-cached.at)/1000,error:String(lastError?.message||lastError)});
   throw lastError||new Error('電池資料查詢失敗');
 }
 function lowMin(r){const lows=Array.isArray(r.low_bikes)?r.low_bikes:[];return lows.length?Math.min(...lows.map(b=>num(b.battery_power,101))):101;}
 function safeBikes(r){return (Array.isArray(r.bikes)?r.bikes:[]).filter(b=>num(b.battery_power,-1)>num(r.threshold,89)).slice().sort(pillarSort);}
 function resultRows(results,priorityEnabled,filter=''){
   const q=norm(filter),rows=Object.values(results||{}).filter(r=>r&&!r.error&&num(r.low_count)>0&&(!q||norm(r.requested_name||r.station_name).includes(q)||norm(districtOf(r)).includes(q)));
   rows.sort((a,b)=>{
     if(hasLocation()){
       const ad=distanceFor(a),bd=distanceFor(b);
       if(Number.isFinite(ad)||Number.isFinite(bd)){const d=ad-bd;if(Math.abs(d)>0.5)return d;}
     }
     return num(b.low_count)-num(a.low_count)||lowMin(a)-lowMin(b)||String(a.requested_name||'').localeCompare(String(b.requested_name||''),'zh-Hant');
   });
   if(!rows.length)return '<div class="empty">目前沒有符合條件的低電 2.0E</div>';
   return rows.map((r,index)=>{
     const key=stationKey(r),encoded=encodedStationKey(r),reverse=reverseStations.has(key),min=lowMin(r),urgent=priorityEnabled?num(r.priority_count):0,safe=safeBikes(r),shown=reverse?safe:(r.low_bikes||[]).slice().sort(pillarSort);
     const bikes=shown.map(b=>{const p=num(b.battery_power,101),hot=priorityEnabled&&p<=num(r.priority_threshold,69);return `<div class="bike ${reverse?'safe':hot?'urgent':''}"><span>柱 ${esc(b.pillar_no||'—')}</span><span>${esc(b.bike_no||'')}</span><strong>${p}%</strong></div>`;}).join('');
     const dist=distanceFor(r),distText=distanceLabel(dist),nearest=index===0&&Number.isFinite(dist)?'<span class="nearest">最近推薦</span>':'';
     const subtitle=reverse?`✅ 不用換 ${safe.length} 台｜${esc(districtOf(r))}${distText?`｜${distText}`:''}`:`${nearest}${esc(districtOf(r))}${distText?`｜${distText}`:''}｜需換 ${num(r.low_count)} 顆${urgent?`｜緊急 ${urgent}`:''}`;
     const badge=reverse?(safe.length?`最低 ${Math.min(...safe.map(b=>num(b.battery_power,101)))}%`:'反向'):`需換 ${num(r.low_count)} 顆`,badgeClass=reverse?'min reverse-min':`min ${priorityEnabled&&min<=num(r.priority_threshold,69)?'hot':''}`;
     const mode=reverse?'<div class="reverse-banner">✅ 反向模式｜以下車輛不用換電池</div>':'';
     const stale=r.source==='stale_cache'?`<div class="stale">⚠ 顯示約 ${Math.round(num(r.age_seconds))} 秒前快取</div>`:'';
     return `<details class="station ${reverse?'reverse':''}" data-station-key="${esc(encoded)}"><summary><div class="station-main"><strong>${esc(r.requested_name||r.station_name||'')}</strong><small>${subtitle}</small></div><button type="button" class="reverse-toggle ${reverse?'active':''}" data-reverse-key="${esc(encoded)}">⇄</button><div class="${badgeClass}">${badge}</div></summary><div class="bike-list">${mode}${bikes}${reverse&&!shown.length?'<div class="reverse-empty">本場站沒有高於目前門檻的車輛</div>':''}${stale}</div></details>`;
   }).join('');
 }
 function districtSummaryHtml(results){
   const grouped=new Map();
   for(const r of Object.values(results||{})){
     if(!r||r.error||num(r.low_count)<=0)continue;
     const district=districtOf(r),cur=grouped.get(district)||{district,total:0,stations:0,minDistance:Infinity};
     cur.total+=num(r.low_count);cur.stations+=1;cur.minDistance=Math.min(cur.minDistance,distanceFor(r));grouped.set(district,cur);
   }
   const items=[...grouped.values()].sort((a,b)=>b.total-a.total||a.minDistance-b.minDistance||a.district.localeCompare(b.district,'zh-Hant'));
   if(!items.length)return '';
   const total=items.reduce((s,x)=>s+x.total,0);
   return `<section class="district-box"><div class="district-head"><strong>📍 各行政區需更換電池｜共 ${total} 顆</strong><small>只顯示至少有 1 顆需要更換的行政區</small></div><div class="district-grid">${items.map(x=>`<div class="district-card"><span>${esc(x.district)}｜需更換</span><strong>${x.total} 顆</strong><small>${x.stations} 個場站${Number.isFinite(x.minDistance)?`｜最近 ${distanceLabel(x.minDistance)}`:''}</small></div>`).join('')}</div></section>`;
 }
 function styleText(){return `
#${ROOT}{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft JhengHei",sans-serif;color:#eaf7ff}
#ub-v29-fab{position:fixed!important;right:18px!important;bottom:110px!important;z-index:2147483000!important;width:54px!important;height:54px!important;display:flex;align-items:center;justify-content:center;border:1px solid rgba(99,235,255,.55);border-radius:50%;background:linear-gradient(145deg,#081524,#101f35);color:#6cf1ff;font-size:24px;box-shadow:0 10px 30px rgba(0,0,0,.4),0 0 22px rgba(44,215,255,.18);visibility:visible!important;opacity:1!important;pointer-events:auto!important}
#ub-v29-page{display:none;position:fixed;inset:0;z-index:2147483001;background:radial-gradient(circle at 80% 0%,rgba(255,54,147,.16),transparent 32%),radial-gradient(circle at 0% 20%,rgba(28,224,255,.13),transparent 30%),#07101c;overflow:auto;overscroll-behavior:contain}
#ub-v29-page.open{display:block}.shell{max-width:900px;margin:0 auto;padding:max(10px,env(safe-area-inset-top)) 12px max(30px,env(safe-area-inset-bottom))}
header{position:sticky;top:0;z-index:5;display:flex;justify-content:space-between;align-items:center;gap:14px;padding:9px 0 12px;background:linear-gradient(#07101c 75%,transparent)}header h1{font-size:24px;margin:0}header p{margin:3px 0 0;font-size:12px;color:#89a9bd}#ub-v29-close{min-width:82px;min-height:44px;border:1px solid #27465d;border-radius:999px;background:#0d1c2b;color:#e9fbff;font-size:16px;font-weight:800}
.panel{background:rgba(12,28,43,.92);border:1px solid rgba(94,205,229,.23);border-radius:17px;padding:14px;margin:8px 0}.title-row,.actions,.summary{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.title-row{justify-content:space-between}.mini{border:1px solid #294b63;background:#10283a;color:#bdeffc;border-radius:999px;padding:6px 10px;font-weight:750}.zones{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}.zone{border:1px solid #2a4e66;border-radius:999px;padding:7px 10px;background:#0a1c2a;color:#d8f6ff}.zone input{accent-color:#4feaff}.controls{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:12px}.field{display:flex;flex-direction:column;gap:5px;color:#9db8c8;font-size:12px}.field input[type=number],#ub-v29-search{box-sizing:border-box;width:100%;min-height:42px;border:1px solid #2c526b;border-radius:11px;background:#081825;color:#eefcff;padding:8px 10px;font-size:16px}.toggle{display:flex;align-items:center;gap:8px;margin-top:10px}.actions{margin-top:12px}.actions button{flex:1;min-width:125px;min-height:44px;border:0;border-radius:12px;font-size:15px;font-weight:900}.query{background:linear-gradient(90deg,#43e6ff,#74f3c5);color:#041018}.force{background:#38152c;color:#ffb8d6;border:1px solid #7e2d59!important}.status{padding:10px 2px;color:#9dc1d1;font-size:13px}.updated{font-size:12px;color:#789aaa}.summary{margin:8px 0}.chip{border:1px solid #2b5269;border-radius:999px;padding:6px 9px;background:#0b2030;color:#bfefff;font-size:12px}.chip.hot{border-color:#8b2d5b;color:#ff9bc7}.location-status{margin:9px 0;padding:10px 12px;border:1px solid #36546b;border-radius:12px;background:rgba(9,26,39,.9);color:#b7cbd7;font-size:13px}.location-status.ok{border-color:#287f69;color:#72f2bd}.location-status.bad{border-color:#8c3c60;color:#ff9bc2}
.district-box{margin:13px 0 16px;padding:13px;border:2px solid rgba(255,58,195,.75);border-radius:20px;background:linear-gradient(145deg,rgba(44,9,39,.88),rgba(5,30,37,.9));box-shadow:0 0 22px rgba(255,54,198,.12)}.district-head{display:flex;flex-direction:column;gap:4px;margin-bottom:11px}.district-head strong{font-size:18px}.district-head small{color:#e6a4d1}.district-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:9px}.district-card{min-height:90px;padding:12px;border:2px solid rgba(224,72,174,.5);border-radius:16px;background:linear-gradient(145deg,rgba(63,17,54,.75),rgba(8,25,40,.88));display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center}.district-card span{color:#f5a8d3;font-weight:800}.district-card strong{margin-top:4px;font-size:28px;color:#f7ff49}.district-card small{margin-top:4px;color:#9cb6c6}
.station{background:rgba(11,26,40,.96);border:1px solid #24495e;border-radius:15px;margin:9px 0;overflow:hidden}.station summary{list-style:none;display:grid;grid-template-columns:minmax(0,1fr) 42px auto;align-items:center;gap:10px;padding:13px 14px}.station summary::-webkit-details-marker{display:none}.station-main{display:flex;min-width:0;flex-direction:column;gap:4px}.station-main strong{overflow-wrap:anywhere}.station summary small{color:#8db0c1}.nearest{display:inline-block;margin-right:6px;padding:2px 6px;border-radius:999px;background:#0b5d5c;color:#73ffed;font-size:11px;font-weight:900}.reverse-toggle{width:40px;height:36px;padding:0;border:1px solid #315c72;border-radius:10px;background:#0d2637;color:#81eaff;font-size:21px;font-weight:900}.reverse-toggle.active{border-color:#55efb0;background:#0e4939;color:#b8ffe3}.min{min-width:74px;text-align:center;border-radius:999px;padding:7px 9px;background:#103447;color:#6cecff;font-weight:900;white-space:nowrap}.min.hot{background:#4d1734;color:#ff8cbd}.bike-list{border-top:1px solid #1f3b4d;padding:8px 13px 12px}.bike{display:grid;grid-template-columns:74px 1fr 52px;gap:8px;align-items:center;padding:7px 3px;border-bottom:1px dashed rgba(110,160,180,.18);color:#cce7f2}.bike strong{text-align:right;color:#66edff}.bike.urgent,.bike.urgent strong{color:#ff83b7;font-weight:850}.bike.safe,.bike.safe strong{color:#83f2c2;font-weight:850}.stale{padding:8px 0 0;color:#ffb968;font-size:12px}.empty{padding:30px 12px;text-align:center;color:#7fa4b6}.progress{height:7px;border-radius:999px;background:#102838;overflow:hidden;margin:7px 0}.progress>i{display:block;height:100%;background:linear-gradient(90deg,#42e8ff,#ff52a8);width:0%;transition:width .2s}.station.reverse{background:linear-gradient(145deg,rgba(8,48,43,.98),rgba(8,35,38,.98));border:2px solid rgba(76,239,172,.85)}.reverse-min{min-width:72px;background:#155540;color:#baffdf;border:1px solid rgba(85,239,177,.35);font-size:12px}.reverse-banner{margin:1px 0 8px;padding:9px 10px;border:1px solid rgba(81,241,177,.45);border-radius:10px;background:rgba(32,151,104,.14);color:#b9ffe0;font-weight:900;font-size:13px}.reverse-empty{padding:15px 4px;text-align:center;color:#91d9bc;font-weight:750}
@media(max-width:560px){.controls{grid-template-columns:1fr}.shell{padding-left:10px;padding-right:10px}.bike{grid-template-columns:65px 1fr 48px}.station summary{grid-template-columns:minmax(0,1fr) 38px auto;gap:7px;padding:12px 10px}.reverse-toggle{width:36px;height:34px}.district-card{min-height:82px;padding:10px}.district-card strong{font-size:25px}.district-head strong{font-size:17px}}
`;}
 function buildRoot(){
   let old=doc.getElementById(ROOT);if(old)old.remove();
   const root=doc.createElement('div');root.id=ROOT;
   root.innerHTML=`<button id="ub-v29-fab" type="button" aria-label="電量查詢">⚡</button><section id="ub-v29-page" aria-hidden="true"><div class="shell"><header><button id="ub-v29-close" type="button">‹ 返回</button><div><h1>⚡ 2.0E 電量查詢｜測試版</h1><p>持續定位｜距離排序｜行政區統整</p></div></header><main id="ub-v29-main"></main></div></section>`;
   doc.body.appendChild(root);return root;
 }
 function ensure(){
   let root=doc.getElementById(ROOT);
   const broken=!root||!root.isConnected||!root.querySelector('#ub-v29-fab')||!root.querySelector('#ub-v29-page')||!root.querySelector('#ub-v29-main');
   if(broken)root=buildRoot();
   const fab=root.querySelector('#ub-v29-fab'),page=root.querySelector('#ub-v29-page'),closeButton=root.querySelector('#ub-v29-close');
   fab.onclick=open;closeButton.onclick=close;
   if(!page.classList.contains('open')){fab.style.display='flex';fab.style.visibility='visible';fab.style.opacity='1';fab.style.pointerEvents='auto';fab.hidden=false;fab.removeAttribute('aria-hidden');}
   let style=doc.getElementById(ROOT+'-style');if(!style){style=doc.createElement('style');style.id=ROOT+'-style';doc.head.appendChild(style);}style.textContent=styleText();
   return root;
 }
 function updateLocationUi(){
   const root=ensure(),el=root.querySelector('#ub-location-status');if(!el)return;
   const l=runtime.locationState;
   if(hasLocation()){const age=Math.max(0,Math.round((Date.now()-num(l.updatedAt,Date.now()))/1000));el.className='location-status ok';el.textContent=`定位完成（誤差約 ${Math.round(num(l.accuracy))} 公尺）｜依距離由近到遠排序｜${age} 秒前更新`;}
   else if(l.error){el.className='location-status bad';el.textContent=`定位未取得：${l.error}｜目前改用需更換顆數排序`;}
   else{el.className='location-status';el.textContent='正在取得定位…';}
 }
 function startLocationWatch(){
   try{
     const geo=win.navigator&&win.navigator.geolocation?win.navigator.geolocation:(typeof navigator!=='undefined'?navigator.geolocation:null);
     if(!geo){runtime.locationState.error='此瀏覽器不支援定位';updateLocationUi();return;}
     if(runtime.locationState.watchId!==null&&runtime.locationState.watchId!==undefined){updateLocationUi();return;}
     runtime.locationState.watchId=geo.watchPosition(position=>{
       const c=position?.coords||{},lat=validCoord(c.latitude,-90,90),lon=validCoord(c.longitude,-180,180);if(lat===null||lon===null)return;
       runtime.locationState.lat=lat;runtime.locationState.lon=lon;runtime.locationState.accuracy=Math.max(0,num(c.accuracy,0));runtime.locationState.updatedAt=Date.now();runtime.locationState.error='';updateLocationUi();refreshViews();
     },err=>{
       const messages={1:'定位權限未允許',2:'暫時無法取得位置',3:'定位逾時'};runtime.locationState.error=messages[num(err?.code)]||String(err?.message||'定位失敗');updateLocationUi();refreshViews();
     },{enableHighAccuracy:true,maximumAge:15000,timeout:10000});
   }catch(e){runtime.locationState.error=String(e?.message||e||'定位啟動失敗');updateLocationUi();}
 }
 function open(){const root=ensure(),page=root.querySelector('#ub-v29-page'),fab=root.querySelector('#ub-v29-fab');fab.style.display='none';page.classList.add('open');page.setAttribute('aria-hidden','false');root._htmlOverflow=doc.documentElement.style.overflow;root._bodyOverflow=doc.body.style.overflow;doc.documentElement.style.overflow='hidden';doc.body.style.overflow='hidden';render();startLocationWatch();}
 function close(){const root=ensure(),page=root.querySelector('#ub-v29-page'),fab=root.querySelector('#ub-v29-fab');page.classList.remove('open');page.setAttribute('aria-hidden','true');fab.style.display='flex';doc.documentElement.style.overflow=root._htmlOverflow||'';doc.body.style.overflow=root._bodyOverflow||'';reverseStations.clear();}
 function prefs(){try{return JSON.parse(localStorage.getItem('ubike-v29-fast-battery-pref')||'{}')||{};}catch(_){return {};}}
 function savePrefs(v){try{localStorage.setItem('ubike-v29-fast-battery-pref',JSON.stringify(v));}catch(_){}}
 function refreshViews(filter){
   const root=ensure(),box=root.querySelector('#ub-results');
   if(box){const openKeys=new Set([...box.querySelectorAll('details.station[open]')].map(x=>x.getAttribute('data-station-key')||''));box.innerHTML=resultRows(currentResults,root.querySelector('#ub-pe')?.checked!==false,filter??root.querySelector('#ub-v29-search')?.value??'');for(const item of box.querySelectorAll('details.station'))if(openKeys.has(item.getAttribute('data-station-key')||''))item.open=true;}
   const district=root.querySelector('#ub-district-summary');if(district)district.innerHTML=districtSummaryHtml(currentResults);updateLocationUi();
 }
 function bindResultControls(root){const box=root.querySelector('#ub-results');if(!box)return;box.onclick=e=>{const btn=e.target.closest?.('.reverse-toggle');if(!btn)return;e.preventDefault();e.stopPropagation();let key='';try{key=decodeURIComponent(btn.getAttribute('data-reverse-key')||'');}catch(_){}if(!key)return;if(reverseStations.has(key))reverseStations.delete(key);else reverseStations.add(key);refreshViews();};}
 function render(){
   const root=ensure(),main=root.querySelector('#ub-v29-main'),map=args.route_station_map||{},zones=Object.keys(map),p=prefs();
   const selected=Array.isArray(p.zones)?p.zones.filter(z=>zones.includes(z)):[],th=Number.isFinite(Number(p.threshold))?Number(p.threshold):num(args.threshold,89),pr=Number.isFinite(Number(p.priority_threshold))?Number(p.priority_threshold):num(args.priority_threshold,69),pe=p.priority_enabled!==false;
   main.innerHTML=`<section class="panel"><div class="title-row"><strong>選擇範圍</strong><div><button class="mini" id="ub-all">全選</button> <button class="mini" id="ub-none">取消</button></div></div><div class="zones">${zones.map(z=>`<label class="zone"><input data-zone type="checkbox" value="${esc(z)}" ${selected.includes(z)?'checked':''}> ${esc(z)} <span style="opacity:.55">(${(map[z]||[]).length})</span></label>`).join('')}</div><div class="controls"><label class="field">低電門檻<input id="ub-th" type="number" min="0" max="100" value="${th}"></label><label class="field">紅色門檻<input id="ub-pr" type="number" min="0" max="100" value="${pr}"></label></div><label class="toggle"><input id="ub-pe" type="checkbox" ${pe?'checked':''}>啟用紅色緊急門檻</label><div class="actions"><button class="query" id="ub-query">查詢</button><button class="force" id="ub-force">強制更新</button></div></section><div id="ub-location-status" class="location-status">正在取得定位…</div><div class="status" id="ub-status">${args.catalog_error?`⚠ 站號服務：${esc(args.catalog_error)}`:(running?'正在查詢…':'選擇範圍後開始查詢')}</div><div class="progress"><i id="ub-progress"></i></div><div class="summary" id="ub-summary"></div><div class="updated" id="ub-updated"></div><div id="ub-district-summary">${districtSummaryHtml(currentResults)}</div><div style="margin-top:10px"><input id="ub-v29-search" placeholder="搜尋場站或行政區"></div><div id="ub-results">${resultRows(currentResults,pe,'')}</div>`;
   root.querySelector('#ub-all').onclick=()=>root.querySelectorAll('[data-zone]').forEach(x=>x.checked=true);root.querySelector('#ub-none').onclick=()=>root.querySelectorAll('[data-zone]').forEach(x=>x.checked=false);root.querySelector('#ub-query').onclick=()=>runQuery(false);root.querySelector('#ub-force').onclick=()=>runQuery(true);root.querySelector('#ub-v29-search').oninput=e=>refreshViews(e.target.value);bindResultControls(root);updateLocationUi();
 }
 function summarize(done,total,failed){const vals=Object.values(currentResults).filter(r=>r&&!r.error),lowStations=vals.filter(r=>num(r.low_count)>0).length,lowBikes=vals.reduce((s,r)=>s+num(r.low_count),0),urgent=vals.reduce((s,r)=>s+num(r.priority_count),0),root=ensure();root.querySelector('#ub-summary').innerHTML=`<span class="chip">完成 ${done}/${total}</span><span class="chip">需換場站 ${lowStations}</span><span class="chip">需換 ${lowBikes} 顆</span>${urgent?`<span class="chip hot">緊急 ${urgent} 顆</span>`:''}${failed?`<span class="chip hot">失敗 ${failed}</span>`:''}`;refreshViews();root.querySelector('#ub-progress').style.width=`${total?Math.round(done/total*100):0}%`;}
 async function runQuery(force){
   if(running)return;const root=ensure(),map=args.route_station_map||{},selected=[...root.querySelectorAll('[data-zone]:checked')].map(x=>x.value),status=root.querySelector('#ub-status');if(!selected.length){status.textContent='請至少選擇一個範圍';return;}
   const th=Math.max(0,Math.min(100,num(root.querySelector('#ub-th').value,89))),pr=Math.max(0,Math.min(th,num(root.querySelector('#ub-pr').value,69))),pe=root.querySelector('#ub-pe').checked;savePrefs({zones:selected,threshold:th,priority_threshold:pr,priority_enabled:pe,pref_version:PREF_VERSION});
   const specs=[],seen=new Set();for(const z of selected)for(const item of (map[z]||[])){if(!item?.name||seen.has(item.name))continue;seen.add(item.name);specs.push(item);}currentResults={};running=true;runtime.run+=1;const runId=runtime.run;let index=0,done=0,failed=0;root.querySelectorAll('#ub-query,#ub-force').forEach(b=>b.disabled=true);status.textContent=`正在查詢 ${specs.length} 個場站…`;summarize(0,specs.length,0);
   async function worker(){while(index<specs.length&&runId===runtime.run){const spec=specs[index++];try{currentResults[spec.name]=await queryOne(spec,th,pr,force);}catch(e){failed++;currentResults[spec.name]={requested_name:spec.name,requested_district:spec.district||spec.official_district||'',official_district:spec.official_district||'',latitude:validCoord(spec.latitude,-90,90),longitude:validCoord(spec.longitude,-180,180),error:String(e?.message||e),low_count:0,priority_count:0};}finally{done++;status.textContent=`查詢中：${done}/${specs.length}｜失敗 ${failed}`;summarize(done,specs.length,failed);}}}
   await Promise.all(Array.from({length:Math.min(CONCURRENCY,specs.length)},worker));if(runId!==runtime.run)return;running=false;status.textContent=`查詢完成：${done-failed} 站成功${failed?`｜${failed} 站未取得`:''}`;root.querySelector('#ub-updated').textContent=`最後更新：${new Date().toLocaleString('zh-TW',{hour12:false})}`;root.querySelectorAll('#ub-query,#ub-force').forEach(b=>b.disabled=false);summarize(done,specs.length,failed);
 }
 function repairFloatingButton(){try{const root=ensure(),page=root.querySelector('#ub-v29-page'),fab=root.querySelector('#ub-v29-fab');if(!page.classList.contains('open')){fab.style.display='flex';fab.style.visibility='visible';fab.style.opacity='1';fab.style.pointerEvents='auto';fab.hidden=false;}}catch(_){} }
 if(runtime.safeUiInterval){try{win.clearInterval(runtime.safeUiInterval);}catch(_){}}
 runtime.safeUiInterval=win.setInterval(repairFloatingButton,1000);
 ensure();render();repairFloatingButton();startLocationWatch();
})();
</script></body></html>'''.replace('__ARGS__', payload)
    components.html(html_text, height=0, scrolling=False)
