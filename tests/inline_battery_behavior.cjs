'use strict';
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '..', 'app.py'), 'utf8');
const start = source.indexOf('<script>', source.indexOf('def _build_inline_low_battery_pillars_html'));
const raw = source.slice(start + 8, source.indexOf('</script>', start));
class Element {
  constructor() {
    this.children=[]; this.dataset={}; this.events={}; this.className=''; this.textContent='';
    this.classList={contains:c=>this.className.split(' ').includes(c),
      add:(...cs)=>{this.className += ' '+cs.join(' ');},
      remove:(...cs)=>{this.className=this.className.split(' ').filter(c=>!cs.includes(c)).join(' ');},
      toggle:(c,on)=>on?this.classList.add(c):this.classList.remove(c)};
  }
  appendChild(e){this.children.push(e);return e;}
  replaceChildren(){this.children=[];}
  addEventListener(name,fn){this.events[name]=fn;}
  querySelector(){return null;}
  remove(){}
}
function setup(autoQuery=false) {
  const host=new Element(), timers=[], pending=[];
  const doc={head:new Element(),createElement:()=>new Element(),querySelectorAll:()=>[],getElementById:id=>id==='station'?host:null};
  const service={normalizeStationName:s=>s,queryStationByName:()=>new Promise((resolve,reject)=>pending.push({resolve,reject}))};
  const win={document:doc,matchMedia:()=>({matches:false}),setTimeout:(fn,delay)=>timers.push({fn,delay})};
  const script=raw.replace('__LOW_BATTERY_CLIENT_CORE__','').replace('__STATION_SPECS__',JSON.stringify([{name:'測試站',kind:'analysis',target:'station'}]))
    .replace('__THRESHOLD__','89').replace('__PRIORITY_THRESHOLD__','40').replace('__AUTO_QUERY__',String(autoQuery))
    .replace('__FORCE_STATION__','""').replace('__DISPLAY_MODE__','"desktop"').replace('__FINGERPRINT__','"test"');
  vm.runInNewContext(script,{window:{parent:win},ensureUbikeBatteryService:()=>service});
  return {host,pending,start:()=>timers.find(t=>t.delay===480).fn(),win};
}
const data={bikes:[{pillar_no:'8',battery_power:41},{pillar_no:'3',battery_power:40},{pillar_no:'1',battery_power:10},{pillar_no:'9',battery_power:89},{pillar_no:'10',battery_power:90}]};
const flush=()=>new Promise(resolve=>setImmediate(resolve));
const click=e=>e.events.click({preventDefault(){},stopPropagation(){}});
const pills=h=>h.children[0].children.filter(e=>e.classList.contains('ubike-pillar')).map(e=>e.textContent);
const button=(h,text)=>h.children[0].children.find(e=>e.textContent===text);
(async()=>{
  const s=setup(); const running=s.start();
  assert.ok(button(s.host,'⚡ 查低電量柱號'),'first threshold remains accessible while loading');
  s.pending[0].resolve(data); await running;
  assert.deepEqual(pills(s.host),['⚠ 01柱','⚠ 03柱']);
  click(button(s.host,'↻')); s.pending[1].resolve(data); await flush();
  assert.deepEqual(pills(s.host),['⚠ 01柱','⚠ 03柱'],'refresh must not expand first threshold');
  click(button(s.host,'⚡ 查低電量柱號')); s.pending[2].resolve(data); await flush();
  assert.deepEqual(pills(s.host),['⚠ 01柱','⚠ 03柱','08柱','09柱']);
  const empty=setup(); const e=empty.start(); empty.pending[0].resolve({bikes:[{pillar_no:'8',battery_power:41}]}); await e;
  assert.equal(pills(empty.host).length,0); assert.ok(empty.host.children[0].children.some(n=>n.textContent==='第二門檻 ≤40%：無'));
  const race=setup(); const r=race.start(); click(button(race.host,'⚡ 查低電量柱號'));
  race.pending[1].resolve(data); await flush(); race.pending[0].reject(new Error('old request')); await r;
  assert.equal(pills(race.host).length,4,'old error must not replace newer result');
  const fail=setup(); const f=fail.start(); fail.pending[0].reject(new Error('offline')); await f;
  assert.ok(button(fail.host,'重新查詢')); assert.ok(button(fail.host,'⚡ 查低電量柱號'));
  const smart=setup(true); const q=smart.start(); smart.pending[0].resolve(data); await q;
  assert.equal(pills(smart.host).length,4,'smart dispatch must keep full automatic display');
  console.log('PASS: automatic second threshold, manual first threshold, boundaries, sorting, refresh, empty, errors, request races and smart dispatch');
})().catch(e=>{console.error(e);process.exitCode=1;});
