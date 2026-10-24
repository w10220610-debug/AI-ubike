// Run with jsdom available on NODE_PATH.
const {JSDOM}=require('jsdom'),fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const root=path.join(__dirname,'..','work_order_component');
(async()=>{
 const dom=new JSDOM(fs.readFileSync(path.join(root,'index.html'),'utf8'),{url:'https://local.test/',runScripts:'outside-only'});
 const w=dom.window,events=[];w.postMessage=data=>events.push(data);
 const lookup=fs.readFileSync(path.join(root,'lookup.js'),'utf8');
 const script=fs.readFileSync(path.join(root,'index.html'),'utf8').split('<script>')[1].split('</script>')[0];w.eval(lookup+'\n'+script);
 const q=id=>w.document.getElementById(id),specs=[{name:'A',station_no:'1'},{name:'B',station_no:'bad'},{name:'C',station_no:''}];
 const render=epoch=>w.dispatchEvent(new w.MessageEvent('message',{source:w,data:{type:'streamlit:render',args:{token:'fixture',snapshot:[],reset_epoch:epoch,stations:specs,catalog_pending:false}}}));
 const wait=async predicate=>{for(let i=0;i<200;i++){if(predicate())return;await new Promise(r=>setTimeout(r,5))}throw Error('test timeout')};
 render(0);
 q('bike-numbers').value='０１２３４５６,0123456\n1234567 9999999';
 w.fetch=async url=>({ok:!url.includes('bad'),status:503,json:async()=>({retCode:1,retVal:[{bike_no:'0123456',pillar_no:'03',battery_power:95},{bike_no:'1234567',pillar_no:'07'}]})});
 await q('bike-search').onclick();
 const text=q('bike-status').textContent;
 assert.match(text,/成功查詢 1\/2/);assert.match(text,/失敗 1/);assert.match(text,/未配對 1/);assert.match(text,/9999999/);
 assert.match(q('bike-results').textContent,/95%/);assert.match(q('bike-results').textContent,/未提供/);
 assert.equal(q('bike-use').disabled,true);
 assert.equal(events.filter(x=>x.value?.type==='lookup').length,1);
 assert.equal(events.find(x=>x.value?.type==='lookup').value.matches.length,2);
 assert.ok([...q('bike-results').querySelectorAll('input')].every(x=>x.disabled));
 // API business errors must not count as successful empty stations.
 w.fetch=async()=>({ok:true,json:async()=>({retCode:0,retVal:[]})});
 await q('bike-search').onclick();assert.match(q('bike-status').textContent,/成功查詢 0\/2/);
 // Pending response cannot refill results after reset.
 let release;const gate=new Promise(r=>release=r);
 w.fetch=async()=>{await gate;return{ok:true,json:async()=>[{bike_no:'0123456',pillar_no:'09'}]}};
 const run=q('bike-search').onclick();render(1);release();await run;
 assert.equal(q('bike-numbers').value,'');assert.equal(q('bike-results').textContent,'');assert.equal(q('bike-use').disabled,true);
 // Stale worker creation must terminate without publishing OCR after reset.
 let finishWorker,terminated=false;
 w.Tesseract={createWorker:()=>new Promise(r=>finishWorker=()=>r({terminate:async()=>{terminated=true},setParameters:async()=>{},recognize:async()=>({data:{text:'stale'}})}))};
 Object.defineProperty(q('files'),'files',{configurable:true,value:[{size:100}]});
 const ocr=q('start').onclick();await wait(()=>!!finishWorker);render(2);finishWorker();await ocr;
 assert.equal(terminated,true);assert.equal(events.filter(x=>x.value?.type==='ocr').length,0);
 assert.equal(q('start').disabled,false);assert.equal(q('camera').value,'');
 // A parsed photo must auto-search even when a previous input is present.
 let calls=0;
 w.fetch=async()=>{calls++;return {ok:true,json:async()=>[{bike_no:'1560268',pillar_no:'12'}]}};
 const autoArgs={token:'fixture',snapshot:[],reset_epoch:2,stations:[{name:'NEW',station_no:'1'}],catalog_pending:true,lookup_request:'photo-1',prefill_bike_numbers:['1560268','1590102']};
 const autoRender=()=>w.dispatchEvent(new w.MessageEvent('message',{source:w,data:{type:'streamlit:render',args:JSON.parse(JSON.stringify(autoArgs))}}));
 q('bike-numbers').value='9999999';autoRender();assert.equal(calls,0);
 autoArgs.catalog_pending=false;autoRender();
 await wait(()=>events.some(x=>x.value?.request_id==='photo-1'));
 const autoEvent=events.find(x=>x.value?.request_id==='photo-1').value;
 assert.equal(autoEvent.matches.length,1);assert.equal(autoEvent.matches[0].station_name,'NEW');
 assert.match(q('bike-status').textContent,/1590102/);
 assert.equal(q('bike-numbers').value,'1560268\n1590102');
 autoRender();await new Promise(r=>setTimeout(r,10));assert.equal(calls,1);
 // New draft invalidates an older in-flight response.
 let finishOld;const oldGate=new Promise(r=>finishOld=r);
 w.fetch=async()=>{await oldGate;return {ok:true,json:async()=>[{bike_no:'1560268'}]}};
 autoArgs.lookup_request='old-photo';autoRender();
 autoArgs.lookup_request='new-photo';autoArgs.catalog_pending=true;autoRender();
 finishOld();await new Promise(r=>setTimeout(r,10));
 assert.equal(events.filter(x=>x.value?.request_id==='old-photo').length,0);
 console.log('DOM checks passed: batch/leading-zero lookup, missing power, partial failure, API errors, handoff, reset race, OCR cancellation.');
 dom.window.close();
})().catch(e=>{console.error(e);process.exitCode=1});
