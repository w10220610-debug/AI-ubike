const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '..', 'legacy_ui.py'), 'utf8');
const html = source.split('YOUBIKE_BROWSER_COMPONENT_HTML = ')[1].split('_YOUBIKE_BROWSER_SYNC_COMPONENT = None')[0];
let js = html.slice(html.indexOf('<script>') + 8, html.indexOf('</script>'));
js = js.replace('  send("streamlit:componentReady", { apiVersion: API_VERSION });',
  '  window.testApi = { runSync, selectConfiguredCatalog, setArgs: value => { args = value; } };');

function harness(fetchImpl) {
  let now = 0;
  const messages = [], listeners = {}, timers = new Map();
  const node = { addEventListener() {}, textContent: '', disabled: false };
  const window = { parent: { postMessage: value => messages.push(value) },
    sessionStorage: { getItem: () => null, setItem() {} },
    addEventListener: (name, fn) => { listeners[name] = fn; } };
  let timerId = 0;
  const context = { window, document: { getElementById: () => node },
    performance: { now: () => now }, Date, Intl, Math, Map, Set, Promise,
    AbortController, fetch: (...args) => fetchImpl(...args),
    setTimeout: (fn, delay) => { timers.set(++timerId, {fn, delay}); return timerId; },
    clearTimeout: id => timers.delete(id) };
  vm.runInNewContext(js, context);
  return { api: window.testApi, messages, listeners, setNow: value => { now = value; }, timers };
}

const catalog = [
  { station_no: '1', name_tw: 'YouBike2.0_測試站', district_tw: '台東市' },
  { station_no: '2', name_tw: '測試站', district_tw: '台北市' },
  { station_no: '3', name_tw: '第二站', district_tw: '台東市' },
  { station_no: '4', name_tw: '其他站', district_tw: '台東市' },
];
const response = data => ({ ok: true, text: async () => JSON.stringify(data) });
const values = h => h.messages.filter(m => m.type === 'streamlit:setComponentValue').map(m => m.value);

(async () => {
  const h = harness(async () => { throw new Error('unexpected request'); });
  h.api.setArgs({ station_specs: [{name: '測試站'}] });
  assert.equal(h.api.selectConfiguredCatalog(catalog).length, 0, 'ambiguous names must not match');
  h.api.setArgs({ station_specs: [{name: '測試站', district: '臺東市'}] });
  assert.equal(h.api.selectConfiguredCatalog(catalog)[0].station_no, '1');

  const requests = [];
  const ok = harness(async (url, options) => {
    if (options.method === 'GET') return response(catalog);
    const ids = JSON.parse(options.body).station_no;
    requests.push(ids);
    return response(ids.map(station_no => ({station_no, available_spaces_detail: {yb2: 3, eyb: 0}})));
  });
  ok.api.setArgs({station_specs: [{name:'第二站'}], max_single_rounds:0});
  await ok.api.runSync();
  assert.deepEqual(requests, [['3']], 'only configured stations are queried');
  assert.equal(values(ok)[0].records[0].general_bikes, 3);
  assert.equal(values(ok)[0].records[0].electric_bikes, 0, 'real zero is retained');

  let count = 0;
  const partial = harness(async (url, options) => {
    if (options.method === 'GET') return response(catalog);
    count++;
    partial.setNow(26000);
    return response([{station_no:'3', available_spaces_detail:{yb2:5}}]);
  });
  partial.api.setArgs({station_specs:[{name:'第二站'}, {name:'其他站'}], max_single_rounds:0});
  await partial.api.runSync();
  assert.equal(count, 1, 'no retries after whole-round deadline');
  assert.equal(values(partial)[0].ok, true);
  assert.equal(values(partial)[0].records.length, 1);
  assert.equal(values(partial)[0].missing_station_count, 1);
  assert.equal(values(partial)[0].records[0].electric_bikes, null, 'missing counts must not become zero');

  const failed = harness(async () => ({ok:false, status:403}));
  failed.api.setArgs({station_specs:[{name:'第二站'}]});
  await failed.api.runSync();
  assert.equal(values(failed)[0].ok, false);
  assert.match(values(failed)[0].error, /403/);
  assert.equal(values(failed)[0].records, undefined);

  let initialRequests = 0;
  const immediate = harness(async () => { initialRequests++; return response([]); });
  const event = {data:{type:'streamlit:render', args:{start_immediately:true, auto_refresh:true}}};
  immediate.listeners.message(event);
  immediate.listeners.message(event);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(initialRequests, 1, 'initial request runs immediately, only once per mount');
  console.log('PASS: scoped/ambiguous station matching, real zero vs missing, partial timeout, HTTP error, immediate single-flight start');
})().catch(err => { console.error(err); process.exitCode = 1; });
