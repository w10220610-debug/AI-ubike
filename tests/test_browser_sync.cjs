const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

// Exercise the actual embedded client, including Streamlit render messages and
// clocks. No YouBike network calls or browser dependencies are needed.
const source = fs.readFileSync(path.join(__dirname, '..', 'app.py'), 'utf8');
const html = source.match(/YOUBIKE_BROWSER_COMPONENT_HTML\s*=\s*r"""([\s\S]*?)"""/)[1];
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];

function catalogPayload() {
  return { retCode: 1, retVal: { data: [{
    station_no: '508001', name_tw: '臺東縣政府文化處圖書館',
    county_tw: '臺東縣', lat: 22.755, lng: 121.151, status: 1,
  }] } };
}

function parkingPayload(general = 3, electric = 2) {
  return { retCode: 1, retVal: { data: [{
    station_no: '508001', available_spaces_detail: { yb2: general, eyb: electric },
    available_spaces: general + electric, empty_spaces: 50,
    parking_spaces: 50 + general + electric, status: 1,
  }] } };
}

function response(payload, status = 200) {
  return { ok: status >= 200 && status < 300, status, text: async () => JSON.stringify(payload) };
}

function deferred() {
  let resolve;
  const promise = new Promise(done => { resolve = done; });
  return { promise, resolve };
}

function harness(fetchOverride) {
  let now = 0;
  let timerId = 0;
  const timers = new Map();
  const listeners = new Map();
  const buttonListeners = new Map();
  const messages = [];
  const requests = [];
  const storage = new Map();
  const button = { disabled: false, textContent: '', addEventListener: (type, fn) => buttonListeners.set(type, fn) };
  const status = { textContent: '', className: '' };
  class ClockDate extends Date {
    constructor(...args) { super(...(args.length ? args : [1727875800000 + now])); }
    static now() { return 1727875800000 + now; }
  }
  const clock = {
    setTimeout(fn, milliseconds = 0) {
      const id = ++timerId;
      timers.set(id, { fn, due: now + Math.max(0, Number(milliseconds)) });
      return id;
    },
    clearTimeout(id) { timers.delete(id); },
  };
  const window = {
    parent: { postMessage(message) { messages.push(message); } },
    sessionStorage: {
      getItem(key) { return storage.get(key) ?? null; },
      setItem(key, value) { storage.set(key, String(value)); },
    },
    addEventListener(type, fn) {
      if (!listeners.has(type)) listeners.set(type, []);
      listeners.get(type).push(fn);
    },
    ...clock,
  };
  const context = vm.createContext({
    window,
    document: { getElementById(id) { return id === 'syncButton' ? button : status; } },
    fetch: async (url, options = {}) => {
      requests.push({ url, options });
      if (fetchOverride) return fetchOverride(url, options, requests.length);
      return response(options.method === 'POST' ? parkingPayload() : catalogPayload());
    },
    AbortController,
    performance: { now: () => now },
    Date: ClockDate,
    Intl,
    console,
    ...clock,
  });
  vm.runInContext(script, context, { filename: 'YOUBIKE_BROWSER_COMPONENT_HTML.js' });
  async function flush() {
    // Complete the promise chain of catalog -> batched parking -> host delivery.
    for (let i = 0; i < 60; i += 1) await Promise.resolve();
  }
  async function advance(milliseconds) {
    const end = now + milliseconds;
    await flush();
    while (true) {
      const next = [...timers.entries()].filter(([, timer]) => timer.due <= end)
        .sort((a, b) => a[1].due - b[1].due || a[0] - b[0])[0];
      if (!next) break;
      now = next[1].due;
      timers.delete(next[0]);
      next[1].fn();
      await flush();
    }
    now = end;
    await flush();
  }
  function send(data) {
    for (const listener of listeners.get('message') || []) listener({ data });
  }
  function render(args = {}) {
    send({ type: 'streamlit:render', args: {
      signature_scope: 'configured-workbook', auto_refresh: true,
      auto_refresh_seconds: 60, max_batch_rounds: 1, max_single_rounds: 0,
      ...args,
    } });
  }
  return {
    render, advance, flush, requests, messages, button, status,
    manual: () => send({ type: 'ubike:manual-sync' }),
    values: () => messages.filter(message => message.type === 'streamlit:setComponentValue').map(message => message.value),
    catalogRequests: () => requests.filter(request => request.options.method !== 'POST').length,
  };
}

test('the first render fetches and delivers current data immediately', async () => {
  const app = harness();
  app.render();
  await app.advance(0);
  assert.equal(app.catalogRequests(), 1);
  const [value] = app.values();
  assert.equal(value.ok, true);
  assert.equal(value.records[0].general_bikes, 3);
  assert.equal(value.records[0].electric_bikes, 2);
});

test('30-second Streamlit rerenders do not postpone the minute refresh', async () => {
  const app = harness();
  app.render();
  await app.advance(0);
  await app.advance(30000);
  app.render();
  await app.advance(29999);
  assert.equal(app.catalogRequests(), 1);
  app.render();
  await app.advance(1);
  assert.equal(app.catalogRequests(), 2);
  await app.advance(30000);
  app.render();
  await app.advance(30000);
  assert.equal(app.catalogRequests(), 3);
});

test('manual update remains single-flight while a request is pending', async () => {
  const firstCatalog = deferred();
  const app = harness((url, options, index) => index === 1
    ? firstCatalog.promise
    : response(options.method === 'POST' ? parkingPayload() : catalogPayload()));
  app.render();
  await app.advance(0);
  app.manual();
  app.manual();
  app.render();
  await app.flush();
  assert.equal(app.requests.length, 1);
  assert.equal(app.button.disabled, true);
  firstCatalog.resolve(response(catalogPayload()));
  await app.flush();
  assert.equal(app.values().length, 1);
  assert.equal(app.button.disabled, false);
  app.manual();
  await app.flush();
  assert.equal(app.catalogRequests(), 2);
  assert.equal(app.values().length, 2, 'manual updates force delivery even when counts are unchanged');
});

test('a failed request reports the error without delivering fake zero counts', async () => {
  const app = harness(() => response({ error: 'temporarily unavailable' }, 503));
  app.render();
  await app.advance(5000);
  const [value] = app.values();
  assert.equal(value.ok, false);
  assert.match(value.error, /503/);
  assert.equal(app.values().filter(item => item.ok).length, 0);
  assert.equal(value.records, undefined);
});

test('an HTTP 200 official API rejection reports its retMsg', async () => {
  const app = harness(() => response({ retCode: 0, retMsg: '官方服務暫時不可用', retVal: { data: [] } }));
  app.render();
  await app.advance(5000);
  const [value] = app.values();
  assert.equal(value.ok, false);
  assert.match(value.error, /官方服務暫時不可用/);
  assert.equal(app.values().filter(item => item.ok).length, 0);
});

test('parking API failures retain their cause and do not replace previously fetched counts', async () => {
  let parkingUnavailable = false;
  const app = harness((url, options) => {
    if (options.method !== 'POST') return response(catalogPayload());
    if (parkingUnavailable) return response({ retCode: 0, retMsg: '場站查詢暫時不可用', retVal: { data: [] } });
    return response(parkingPayload());
  });
  app.render();
  await app.advance(0);
  assert.equal(app.values()[0].records[0].general_bikes, 3);
  parkingUnavailable = true;
  app.manual();
  await app.advance(5000);
  const values = app.values();
  assert.equal(values.length, 2);
  assert.equal(values[1].ok, false);
  assert.match(values[1].error, /場站查詢暫時不可用/);
  assert.equal(values[1].records, undefined);
  assert.equal(values.filter(item => item.ok).length, 1);
});

test('an automatic recovery delivers unchanged counts to clear the previous error', async () => {
  let catalogCount = 0;
  const app = harness((url, options) => {
    if (options.method === 'POST') return response(parkingPayload());
    catalogCount += 1;
    if (catalogCount === 2) return response({ retCode: 0, retMsg: '短暫斷線', retVal: { data: [] } });
    return response(catalogPayload());
  });
  app.render({ force_initial_delivery: false });
  await app.advance(0);
  assert.equal(app.values()[0].ok, true);
  await app.advance(60000);
  assert.equal(app.values()[1].ok, false);
  await app.advance(60000);
  assert.equal(app.values().length, 3);
  assert.equal(app.values()[2].ok, true, 'recovery must notify Python before the five-minute heartbeat');
  assert.equal(app.values()[2].records[0].general_bikes, 3);
  assert.equal(app.values()[2].records[0].electric_bikes, 2);
  await app.advance(60000);
  assert.equal(app.catalogRequests(), 4);
  assert.equal(app.values().length, 3, 'later unchanged results should resume suppressing redundant deliveries');
});

test('real zero counts stay zero and are delivered as valid data', async () => {
  const app = harness((url, options) => response(options.method === 'POST' ? parkingPayload(0, 0) : catalogPayload()));
  app.render();
  await app.advance(0);
  const [value] = app.values();
  assert.equal(value.ok, true);
  assert.equal(value.records[0].general_bikes, 0);
  assert.equal(value.records[0].electric_bikes, 0);
  assert.equal(value.records[0].available_spaces, 0);
});

test('changing workbook scope immediately delivers the new context', async () => {
  const app = harness();
  app.render({ signature_scope: 'first-workbook' });
  await app.advance(0);
  app.render({ signature_scope: 'second-workbook' });
  await app.advance(0);
  assert.equal(app.catalogRequests(), 2);
  assert.equal(app.values().filter(item => item.ok).length, 2);
});

test('a context change while busy discards the old result and starts the latest context', async () => {
  const firstCatalog = deferred();
  const app = harness((url, options, index) => index === 1
    ? firstCatalog.promise
    : response(options.method === 'POST' ? parkingPayload() : catalogPayload()));
  app.render({ signature_scope: 'first-workbook' });
  await app.advance(0);
  app.render({ signature_scope: 'intermediate-workbook' });
  app.render({ signature_scope: 'latest-workbook' });
  await app.flush();
  assert.equal(app.catalogRequests(), 1);
  firstCatalog.resolve(response(catalogPayload()));
  await app.advance(0);
  assert.equal(app.catalogRequests(), 2, 'only the latest context should be queued');
  assert.equal(app.values().filter(item => item.ok).length, 1, 'the old context must not overwrite the latest workbook');
});

test('disabling auto refresh clears its timer while retaining manual updates', async () => {
  const app = harness();
  app.render();
  await app.advance(0);
  app.render({ auto_refresh: false });
  await app.advance(60000);
  assert.equal(app.catalogRequests(), 1);
  app.manual();
  await app.advance(0);
  assert.equal(app.catalogRequests(), 2);
  await app.advance(60000);
  assert.equal(app.catalogRequests(), 2);
});

test('official area codes exclude neighbouring counties while preserving legacy Taitung records', async () => {
  const stations = [
    { station_no: '508001', name_tw: '臺東官方站', area_code: 15 },
    { station_no: '508002', name_tw: '臺東文字備援站', county_tw: '臺東縣' },
    { station_no: '508003', name_tw: '座標備援站', area_code: '', lat: 22.755, lng: 121.151 },
    { station_no: '514001', name_tw: '屏東範圍內站', area_code: '14', lat: 22.00, lng: 120.80 },
    { station_no: '514002', name_tw: '臺東字樣的外縣市站', area_code: '14', county_tw: '臺東縣' },
    { station_no: '500001', name_tw: '範圍外站', lat: 25.05, lng: 121.50 },
  ];
  const app = harness((url, options) => {
    if (options.method !== 'POST') return response({ retCode: 1, retVal: { data: stations } });
    const requested = JSON.parse(options.body).station_no;
    return response({ retCode: 1, retVal: { data: requested.map(station_no => ({
      ...parkingPayload().retVal.data[0], station_no,
    })) } });
  });
  app.render();
  await app.advance(0);
  const [value] = app.values();
  assert.equal(value.ok, true);
  assert.deepEqual(Array.from(value.records, record => record.station_id), ['508001', '508002', '508003']);
  const requested = app.requests.filter(request => request.options.method === 'POST')
    .flatMap(request => JSON.parse(request.options.body).station_no);
  assert.deepEqual(requested, ['508001', '508002', '508003']);
});
