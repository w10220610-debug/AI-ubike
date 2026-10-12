// Exercise the actual shared battery service without network access or a DOM.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'app.py'), 'utf8');
const core = source.match(/LOW_BATTERY_CLIENT_CORE_JS\s*=\s*r"""([\s\S]*?)"""/)[1];
const stationNo = '501501003';
const stationName = '臺東市公所';
const batteryPayload = (records = [{ bike_no: 'E001', pillar_no: '03', battery_power: '39' }]) => ({
  retCode: 1, retMsg: '查詢車輛資訊成功', retVal: records,
});
const catalogPayload = [{ station_no: stationNo, name_tw: stationName, area_code: '15' }];
const response = payload => ({ ok: true, status: 200, json: async () => payload });
const plain = value => JSON.parse(JSON.stringify(value));
const flush = async () => { for (let i = 0; i < 30; i += 1) await Promise.resolve(); };
function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((success, failure) => { resolve = success; reject = failure; });
  return { promise, resolve, reject };
}

function harness(fetcher = url => response(url.includes('station-min') ? catalogPayload : batteryPayload()), sharedCatalog) {
  let now = 100000;
  let nextTimer = 0;
  const timers = new Map();
  const requests = [];
  class ClockDate extends Date { static now() { return now; } }
  const win = {
    AbortController,
    __ubikeStationCatalog: sharedCatalog,
    setTimeout(fn, milliseconds = 0) {
      const id = ++nextTimer;
      timers.set(id, { fn, due: now + Math.max(0, Number(milliseconds)) });
      return id;
    },
    clearTimeout(id) { timers.delete(id); },
    async fetch(url, options) {
      requests.push({ url, options });
      return fetcher(url, options, requests.length);
    },
  };
  const context = vm.createContext({ Date: ClockDate, AbortController, TypeError });
  vm.runInContext(core, context, { filename: 'LOW_BATTERY_CLIENT_CORE_JS.js' });
  const service = context.ensureUbikeBatteryService(win);
  async function advance(milliseconds) {
    await flush();
    const end = now + milliseconds;
    for (;;) {
      const next = [...timers.entries()].filter(([, timer]) => timer.due <= end)
        .sort((left, right) => left[1].due - right[1].due || left[0] - right[0])[0];
      if (!next) break;
      now = next[1].due;
      timers.delete(next[0]);
      next[1].fn();
      await flush();
    }
    now = end;
    await flush();
  }
  return { service, requests, advance, win, context };
}

test('healthy station lookup preserves battery values and exposes a reusable cached result', async () => {
  const app = harness();
  const result = await app.service.queryStationByName(stationName);
  assert.equal(result.matched, true);
  assert.equal(result.stationNo, stationNo);
  assert.deepEqual(plain(result.bikes), [{ bike_no: 'E001', pillar_no: '03', battery_power: 39 }]);
  assert.equal(app.requests.length, 2, 'one catalogue and one station query');
  const cached = app.service.getCachedStationResult('台東市公所');
  assert.equal(cached.stationNo, stationNo);
  assert.deepEqual(plain(cached.bikes), plain(result.bikes));
  await app.service.queryStationByName(stationName);
  assert.equal(app.requests.length, 2);
});

test('known station ID bypasses the large catalogue without altering the response shape', async () => {
  const app = harness(url => {
    assert.ok(url.includes('bike/lists'), 'direct station ID must not fetch the catalogue');
    assert.ok(url.includes(`station_no=${stationNo}`));
    return response(batteryPayload());
  });
  const result = await app.service.queryStationByName(stationName, { stationNo });
  assert.equal(result.stationNo, stationNo);
  assert.equal(result.bikes.length, 1);
  assert.equal(app.requests.length, 1);
  assert.equal(app.service.getCachedStationResult(stationName, { stationNo }).bikes.length, 1);
});

test('successful empty battery list remains a cacheable empty result', async () => {
  const app = harness(() => response(batteryPayload([])));
  const result = await app.service.queryStationByName(stationName, { stationNo });
  assert.deepEqual(plain(result.bikes), []);
  const cached = app.service.getCachedStationResult(stationName, { stationNo });
  assert.ok(cached, 'successful empty lists must differ from unavailable data');
  assert.deepEqual(plain(cached.bikes), []);
});

test('HTTP 200 application errors reject and cannot be cached as an empty success', async () => {
  const app = harness(() => response({ retCode: 0, retMsg: '暫時無法查詢', retVal: [] }));
  await assert.rejects(app.service.queryStationByName(stationName, { stationNo, attempts: 1 }));
  assert.equal(app.service.getCachedStationResult(stationName, { stationNo }), null);
});

test('malformed battery data rejects rather than displaying a false empty result', async () => {
  for (const retVal of [undefined, null, {}, { data: [] }, 'not a list']) {
    const app = harness(() => response({ retCode: 1, retVal }));
    await assert.rejects(app.service.getBatteryListByStationNo(stationNo, { attempts: 1 }));
    assert.equal(app.service.getCachedStationResult(stationName, { stationNo }), null);
  }
});

test('HTTP failures reject and leave no false cached result', async () => {
  const app = harness(() => ({ ok: false, status: 503, json: async () => batteryPayload([]) }));
  await assert.rejects(app.service.getBatteryListByStationNo(stationNo, { attempts: 1 }));
  assert.equal(app.service.getCachedStationResult(stationName, { stationNo }), null);
});

test('missing battery values are excluded while actual zero and valid numeric text are retained', async () => {
  const values = [null, undefined, '', '   ', 'not a number', 0, '0', 40.9, -2, 120];
  const app = harness(() => response(batteryPayload(values.map((battery_power, index) => ({
    bike_no: `E${index}`, pillar_no: String(index), battery_power,
  })))));
  const result = await app.service.getBatteryListByStationNo(stationNo);
  assert.deepEqual(plain(result.map(bike => bike.battery_power)), [0, 0, 40, 0, 100]);
});

test('simultaneous forced refreshes join one station request', async () => {
  const pending = deferred();
  const app = harness(() => pending.promise);
  const first = app.service.queryStationByName(stationName, { stationNo, force: true });
  const second = app.service.queryStationByName(stationName, { stationNo, force: true });
  await flush();
  assert.equal(app.requests.length, 1);
  pending.resolve(response(batteryPayload()));
  const results = await Promise.all([first, second]);
  assert.equal(app.requests.length, 1, 'joining force must not issue another request after the first completes');
  assert.deepEqual(plain(results[0].bikes), plain(results[1].bikes));
});

test('hard deadline settles a fetch that ignores abort and permits a later recovery', async () => {
  let recover = false;
  const app = harness(() => recover ? response(batteryPayload()) : new Promise(() => {}));
  let outcome = 'pending';
  const waiting = app.service.getBatteryListByStationNo(stationNo, { attempts: 3, timeoutMs: 5000 })
    .then(() => { outcome = 'success'; }, () => { outcome = 'failure'; });
  await app.advance(18000);
  assert.equal(outcome, 'failure', 'total waiting must end at the default 18-second hard deadline');
  await waiting;
  assert.equal(app.service.getCachedStationResult(stationName, { stationNo }), null);
  recover = true;
  const result = await app.service.getBatteryListByStationNo(stationNo);
  assert.equal(result.length, 1);
});

test('hard deadline also includes response JSON parsing', async () => {
  const app = harness(() => ({ ok: true, status: 200, json: () => new Promise(() => {}) }));
  let rejected = false;
  const waiting = app.service.getBatteryListByStationNo(stationNo, { attempts: 1 })
    .then(() => {}, () => { rejected = true; });
  await app.advance(18000);
  assert.equal(rejected, true, 'body parsing must not keep the station card loading forever');
  await waiting;
  assert.equal(app.service.getCachedStationResult(stationName, { stationNo }), null);
});

test('catalogue and battery body share one total deadline and late data cannot populate the cache', async () => {
  const app = harness(url => {
    if (url.includes('station-min')) {
      return new Promise(resolve => app.win.setTimeout(() => resolve(response(catalogPayload)), 6000));
    }
    return {
      ok: true, status: 200,
      json: () => new Promise(resolve => app.win.setTimeout(() => resolve(batteryPayload()), 10000)),
    };
  });
  let outcome = 'pending';
  const waiting = app.service.queryStationByName(stationName, {
    attempts: 3, timeoutMs: 7000, totalTimeoutMs: 9000,
  }).then(() => { outcome = 'success'; }, () => { outcome = 'failure'; });
  await app.advance(6000);
  assert.equal(app.requests.length, 2, 'battery query starts after the slow catalogue succeeds');
  assert.equal(outcome, 'pending');
  await app.advance(3000);
  assert.equal(outcome, 'failure', 'battery must use the remaining three seconds instead of a fresh budget');
  await waiting;
  assert.equal(app.service.getCachedStationResult(stationName, { stationNo }), null);
  await app.advance(8000);
  assert.equal(app.service.getCachedStationResult(stationName, { stationNo }), null,
    'a late body from the timed-out request must never be cached');
});

test('a service error clears its pending slot so the next request can recover', async () => {
  let failed = true;
  const app = harness(() => response(failed ? { retCode: -1, retVal: [] } : batteryPayload()));
  await assert.rejects(app.service.getBatteryListByStationNo(stationNo, { attempts: 1 }));
  failed = false;
  const result = await app.service.queryStationByName(stationName, { stationNo });
  assert.equal(result.bikes.length, 1);
  assert.equal(app.requests.length, 2);
});

test('expired cached results are unavailable and normal querying fetches fresh data', async () => {
  const app = harness();
  await app.service.queryStationByName(stationName, { stationNo, ttlMs: 30000 });
  await app.advance(29999);
  assert.ok(app.service.getCachedStationResult(stationName, { stationNo, ttlMs: 30000 }));
  await app.advance(2);
  assert.equal(app.service.getCachedStationResult(stationName, { stationNo, ttlMs: 30000 }), null);
  await app.service.queryStationByName(stationName, { stationNo, ttlMs: 30000 });
  assert.equal(app.requests.length, 2);
});

test('shared minimal synchronization catalogue supplies station IDs without another catalogue download', async () => {
  const app = harness(url => {
    assert.ok(url.includes('bike/lists'));
    return response(batteryPayload());
  }, [{ station_id: stationNo, station_name: stationName }]);
  const result = await app.service.queryStationByName('台東市公所');
  assert.equal(result.stationNo, stationNo);
  assert.equal(result.bikes.length, 1);
  assert.equal(app.requests.length, 1);
  await app.service.queryStationByName('台東市公所', { force: true });
  assert.equal(app.requests.length, 2, 'battery refresh must continue reusing the shared catalogue');
});

test('the new service version replaces an existing older service object', () => {
  const app = harness();
  const old = { version: 'v27.5.1' };
  app.win.__ubikeBatteryService = old;
  const upgraded = app.context.ensureUbikeBatteryService(app.win);
  assert.notEqual(upgraded, old);
  assert.equal(upgraded.version, 'v27.5.2');
  assert.equal(app.context.ensureUbikeBatteryService(app.win), upgraded);
});
