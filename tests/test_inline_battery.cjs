const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'app.py'), 'utf8');
const core = source.match(/LOW_BATTERY_CLIENT_CORE_JS\s*=\s*r"""([\s\S]*?)"""/)[1];
const template = source.match(/def _build_inline_low_battery_pillars_html\([\s\S]*?component_html = r"""([\s\S]*?)"""/)[1];
const A = { name: '臺東縣政府', kind: 'candidate', target: 'st-key-candidate_card_select_A', stationNo: '508001' };
const B = { name: '臺東轉運站', kind: 'candidate', target: 'st-key-candidate_card_select_B', stationNo: '508002' };
const C = { name: '臺東市公所', kind: 'candidate', target: 'st-key-candidate_card_select_C', stationNo: '508003' };

function deferred() {
  let resolve;
  const promise = new Promise(done => { resolve = done; });
  return { promise, resolve };
}

function response(bikes) {
  return { ok: true, status: 200, json: async () => ({ retCode: 1, retVal: bikes }) };
}

function harness(batteryResponse = () => response([{ bike_no: 'E1', pillar_no: '7', battery_power: 12 }])) {
  let now = 0;
  let nextTimer = 0;
  const timers = new Map();
  const requests = [];
  class ClockDate extends Date {
    constructor(...args) { super(...(args.length ? args : [1727875800000 + now])); }
    static now() { return 1727875800000 + now; }
  }
  const clock = {
    setTimeout(fn, milliseconds = 0) {
      const id = ++nextTimer;
      timers.set(id, { fn, due: now + Math.max(0, Number(milliseconds)) });
      return id;
    },
    clearTimeout(id) { timers.delete(id); },
  };
  class Element {
    constructor(tagName) {
      this.tagName = tagName;
      this.children = [];
      this.dataset = {};
      this.style = {};
      this.listeners = new Map();
      this.classes = new Set();
      this.ownText = '';
      this.classList = {
        contains: value => this.classes.has(value),
        add: (...values) => values.forEach(value => this.classes.add(value)),
        remove: (...values) => values.forEach(value => this.classes.delete(value)),
        toggle: (value, enabled = !this.classes.has(value)) => {
          if (enabled) this.classes.add(value);
          else this.classes.delete(value);
          return enabled;
        },
      };
    }
    set className(value) { this.classes = new Set(String(value).split(/\s+/).filter(Boolean)); }
    get className() { return [...this.classes].join(' '); }
    set textContent(value) { this.ownText = String(value); this.replaceChildren(); }
    get textContent() { return this.ownText + this.children.map(child => child.textContent).join(''); }
    get isConnected() { return this === doc.body || this === doc.head || Boolean(this.parent?.isConnected); }
    appendChild(child) { child.parent = this; this.children.push(child); return child; }
    append(...children) { children.forEach(child => this.appendChild(child)); }
    replaceChildren(...children) {
      for (const child of this.children) child.parent = null;
      this.children = [];
      this.append(...children);
    }
    remove() {
      if (this.parent) this.parent.children = this.parent.children.filter(child => child !== this);
      this.parent = null;
    }
    setAttribute(name, value) { this[name] = String(value); }
    addEventListener(type, fn) { this.listeners.set(type, fn); }
    querySelectorAll(selector) {
      const matches = element => {
        if (selector.startsWith('.dispatch-plan-card')) return element.classes.has('dispatch-plan-card') && 'ubikeStationName' in element.dataset;
        if (selector.startsWith('.')) return element.classes.has(selector.slice(1));
        if (selector === '[class*="st-key-candidate_card_select_"]') return element.className.includes('st-key-candidate_card_select_');
        return element.tagName === selector;
      };
      const output = [];
      for (const child of this.children) {
        if (matches(child)) output.push(child);
        output.push(...child.querySelectorAll(selector));
      }
      return output;
    }
    querySelector(selector) { return this.querySelectorAll(selector)[0] || null; }
    click() { return this.listeners.get('click')?.({ preventDefault() {}, stopPropagation() {} }); }
  }
  const doc = {
    head: null, body: null,
    createElement: tagName => new Element(tagName),
    createTextNode(text) { const node = new Element('#text'); node.textContent = text; return node; },
    getElementById(id) {
      function find(node) {
        if (node.id === id) return node;
        for (const child of node.children) { const result = find(child); if (result) return result; }
        return null;
      }
      return find(doc.head) || find(doc.body);
    },
    querySelectorAll(selector) { return [...doc.head.querySelectorAll(selector), ...doc.body.querySelectorAll(selector)]; },
    querySelector(selector) { return doc.querySelectorAll(selector)[0] || null; },
  };
  doc.head = new Element('head');
  doc.body = new Element('body');
  const win = {
    document: doc, AbortController, ...clock,
    matchMedia: () => ({ matches: true }),
    fetch: async (url, options = {}) => {
      requests.push({ url, options });
      if (url.includes('station-min-yb2.json')) {
        return { ok: true, status: 200, json: async () => ({ retCode: 1, retVal: { data: [A, B, C].map(spec => ({
          station_no: spec.stationNo, name_tw: spec.name, area_code: '15',
        })) } }) };
      }
      const stationNo = new URL(url).searchParams.get('station_no');
      return batteryResponse(stationNo, requests.filter(request => request.url.includes('/bike/lists')).length, options);
    },
  };
  const context = vm.createContext({ window: { parent: win }, Date: ClockDate, performance: { now: () => now }, AbortController, URL, console, ...clock });
  async function flush() { for (let i = 0; i < 60; i += 1) await Promise.resolve(); }
  async function advance(milliseconds) {
    const end = now + milliseconds;
    await flush();
    for (;;) {
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
  function target(spec) {
    const node = new Element('div');
    node.id = spec.target;
    node.className = spec.target;
    doc.body.appendChild(node);
    return node;
  }
  function render({ specs = [A], threshold = 89, priorityThreshold = 40, forceStation = '', autoQuery = true } = {}) {
    const html = template.replace('__LOW_BATTERY_CLIENT_CORE__', core)
      .replace('__STATION_SPECS__', JSON.stringify(specs))
      .replace('__THRESHOLD__', String(threshold))
      .replace('__PRIORITY_THRESHOLD__', String(priorityThreshold))
      .replace('__AUTO_QUERY__', String(autoQuery))
      .replace('__FORCE_STATION__', JSON.stringify(forceStation))
      .replace('__DISPLAY_MODE__', '"mobile"')
      .replace('__FINGERPRINT__', JSON.stringify(JSON.stringify({ specs, threshold, priorityThreshold, forceStation, autoQuery })));
    const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
    vm.runInContext(script, context, { filename: '_build_inline_low_battery_pillars_html.js' });
  }
  return { win, doc, target, render, advance, flush,
    batteryRequests: () => requests.filter(request => request.url.includes('/bike/lists')),
    pillars: node => node.querySelectorAll('.ubike-pillar'),
    wrapper: node => node.querySelector('.ubike-inline-battery'),
  };
}

test('inline battery fast results hydrate targets that attach after the request completed', async () => {
  const app = harness();
  app.render();
  await app.advance(1);
  const target = app.target(A);
  await app.advance(600);
  assert.match(target.textContent, /07柱/);
  assert.equal(app.batteryRequests().length, 1);
  assert.doesNotMatch(target.textContent, /準備中|正在讀取/);
});

test('inline battery pending requests survive rerenders and forced refresh joins the same station', async () => {
  const pending = deferred();
  const app = harness(() => pending.promise);
  const target = app.target(A);
  app.render();
  await app.advance(500);
  assert.equal(app.batteryRequests().length, 1);
  for (let i = 0; i < 3; i += 1) {
    app.render({ forceStation: A.name });
    await app.advance(500);
  }
  assert.equal(app.batteryRequests().length, 1, 'rerenders and force must share the running official request');
  pending.resolve(response([{ bike_no: 'E1', pillar_no: '7', battery_power: 12 }]));
  await app.advance(0);
  assert.match(target.textContent, /07柱/);
  assert.doesNotMatch(target.textContent, /正在讀取/);
});

test('inline battery fresh cache appears immediately on rerender without another request', async () => {
  const app = harness();
  const oldTarget = app.target(A);
  app.render();
  await app.advance(600);
  assert.match(oldTarget.textContent, /07柱/);
  oldTarget.remove();
  const target = app.target(A);
  app.render();
  await app.advance(0);
  assert.match(target.textContent, /07柱/, 'fresh cached pillars should display before the background queue delay');
  assert.equal(app.batteryRequests().length, 1);
  await app.advance(600);
  assert.equal(app.batteryRequests().length, 1);
});

test('changing station discards old results and stops launching requests from an obsolete queue', async () => {
  const pending = deferred();
  const app = harness(stationNo => stationNo === A.stationNo ? pending.promise : response([{ bike_no: 'E2', pillar_no: '9', battery_power: 20 }]));
  const oldTarget = app.target(A);
  const previousBTarget = app.target(B);
  app.target(C);
  app.render({ specs: [A, B, C] });
  await app.advance(1);
  oldTarget.remove();
  previousBTarget.remove();
  const currentTarget = app.target(B);
  app.render({ specs: [B] });
  await app.advance(600);
  pending.resolve(response([{ bike_no: 'OLD', pillar_no: '7', battery_power: 12 }]));
  await app.advance(300);
  assert.match(currentTarget.textContent, /09柱/);
  assert.doesNotMatch(currentTarget.textContent, /07柱/);
  assert.equal(app.batteryRequests().filter(request => request.url.includes(`station_no=${C.stationNo}`)).length, 0,
    'the superseded mobile worker queue must not launch its third station');
});

test('inline battery hung requests become retryable and retry can return valid pillars', async () => {
  const pending = deferred();
  let hung = true;
  const app = harness(() => hung ? pending.promise : response([{ bike_no: 'E1', pillar_no: '7', battery_power: 12 }]));
  const target = app.target(A);
  app.render();
  await app.advance(20000);
  assert.match(target.textContent, /暫未取得|失敗|逾時/);
  const retry = target.querySelector('button');
  assert.ok(retry, 'a timeout must leave a retry action');
  const timedOutRequestCount = app.batteryRequests().length;
  hung = false;
  retry.click();
  await app.advance(0);
  assert.match(target.textContent, /07柱/);
  assert.equal(app.batteryRequests().length, timedOutRequestCount + 1);
});

test('inline battery keeps inclusive thresholds, deduplicated sorted pillars and urgent markings', async () => {
  const app = harness(() => response([
    { bike_no: 'E1', pillar_no: '10', battery_power: 89 },
    { bike_no: 'E2', pillar_no: '2', battery_power: 40 },
    { bike_no: 'E3', pillar_no: '1', battery_power: 41 },
    { bike_no: 'E4', pillar_no: '2', battery_power: 20 },
    { bike_no: 'E5', pillar_no: '3', battery_power: 90 },
  ]));
  const target = app.target(A);
  app.render({ threshold: 89, priorityThreshold: 40 });
  await app.advance(600);
  const pillars = app.pillars(target);
  assert.deepEqual(pillars.map(pillar => pillar.textContent), ['01柱', '⚠ 02柱', '10柱']);
  assert.deepEqual(pillars.map(pillar => pillar.classList.contains('urgent')), [false, true, false]);
});

test('legacy specs retain manual battery queries and resolve their station through the catalog', async () => {
  const app = harness();
  const target = app.target(A);
  const { stationNo, ...legacySpec } = A;
  app.render({ specs: [legacySpec], autoQuery: false });
  await app.advance(600);
  assert.equal(app.batteryRequests().length, 0);
  assert.match(target.textContent, /查低電量柱號/);
  target.querySelector('button').click();
  await app.advance(0);
  assert.match(target.textContent, /07柱/);
  assert.equal(app.batteryRequests().length, 1);
  assert.match(app.batteryRequests()[0].url, /station_no=508001/);
});

test('rerenders join a forced refresh instead of displaying the previous cached pillars', async () => {
  const pending = deferred();
  const app = harness((_stationNo, requestCount) => requestCount === 1
    ? response([{ bike_no: 'BEFORE', pillar_no: '3', battery_power: 12 }])
    : pending.promise);
  const target = app.target(A);
  app.render();
  await app.advance(600);
  assert.match(target.textContent, /03柱/);
  target.querySelector('button').click();
  await app.advance(0);
  assert.equal(app.batteryRequests().length, 2);
  app.render();
  await app.advance(0);
  assert.doesNotMatch(target.textContent, /03柱/, 'an in-progress refresh supersedes the completed cache');
  assert.match(target.textContent, /正在讀取/);
  pending.resolve(response([{ bike_no: 'AFTER', pillar_no: '9', battery_power: 12 }]));
  await app.advance(0);
  assert.match(target.textContent, /09柱/);
  assert.doesNotMatch(target.textContent, /03柱|正在讀取/);
  assert.equal(app.batteryRequests().length, 2, 'the replacement UI must join the forced request');
});
