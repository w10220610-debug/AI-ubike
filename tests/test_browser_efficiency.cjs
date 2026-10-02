const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'app.py'), 'utf8');
const locationHtml = source.match(/DISPATCH_GEOLOCATION_COMPONENT_HTML\s*=\s*r"""([\s\S]*?)"""/)[1];
const locationScript = locationHtml.match(/<script>([\s\S]*?)<\/script>/)[1];
const batteryCore = source.match(/LOW_BATTERY_CLIENT_CORE_JS\s*=\s*r"""([\s\S]*?)"""/)[1];
const batteryHtml = source.match(/def _build_floating_battery_query_html\([\s\S]*?component_html = r"""([\s\S]*?)"""/)[1];

function classList(initial = '') {
  const values = new Set(initial.split(/\s+/).filter(Boolean));
  return {
    contains: value => values.has(value),
    add: (...items) => items.forEach(value => values.add(value)),
    remove: (...items) => items.forEach(value => values.delete(value)),
    toggle(value, enabled = !values.has(value)) {
      if (enabled) values.add(value);
      else values.delete(value);
      return enabled;
    },
  };
}

function locationHarness({ pending = false } = {}) {
  let now = 0;
  let timerId = 0;
  const timers = new Map();
  const listeners = new Map();
  const buttonListeners = new Map();
  const messages = [];
  const requests = [];
  let nextCoords = { latitude: 22.755, longitude: 121.151, accuracy: 20 };
  let nextError = null;
  class ClockDate extends Date {
    constructor(...args) { super(...(args.length ? args : [1727875800000 + now])); }
    static now() { return 1727875800000 + now; }
  }
  const button = { disabled: false, textContent: '', addEventListener: (type, fn) => buttonListeners.set(type, fn) };
  const status = { textContent: '', className: '' };
  const clock = {
    setTimeout(fn, milliseconds = 0) {
      const id = ++timerId;
      timers.set(id, { fn, due: now + Math.max(0, Number(milliseconds)) });
      return id;
    },
    clearTimeout(id) { timers.delete(id); },
  };
  const window = {
    parent: { postMessage: message => messages.push(message) },
    addEventListener(type, fn) { listeners.set(type, fn); },
    ...clock,
  };
  const navigator = { geolocation: {
    getCurrentPosition(success, error, options) {
      requests.push({ success, error, options });
      if (!pending) {
        if (nextError) error(nextError);
        else success({ coords: { ...nextCoords }, timestamp: ClockDate.now() });
      }
    },
  } };
  vm.runInNewContext(locationScript, {
    window, navigator, Date: ClockDate,
    document: { body: { classList: classList(), scrollHeight: 100 }, getElementById: id => id === 'locateButton' ? button : status },
    ...clock,
  }, { filename: 'DISPATCH_GEOLOCATION_COMPONENT_HTML.js' });
  async function advance(milliseconds) {
    const end = now + milliseconds;
    for (;;) {
      const next = [...timers.entries()].filter(([, timer]) => timer.due <= end)
        .sort((a, b) => a[1].due - b[1].due || a[0] - b[0])[0];
      if (!next) break;
      now = next[1].due;
      timers.delete(next[0]);
      next[1].fn();
      await Promise.resolve();
    }
    now = end;
  }
  return {
    render(args = {}) { listeners.get('message')({ data: { type: 'streamlit:render', args: {
      auto_start: true, auto_refresh: true, auto_refresh_seconds: 30, ...args,
    } } }); },
    advance, requests, button, status,
    manual: () => buttonListeners.get('click')(),
    coordinates(coords) { nextCoords = { ...nextCoords, ...coords }; },
    fail(error) { nextError = error; },
    values: () => messages.filter(message => message.type === 'streamlit:setComponentValue').map(message => message.value),
  };
}

test('GPS retains its 30-second cadence through ordinary Streamlit rerenders', async () => {
  const app = locationHarness();
  app.render();
  await app.advance(0);
  for (let i = 0; i < 2; i += 1) {
    await app.advance(10000);
    app.render();
  }
  await app.advance(9999);
  assert.equal(app.requests.length, 1);
  app.render();
  await app.advance(1);
  assert.equal(app.requests.length, 2);
  await app.advance(30000);
  assert.equal(app.requests.length, 3);
  assert.equal(app.values().length, 1, 'stationary automatic updates still suppress page rerenders');
  assert.deepEqual({ ...app.requests[0].options }, { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 });
});

test('GPS auto refresh can be disabled while manual updates still deliver unchanged locations', async () => {
  const app = locationHarness();
  app.render();
  await app.advance(0);
  app.render({ auto_refresh: false });
  await app.advance(60000);
  assert.equal(app.requests.length, 1);
  app.manual();
  assert.equal(app.requests.length, 2);
  assert.equal(app.values().length, 2);
  await app.advance(60000);
  assert.equal(app.requests.length, 2);
});

test('GPS remains single-flight and starts its next timer after completion', async () => {
  const app = locationHarness({ pending: true });
  app.render();
  await app.advance(0);
  app.manual();
  app.render();
  await app.advance(60000);
  assert.equal(app.requests.length, 1);
  assert.equal(app.button.disabled, true);
  app.requests[0].success({ coords: { latitude: 22.755, longitude: 121.151, accuracy: 20 }, timestamp: Date.now() });
  await app.advance(29999);
  assert.equal(app.requests.length, 1);
  await app.advance(1);
  assert.equal(app.requests.length, 2);
});

test('GPS movement, accuracy improvements and five-minute heartbeats continue to deliver', async () => {
  const app = locationHarness();
  app.render();
  await app.advance(0);
  app.coordinates({ latitude: 22.7556, accuracy: 50 });
  await app.advance(30000);
  assert.equal(app.values().length, 2, 'movement of about 67 metres must update routing');
  app.coordinates({ accuracy: 20 });
  await app.advance(30000);
  assert.equal(app.values().length, 3, 'accuracy improvement must still deliver');
  await app.advance(299999);
  assert.equal(app.values().length, 3);
  await app.advance(1);
  assert.equal(app.values().length, 4, 'stationary GPS must retain its five-minute heartbeat');
});

test('GPS errors are reported and automatic retries still run', async () => {
  const app = locationHarness();
  app.fail({ message: '定位暫時失敗' });
  app.render();
  await app.advance(0);
  const [value] = app.values();
  assert.equal(value.ok, false);
  assert.match(value.error, /定位暫時失敗/);
  assert.equal(value.latitude, undefined);
  assert.equal(app.button.disabled, false);
  app.fail(null);
  await app.advance(30000);
  assert.equal(app.requests.length, 2);
  assert.equal(app.values()[1].ok, true);
});

// A minimal DOM exercises the complete embedded battery script. It models the
// controls and owned nodes rather than requiring a browser or making API calls.
function batteryHarness({ fetchImpl, sharedCatalog } = {}) {
  let created = 0;
  const listeners = new Map();
  const storage = new Map([['ubike-battery-query-preferences-v7', JSON.stringify({ sort_mode: 'district' })]]);
  class Element {
    constructor(tagName) {
      this.tagName = tagName;
      this.children = [];
      this.style = {};
      this.dataset = {};
      this.classList = classList();
      this.listeners = new Map();
      this.value = '';
      this.textContent = '';
      this.checked = false;
      created += 1;
    }
    set className(value) { this.classList = classList(value); }
    set innerHTML(html) {
      this.replaceChildren();
      for (const match of html.matchAll(/<([a-z][a-z0-9]*)\b([^>]*)>/gi)) {
        const child = new Element(match[1]);
        const attributes = match[2];
        for (const name of ['class', 'type', 'value']) {
          const value = attributes.match(new RegExp(`\\b${name}="([^"]*)"`));
          if (value) child[name === 'class' ? 'className' : name] = value[1];
        }
        this.appendChild(child);
      }
    }
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
      const result = [];
      const matches = element => selector.startsWith('.')
        ? element.classList.contains(selector.slice(1))
        : selector === 'input:checked' && element.tagName === 'input' && element.checked;
      for (const child of this.children) {
        if (matches(child)) result.push(child);
        result.push(...child.querySelectorAll(selector));
      }
      return result;
    }
    querySelector(selector) { return this.querySelectorAll(selector)[0] || null; }
    click() { return this.listeners.get('click')?.({ preventDefault() {}, stopPropagation() {} }); }
  }
  const doc = {
    head: new Element('head'), body: new Element('body'),
    createElement: tagName => new Element(tagName),
    createTextNode(text) { const element = new Element('#text'); element.textContent = text; return element; },
    getElementById(id) {
      function find(element) {
        if (element.id === id) return element;
        for (const child of element.children) { const result = find(child); if (result) return result; }
        return null;
      }
      return find(doc.head) || find(doc.body);
    },
  };
  doc.body.style.overflow = 'auto';
  const win = {
    document: doc, AbortController,
    matchMedia: () => ({ matches: false }),
    localStorage: { getItem: key => storage.get(key) ?? null, setItem: (key, value) => storage.set(key, String(value)) },
    setTimeout, clearTimeout,
    addEventListener(type, fn) {
      if (!listeners.has(type)) listeners.set(type, new Set());
      listeners.get(type).add(fn);
    },
    removeEventListener(type, fn) { listeners.get(type)?.delete(fn); },
    fetch: fetchImpl || (() => { throw new Error('This UI test must not request network data'); }),
    __ubikeStationCatalog: sharedCatalog,
  };
  const context = vm.createContext({ window: { parent: win }, navigator: {}, Date, Intl, AbortController });
  function render({ fingerprint = 'configured-routes', mode = 'desktop', routes = { D1: [{ name: '臺東縣政府', district: '臺東市' }] } } = {}) {
    const script = batteryHtml.replace('__LOW_BATTERY_CLIENT_CORE__', batteryCore)
      .replace('__ROUTE_STATIONS__', JSON.stringify(routes))
      .replace('__DISPLAY_MODE__', JSON.stringify(mode))
      .replace('__BATTERY_FINGERPRINT__', JSON.stringify(fingerprint))
      .match(/<script>([\s\S]*?)<\/script>/)[1];
    vm.runInContext(script, context, { filename: '_build_floating_battery_query_html.js' });
  }
  return { doc, win, render, created: () => created,
    escape() { for (const listener of listeners.get('keydown') || []) listener({ key: 'Escape' }); },
    keyListenerCount: () => listeners.get('keydown')?.size || 0,
  };
}

test('ordinary rerenders reuse the battery panel, selected scopes and results without new DOM work', () => {
  const app = batteryHarness();
  app.render();
  const page = app.doc.getElementById('ubike-battery-page');
  const fab = app.doc.getElementById('ubike-battery-fab');
  const service = app.win.__ubikeBatteryService;
  const scope = page.querySelector('.battery-scope-list').children[0].children[0];
  scope.checked = true;
  const result = app.doc.createElement('div');
  result.textContent = '既有電量查詢結果';
  page.querySelector('.battery-results').appendChild(result);
  fab.click();
  const created = app.created();
  app.render();
  app.render();
  assert.equal(app.doc.getElementById('ubike-battery-page'), page);
  assert.equal(app.doc.getElementById('ubike-battery-fab'), fab);
  assert.equal(app.created(), created, 'same configuration should allocate no DOM nodes');
  assert.equal(scope.checked, true);
  assert.equal(page.querySelector('.battery-results').children[0], result);
  assert.equal(page.classList.contains('open'), true);
  assert.equal(app.doc.body.style.overflow, 'hidden');
  assert.equal(app.win.__ubikeBatteryService, service);
  assert.equal(app.keyListenerCount(), 1);
  app.escape();
  assert.equal(page.classList.contains('open'), false);
  assert.equal(app.doc.body.style.overflow, 'auto');
});

test('changed routes rebuild an open battery panel and replace only its owned Escape listener', () => {
  const app = batteryHarness();
  app.render();
  const previousPage = app.doc.getElementById('ubike-battery-page');
  const service = app.win.__ubikeBatteryService;
  app.doc.getElementById('ubike-battery-fab').click();
  let otherEscapeCalls = 0;
  app.win.addEventListener('keydown', () => { otherEscapeCalls += 1; });
  app.render({ fingerprint: 'updated-routes', routes: { D2: [{ name: '臺東轉運站' }] } });
  const page = app.doc.getElementById('ubike-battery-page');
  assert.notEqual(page, previousPage);
  assert.equal(previousPage.parent, null);
  assert.equal(page.classList.contains('open'), true);
  assert.equal(app.win.__ubikeBatteryService, service);
  assert.equal(page.querySelector('.battery-scope-list').children[0].children[0].value, 'D2');
  assert.equal(app.keyListenerCount(), 2, 'one panel listener plus the unrelated listener');
  app.escape();
  assert.equal(otherEscapeCalls, 1);
  assert.equal(page.classList.contains('open'), false);
  assert.equal(app.doc.body.style.overflow, 'auto');
});

test('missing battery nodes recover by rebuilding instead of reusing an incomplete panel', () => {
  const app = batteryHarness();
  app.render();
  const oldPage = app.doc.getElementById('ubike-battery-page');
  app.doc.getElementById('ubike-battery-style').remove();
  app.render();
  assert.notEqual(app.doc.getElementById('ubike-battery-page'), oldPage);
  assert.ok(app.doc.getElementById('ubike-battery-style'));
  assert.equal(app.keyListenerCount(), 1);
});

test('battery panel schema marker upgrades a panel with a legacy fingerprint', () => {
  const app = batteryHarness();
  app.render();
  const oldPage = app.doc.getElementById('ubike-battery-page');
  app.win.__ubikeBatteryFingerprint = 'configured-routes';
  app.render();
  assert.notEqual(app.doc.getElementById('ubike-battery-page'), oldPage);
  assert.equal(app.win.__ubikeBatteryFingerprint, 'configured-routes-panel-v2-v27.5.2');
  assert.equal(app.keyListenerCount(), 1);
});

test('battery FAB keeps responsive CSS available on reuse and honours explicit mobile mode', () => {
  const app = batteryHarness();
  app.render();
  const fab = app.doc.getElementById('ubike-battery-fab');
  fab.style.bottom = 'calc(312px + env(safe-area-inset-bottom, 0px))';
  app.render();
  assert.equal(app.doc.getElementById('ubike-battery-fab'), fab);
  assert.equal(fab.style.bottom, '', 'desktop viewport sizes must remain governed by the responsive CSS');
  app.render({ fingerprint: 'mobile-routes', mode: 'mobile' });
  assert.equal(app.doc.getElementById('ubike-battery-fab').style.bottom, 'calc(312px + env(safe-area-inset-bottom, 0px))');
});

test('floating battery queries keep matching stations from the shared synchronization catalogue', async () => {
  const requests = [];
  const app = batteryHarness({
    sharedCatalog: [{ station_id: 'fixture-001', station_name: '臺東縣政府', latitude: 22.75, longitude: 121.14 }],
    fetchImpl: async url => {
      requests.push(url);
      return { ok: true, json: async () => ({ retCode: 1, retVal: [
        { bike_no: 'E1', pillar_no: '7', battery_power: 20 },
      ] }) };
    },
  });
  app.render();
  const page = app.doc.getElementById('ubike-battery-page');
  page.querySelector('.battery-scope-list').children[0].children[0].checked = true;
  await page.querySelector('.battery-refresh').click();
  assert.equal(requests.length, 1, 'shared catalogue must avoid a second catalogue request');
  assert.match(requests[0], /station_no=fixture-001/);
  assert.match(page.querySelector('.battery-status').textContent, /1 個場站完成/);
  assert.doesNotMatch(page.querySelector('.battery-status').textContent, /未配對|失敗/);
});
