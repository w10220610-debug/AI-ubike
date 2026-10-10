// Run: node tests/battery_browser_syntax.cjs
// No npm dependencies: protects the embedded battery browser script from deployment-time syntax failures.
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'battery_upgrade.py'), 'utf8');
const start = source.indexOf('<script>', source.indexOf("html_text = r'''"));
const end = source.indexOf('</script>', start);
assert.ok(start >= 0 && end > start, 'Battery inline script must be present');
const script = source.slice(start + '<script>'.length, end);
assert.ok(!script.includes(';\\n '), 'Literal \\n between JavaScript statements breaks the whole battery popup');
assert.doesNotThrow(() => new vm.Script(script.replace('__ARGS__', JSON.stringify({
  route_station_map: {D1: [{name: '測試場站', station_no: 'TEST-01'}]},
  catalog_pending: false,
  threshold: 89,
  priority_threshold: 69,
  mobile: true,
})), {filename: 'battery_upgrade.browser.js'}), 'Browser JS must parse');
assert.ok(script.includes('function applyThresholds('), 'Fresh and stale cache must adapt to active thresholds');
assert.ok(script.includes("source:'stale_cache'"), 'Stale battery snapshot fallback must remain');
assert.ok(!script.includes('if(args.catalog_pending){const status='), 'Pending server catalog must not block browser fallback');
assert.ok(script.includes('if(runId!==runtime.run)return;'), 'Old workers must not override a remounted browser page');
console.log('PASS: embedded V29 battery JS syntax, catalog fallback, cache thresholds, stale data, rerun guards');
