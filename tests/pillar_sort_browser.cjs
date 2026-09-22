// Run: node tests/pillar_sort_browser.cjs (Playwright required for this optional test).
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const {chromium} = require(require.resolve('playwright', {
  paths: [process.cwd(), process.env.CODEX_PRIMARY_RUNTIME_NODE_MODULES].filter(Boolean)
}));
const app = fs.readFileSync(path.join(__dirname, '..', 'app.py'), 'utf8');
const shim = app.match(/icon_html = r'''([\s\S]*?)'''/)[1];
const script = shim.match(/<script>([\s\S]*?)<\/script>/)[1];
(async () => {
  const browser = await chromium.launch({headless: true});
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    await page.setContent('<div id="ub-v29-fab"></div><div id="ubike-battery-v29-upgrade"><div class="bike-list"><div class="bike"><span>03</span></div><div class="bike"><span>01</span></div><div class="bike"><span>02</span></div></div></div>');
    await page.evaluate(() => {
      window.sortScans = 0;
      const original = document.querySelectorAll.bind(document);
      document.querySelectorAll = selector => {
        if (selector === '#ubike-battery-v29-upgrade .bike-list') window.sortScans++;
        return original(selector);
      };
      // Exercise the upgrade path from the prior observer.
      window.oldDisconnected = false;
      window.__ubikePillarSortObserver = {disconnect(){window.oldDisconnected = true;}};
    });
    await page.evaluate(script);
    await page.waitForTimeout(650);
    const labels = () => page.locator('.bike span').allTextContents();
    assert.deepEqual(await labels(), ['01', '02', '03']);
    assert.equal(await page.evaluate(() => window.oldDisconnected), true);
    const before = await page.evaluate(() => window.sortScans);
    await page.evaluate(() => {
      for (let i = 0; i < 200; i++) document.body.appendChild(document.createElement('p'));
    });
    await page.waitForTimeout(100);
    assert.equal(await page.evaluate(() => window.sortScans), before);
    await page.evaluate(() => {
      document.querySelector('.bike-list').insertAdjacentHTML('beforeend', '<div class="bike"><span>00</span></div>');
    });
    await page.waitForTimeout(100);
    assert.deepEqual(await labels(), ['00', '01', '02', '03']);
    await page.evaluate(() => {
      document.getElementById('ubike-battery-v29-upgrade').remove();
      document.body.insertAdjacentHTML('beforeend', '<div id="ubike-battery-v29-upgrade"><div class="bike-list"><div class="bike"><span>10</span></div><div class="bike"><span>02</span></div></div></div>');
      window.savedObserver = window.__ubikePillarSortObserver;
    });
    await page.waitForTimeout(100);
    assert.deepEqual(await labels(), ['02', '10']);
    await page.evaluate(script);
    assert.equal(await page.evaluate(() => window.savedObserver === window.__ubikePillarSortObserver), true);
    assert.deepEqual(errors, []);
    console.log('PASS: numeric sort, observer upgrade, 200 unrelated mutations ignored, new rows, root replacement, singleton observer, no JS errors');
  } finally { await browser.close(); }
})().catch(e => { console.error(e); process.exitCode = 1; });
