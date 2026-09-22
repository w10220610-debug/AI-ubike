// npm install jsdom; node tests/pillar_sort_dom.cjs
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const {JSDOM} = require('jsdom');
const app = fs.readFileSync(path.join(__dirname, '..', 'app.py'), 'utf8');
const script = app.match(/icon_html = r'''([\s\S]*?)'''/)[1].match(/<script>([\s\S]*?)<\/script>/)[1];
const row = n => `<div class="bike"><span>${n}</span></div>`;
const root = rows => `<div id="ubike-battery-v29-upgrade"><div class="bike-list">${rows.map(row).join('')}</div></div>`;
(async () => {
  const dom = new JSDOM(`<div id="ub-v29-fab"></div>${root(['03', '01', '02'])}`, {runScripts:'outside-only'});
  const win = dom.window, doc = win.document;
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const labels = () => Array.from(doc.querySelectorAll('.bike span'), node => node.textContent);
  let scans = 0, disconnected = false;
  const query = doc.querySelectorAll.bind(doc);
  doc.querySelectorAll = selector => {
    if (selector === '#ubike-battery-v29-upgrade .bike-list') scans++;
    return query(selector);
  };
  win.__ubikePillarSortObserver = {disconnect(){disconnected = true;}};
  try {
    win.eval(script);
    await wait(650);
    assert.deepEqual(labels(), ['01', '02', '03']);
    assert.ok(disconnected);
    const before = scans;
    for(let i=0;i<200;i++) doc.body.appendChild(doc.createElement('p'));
    await wait(100);
    assert.equal(scans, before, 'unrelated page edits must not trigger a list scan');
    doc.querySelector('.bike-list').insertAdjacentHTML('beforeend', row('00'));
    await wait(100);
    assert.deepEqual(labels(), ['00', '01', '02', '03']);
    doc.querySelector('.bike span').textContent = '09';
    await wait(100);
    assert.deepEqual(labels(), ['01', '02', '03', '09']);
    doc.getElementById('ubike-battery-v29-upgrade').remove();
    doc.body.insertAdjacentHTML('beforeend', root(['10', '02']));
    await wait(100);
    assert.deepEqual(labels(), ['02', '10']);
    const observer = win.__ubikePillarSortObserver;
    win.eval(script);
    assert.equal(win.__ubikePillarSortObserver, observer);
    console.log('PASS: DOM numeric sorting; old observer cleanup; 200 unrelated mutations ignored; inserted row; changed label; root remount; singleton after rerun');
  } finally { win.close(); }
})().catch(error => {console.error(error); process.exitCode = 1;});
