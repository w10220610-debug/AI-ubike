/* Vehicle lookup is user-triggered, bounded and independent of page rendering. */
const WorkOrderLookup = (() => {
  const $ = id => document.getElementById(id);
  let args = {}, emit, resize, generation = 0, controllers = new Set(), matches = [];
  const norm = x => String(x ?? '').normalize('NFKC').trim();
  function parseNumbers(text) {
    const ids = [...new Set(norm(text).split(/[\s,，、;；]+/).filter(Boolean))];
    if (!ids.length || ids.length > 50 || ids.some(x => !/^[0-9]{6,10}$/.test(x)))
      throw Error('請輸入 6–10 位完整數字車號，一次最多 50 台；保留開頭的 0。');
    return ids;
  }
  function records(data) {
    if (data && Object.hasOwn(data, 'retCode') && String(data.retCode) !== '1')
      throw Error('官方服務回報查詢失敗');
    if (Array.isArray(data)) return data;
    if (data && typeof data === 'object') {
      for (const k of ['retVal', 'data', 'items', 'result', 'results']) {
        const v = data[k];
        if (Array.isArray(v)) return v;
        if (v && typeof v === 'object') {
          for (const n of ['items', 'data', 'list', 'results']) if (Array.isArray(v[n])) return v[n];
        }
      }
    }
    throw Error('回應格式無法辨識');
  }
  function selectMatches(data, wanted, spec, queriedAt) {
    const seen = new Set();
    return records(data).filter(x => x && wanted.has(norm(x.bike_no))).flatMap(x => {
      const bike = norm(x.bike_no), pillar = norm(x.pillar_no);
      const key = bike + ':' + pillar;
      if (seen.has(key)) return [];
      seen.add(key);
      const power = x.battery_power == null || norm(x.battery_power) === '' ? NaN : Number(x.battery_power);
      return [{bike_no: bike, station_name: spec.name, pillar_no: pillar,
        battery_power: Number.isFinite(power) && power >= 0 && power <= 100 ? power : null,
        queried_at: queriedAt}];
    });
  }
  function report(message) { $('bike-status').textContent = message; resize(); }
  function stop() {
    generation++;
    for (const c of controllers) c.abort();
    controllers.clear();
    $('bike-search').disabled = !!args.catalog_pending;
    $('bike-cancel').disabled = true;
  }
  function reset() {
    stop(); matches = [];
    $('bike-numbers').value = ''; $('bike-results').replaceChildren();
    $('bike-use').disabled = true; $('bike-status').textContent = '';
  }
  function configure(next) {
    const scope = x => JSON.stringify((x.stations || []).map(s => [s.name, s.station_no]));
    if (scope(args) !== scope(next) && (matches.length || !$('bike-cancel').disabled)) {
      stop(); matches = []; renderMatches();
      report('配置範圍已變更，請重新查詢。');
    }
    args = next;
    $('lookup-scope').textContent = `目前配置範圍：${(args.stations || []).length} 站` +
      (args.catalog_pending ? '；站號清單準備中，稍後即可查詢。' : '') +
      (args.catalog_error ? '；場站清單更新失敗。' : '');
    const prefill = [...new Set((args.prefill_bike_numbers || []).map(norm).filter(x => /^[0-9]{6,10}$/.test(x)))];
    if (prefill.length && !$('bike-numbers').value.trim() && $('bike-cancel').disabled) {
      $('bike-numbers').value = prefill.join('\n');
      $('lookup-panel').open = true;
      report(`已從派工內容抓到 ${prefill.length} 個車號；按「找場站並帶入派工」即可自動回填。`);
    }
    if ($('bike-cancel').disabled) $('bike-search').disabled = !!args.catalog_pending;
    resize();
  }
  function renderMatches() {
    const root = $('bike-results'); root.replaceChildren();
    const counts = new Map();
    for (const m of matches) counts.set(m.bike_no, (counts.get(m.bike_no) || 0) + 1);
    matches.forEach((m, i) => {
      const label = document.createElement('label'), box = document.createElement('input');
      box.type = 'checkbox'; box.dataset.match = String(i); box.checked = counts.get(m.bike_no) === 1;
      label.append(box, document.createTextNode(`${m.bike_no}｜${m.station_name}｜柱號 ${m.pillar_no || '未知'}｜電量 ${m.battery_power == null ? '未提供' : m.battery_power + '%'}｜${m.queried_at}${counts.get(m.bike_no) > 1 ? ' ⚠ 多筆位置，請重新查詢或核對後擇一' : ''}`));
      root.append(label);
    });
    $('bike-use').disabled = !matches.length; resize();
  }
  async function search() {
    let ids;
    try { ids = parseNumbers($('bike-numbers').value); } catch (e) { report(e.message); return; }
    if (args.catalog_pending) { report('場站清單尚未準備完成，請稍後重試。'); return; }
    stop(); const run = generation;
    matches = []; renderMatches(); $('bike-search').disabled = true; $('bike-cancel').disabled = false;
    const wanted = new Set(ids), unique = new Map(), unmapped = [];
    for (const spec of args.stations || []) {
      if (!spec.station_no) unmapped.push(spec.name);
      else if (!unique.has(String(spec.station_no))) unique.set(String(spec.station_no), spec);
    }
    const specs = [...unique.values()], failed = []; let cursor = 0, finished = 0, succeeded = 0, timedOut = false;
    const deadline = setTimeout(() => {
      if (generation !== run) return;
      timedOut = true;
      for (const c of controllers) c.abort();
    }, 60000);
    report(`搜尋中 0/${specs.length} 站…`);
    async function worker() {
      while (generation === run && !timedOut && cursor < specs.length) {
        const spec = specs[cursor++], controller = new AbortController(); controllers.add(controller);
        const timeout = setTimeout(() => controller.abort(), 8000);
        try {
          const response = await fetch('https://apis.youbike.com.tw/api/front/bike/lists?station_no=' + encodeURIComponent(spec.station_no), {
            signal: controller.signal, cache: 'no-store', credentials: 'omit', headers: {'Accept': 'application/json'}
          });
          if (!response.ok) throw Error('HTTP ' + response.status);
          const data = await response.json();
          if (generation !== run || timedOut) return;
          matches.push(...selectMatches(data, wanted, spec, new Date().toLocaleString('zh-TW', {timeZone: 'Asia/Taipei', hour12: false})));
          succeeded++;
        } catch (e) {
          if (generation === run) failed.push(spec.name + (e.name === 'AbortError' ? '（逾時）' : '（查詢失敗）'));
        } finally { clearTimeout(timeout); controllers.delete(controller); finished++; }
        if (generation === run) report(`搜尋中 ${finished}/${specs.length} 站，找到 ${matches.length} 筆…`);
      }
    }
    try { await Promise.all(Array.from({length: Math.min(4, specs.length)}, worker)); }
    finally { clearTimeout(deadline); }
    if (generation !== run) return;
    $('bike-search').disabled = false; $('bike-cancel').disabled = true;
    renderMatches();
    const found = new Set(matches.map(x => x.bike_no)), missing = ids.filter(x => !found.has(x));
    const counts = new Map();
    for (const m of matches) counts.set(m.bike_no, (counts.get(m.bike_no) || 0) + 1);
    const uniqueMatches = matches.filter(m => counts.get(m.bike_no) === 1);
    const autoApply = !timedOut && !failed.length && !unmapped.length && !missing.length &&
      uniqueMatches.length === ids.length;
    if (autoApply) {
      emit({type: 'lookup', matches: uniqueMatches, kind: $('bike-kind').value});
      report(`已找到 ${ids.length} 台，唯一場站已自動帶入派工待確認清單。`);
    } else {
      report(`${timedOut ? '搜尋達 60 秒上限，結果不完整。' : '本次搜尋結束。'}成功查詢 ${succeeded}/${specs.length} 站，找到 ${found.size} 台。` +
        (missing.length ? ` 本次未找到：${missing.join('、')}；不代表車輛不在或已借出。` : '') +
        (failed.length ? ` 失敗 ${failed.length} 站：${failed.join('、')}。` : '') +
        (specs.length > finished ? ` 尚有 ${specs.length - finished} 站未查。` : '') +
        (unmapped.length ? ` 未配對 ${unmapped.length} 站：${unmapped.join('、')}。` : '') +
        ' 有多位置或查詢不完整時不會自動套用，請核對勾選結果後再帶入。');
    }
  }
  function init(send, onResize) {
    emit = send; resize = onResize;
    $('bike-search').onclick = search;
    $('bike-cancel').onclick = () => { stop(); renderMatches(); report('已停止，僅顯示已查到的部分結果；請重新查詢以完成搜尋。'); };
    $('bike-use').onclick = () => {
      const chosen = [...$('bike-results').querySelectorAll('input:checked')].map(x => matches[Number(x.dataset.match)]);
      if (!chosen.length) { report('請先勾選查詢結果。'); return; }
      if (new Set(chosen.map(x => x.bike_no)).size !== chosen.length) { report('同一車號出現多個位置，請核對後只選一筆。'); return; }
      emit({type: 'lookup', matches: chosen, kind: $('bike-kind').value});
      report('已放入下方待確認清單，請核對後按「確認加入勾選派工」。');
    };
  }
  return {init, configure, reset, parseNumbers, records, selectMatches};
})();
