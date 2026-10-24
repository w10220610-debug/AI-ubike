"""Work orders: explicit review, separate completion, cross-shift persistence."""
from __future__ import annotations
import hashlib
import json
import re
import time
import unicodedata
from pathlib import Path

KEY = '__work_orders_v1__'
DRAFT_SCHEMA = 'vehicle-location-v3'

def is_station_code(value):
    return bool(re.fullmatch(r'50\d{7}', str(value or '').strip()))

def is_bike_number(value):
    value = str(value or '').strip()
    return bool(re.fullmatch(r'[0-9]{6,10}', value)) and not is_station_code(value)

FIELDS = ('station_name', 'kind', 'equipment', 'issued_at', 'description')
LABELS = dict(zip(FIELDS, ('場站', '派工類型', '設備編號', '派工時間', '問題描述')))


def station_key(value):
    return ''.join(unicodedata.normalize('NFKC', str(value or '')).replace('臺', '台').split()).casefold()


def orders(cache):
    return cache.setdefault('metadata', {}).setdefault(KEY, [])


def pending(cache, station=None):
    return [x for x in orders(cache) if x.get('state') not in ('done', 'deleted') and
            (station is None or station_key(x.get('station_name')) == station_key(station))]


def pending_names(cache):
    return list(dict.fromkeys(x['station_name'] for x in pending(cache)))


def identity(row):
    # A timestamp distinguishes later assignments to the same equipment.
    parts = [station_key(row.get(k)) for k in FIELDS[:4]]
    if not row.get('equipment'):
        parts.append(station_key(row.get('description')))
    return hashlib.sha256('|'.join(parts).encode()).hexdigest()[:24]


def merge(cache, incoming):
    target = orders(cache)
    by_id = {x['id']: x for x in target}
    changed = 0
    for raw in incoming[:2000]:
        if not isinstance(raw, dict) or not raw.get('station_name'):
            continue
        row = {k: str(raw.get(k) or '')[:2000] for k in FIELDS}
        row['id'] = identity(row)
        row['state'] = raw.get('state') if raw.get('state') in ('pending', 'blocked', 'done', 'deleted') else 'pending'
        try:
            row['updated_at'] = min(float(raw.get('updated_at') or 0), time.time() + 60)
        except (TypeError, ValueError):
            continue
        old = by_id.get(row['id'])
        if old is None:
            target.append(row)
            by_id[row['id']] = row
            changed += 1
        elif row['updated_at'] > old.get('updated_at', 0):
            old.update(row)
            changed += 1
    return changed


def parse_text(text, station_names):
    """Keep uncertain rows for review; never silently fuzzy-match a station."""
    names = {}
    for name in station_names:
        names.setdefault(station_key(name), []).append(name)
    results = []
    issued = ''
    date_re = re.compile(r'20\d{2}[./-]\d{1,2}[./-]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?')
    for raw in str(text).splitlines():
        line = unicodedata.normalize('NFKC', raw).strip()
        if not line:
            continue
        if '車輛號碼' in line and ('場站' in line or '故障資訊' in line):
            continue
        stamp = date_re.search(line)
        if stamp:
            issued = stamp.group().replace('.', '-').replace('/', '-')
            line = line[:stamp.start()] + line[stamp.end():]
        compact = station_key(line)
        if not compact:
            continue
        matches = [k for k in names if k and k in compact]
        # Prefer a longer exact station name over its substring.
        matches = [k for k in matches if not any(k != other and k in other for other in matches)]
        kind = re.search(r'車\s*輛\s*電\s*池|換\s*電(?:\s*池)?|車\s*輛\s*故\s*障|故\s*障|補\s*車|收\s*車|綁\s*車|清\s*潔|巡\s*檢|電\s*池', line)
        # YouBike station codes (50 + seven digits) are not vehicle numbers.
        equipment_ids = [x for x in re.findall(r'(?<!\d)\d{6,10}(?!\d)', line)
                         if not re.fullmatch(r'50\d{7}', x)]
        equipment = equipment_ids[0] if equipment_ids else ''
        if not matches and not kind and not equipment:
            if results and line not in ('派工', '派工列表'):
                results[-1]['description'] += ' ' + line
            continue
        name = names[matches[0]][0] if len(matches) == 1 and len(names[matches[0]]) == 1 else ''
        guessed = line[:kind.start()].strip(' |｜') if kind else ''
        # Station-name continuation lines belong to the printed reference only.
        if name and not kind and not equipment and results:
            results[-1]['description'] += ' ' + line
            continue
        for bike in equipment_ids or ['']:
            results.append(dict(station_name='' if bike else name,
                                kind=re.sub(r'\s+', '', kind.group()) if kind else '',
                                equipment=bike, issued_at=issued, description=line,
                                detected_name=guessed, source_station=name,
                                selected=bool(name) and not bike))
        issued = ""
    return results


def badge_html(cache, station):
    import html
    tasks = pending(cache, station)
    if not tasks:
        return ''
    details = '；'.join(f"{x['kind']} {x['equipment']}" for x in tasks)
    return f'<small style="display:block;color:#a855f7;font-weight:800">🟣 派工 {len(tasks)} 件｜{html.escape(details)}</small>'


def clear_draft(state, prefix):
    """Reset widget state in a callback, before widgets are instantiated."""
    for key in list(state):
        if key in {prefix + '::text', prefix + '::draft', prefix + '::draft_version',
                   prefix + '::notice', prefix + '::lookup_request', prefix + '::legacy_draft'} or key.startswith(prefix + '::editor::'):
            del state[key]
    state[prefix + '::reset_epoch'] = int(state.get(prefix + '::reset_epoch', 0)) + 1


def lookup_draft(matches, names, kind):
    """Accept only configured stations and exact numeric equipment IDs."""
    output = []
    seen = set()
    if kind not in ('換電', '收車', '車輛故障', '巡檢'):
        return output
    for row in matches[:100] if isinstance(matches, list) else []:
        if not isinstance(row, dict):
            continue
        name, bike = str(row.get('station_name', '')), str(row.get('bike_no', ''))
        if name not in names or not is_bike_number(bike):
            continue
        key = (name, bike)
        if key in seen:
            continue
        seen.add(key)
        description = f"車號查詢：柱號 {str(row.get('pillar_no') or '未知')[:20]}；查詢時間 {str(row.get('queried_at') or '')[:40]}"
        output.append(dict(station_name=name, equipment=bike, kind=kind, issued_at='',
                           description=description, selected=True, location_confirmed=True))
    return output


def apply_lookup_to_draft(draft, matches, names, kind):
    """Fill current parsed rows by exact bike number; append vehicle-only rows."""
    base = [dict(row) for row in draft] if isinstance(draft, list) else []
    found_rows = lookup_draft(matches, names, kind)
    by_bike = {}
    for index, row in enumerate(base):
        bike = str(row.get('equipment') or '').strip()
        if re.fullmatch(r'[0-9]{6,10}', bike):
            by_bike.setdefault(bike, []).append(index)

    for found in found_rows:
        bike = found['equipment']
        indexes = by_bike.get(bike, [])
        if indexes:
            for index in indexes:
                row = base[index]
                row['station_name'] = found['station_name']
                row['location_confirmed'] = True
                if not str(row.get('kind') or '').strip():
                    row['kind'] = found['kind']
                detail = found['description']
                current = str(row.get('description') or '').strip()
                if detail and detail not in current:
                    row['description'] = (current + '；' + detail).strip('；')
                row['selected'] = True
        else:
            base.append(found)
            by_bike[bike] = [len(base) - 1]
    return base


def prepare_vehicle_draft(state, prefix, names):
    """Migrate transient old OCR results, never mutate saved work orders."""
    if state.get(prefix + '::draft_schema') == DRAFT_SCHEMA:
        return
    state[prefix + '::draft_schema'] = DRAFT_SCHEMA
    draft = state.get(prefix + '::draft')
    if not isinstance(draft, list) or not draft:
        return
    state[prefix + '::legacy_draft'] = draft
    if any(is_station_code(row.get('equipment')) for row in draft):
        original = state.get(prefix + '::text') or '\n'.join(str(row.get('description') or '') for row in draft)
        draft = parse_text(original, names)
    else:
        draft = [dict(row) for row in draft]
    for row in draft:
        if is_bike_number(row.get('equipment')) and not row.get('location_confirmed'):
            row.update(station_name='', selected=False)
    state[prefix + '::draft'] = draft
    state[prefix + '::draft_version'] = str(time.time_ns())
    state[prefix + '::lookup_request'] = state[prefix + '::draft_version']
    state[prefix + '::notice'] = '已更新舊辨識清單，正在依車號自動查場站；不需要手動選場站。'


def lookup_station_args(route_station_map):
    # Reuse the battery map's bounded background job; never wait for network in the UI.
    from battery_upgrade import _clean_route_map, _load_station_map, _monitor_station_map
    from background_refresh import station_map_refresh
    clean = _clean_route_map(route_station_map)
    scope = tuple((zone, tuple((x['name'], x['district']) for x in items))
                  for zone, items in sorted(clean.items()))
    state = station_map_refresh.poll(scope, lambda: _load_station_map(clean))
    if state.pending and state.value is None:
        _monitor_station_map(scope, clean, 'work_order_station_map::' + hashlib.sha256(repr(scope).encode()).hexdigest())
    return dict(stations=[x for items in (state.value or clean).values() for x in items],
                catalog_pending=state.value is None,
                catalog_error=state.error or '')


def render_manager(status_df, *, cache, token, page_mode, save, route_station_map=None):
    import pandas as pd
    import streamlit as st
    import streamlit.components.v1 as components
    names = sorted(set(status_df['場站名稱'].dropna().astype(str)))
    prefix = f'work_orders::{token}'
    prepare_vehicle_draft(st.session_state, prefix, names)
    component = components.declare_component('work_order_import', path=str(Path(__file__).with_name('work_order_component')))
    with st.expander(f'🟣 派工待辦｜{len(pending(cache))} 件未完成', expanded=False):
        st.caption('車號自動定位 v3｜照片 → 辨識車號 → 自動填入目前場站。場站由系統查詢，不需要手動選擇。')
        epoch = st.session_state.get(prefix + '::reset_epoch', 0)
        current_draft = st.session_state.get(prefix + '::draft')
        prefill_bike_numbers = []
        if isinstance(current_draft, list):
            prefill_bike_numbers = list(dict.fromkeys(
                str(row.get('equipment') or '').strip()
                for row in current_draft
                if isinstance(row, dict)
                and is_bike_number(row.get('equipment'))
            ))
        event = component(token=token, snapshot=orders(cache), reset_epoch=epoch,
                          prefill_bike_numbers=prefill_bike_numbers,
                          lookup_request=st.session_state.get(prefix + '::lookup_request', ''),
                          **lookup_station_args(route_station_map or {}),
                          key=prefix + '::ocr::' + DRAFT_SCHEMA, default=None)
        if isinstance(event, dict) and event.get('event_id') != st.session_state.get(prefix + '::event'):
            st.session_state[prefix + '::event'] = event.get('event_id')
            if event.get('type') == 'restore' and isinstance(event.get('orders'), list):
                if merge(cache, event['orders']):
                    save()
                    st.rerun()
            elif (event.get('type') == 'lookup' and event.get('reset_epoch') == epoch
                  and event.get('request_id', '') == st.session_state.get(prefix + '::lookup_request', '')):
                st.session_state[prefix + '::draft'] = apply_lookup_to_draft(
                    st.session_state.get(prefix + '::draft'),
                    event.get('matches'),
                    names,
                    event.get('kind'),
                )
                st.session_state[prefix + '::draft_version'] = str(event.get('event_id'))
            elif event.get('type') == 'ocr' and event.get('reset_epoch') == epoch:
                st.session_state[prefix + '::text'] = str(event.get('text') or '')[:100000]
                st.session_state[prefix + '::draft'] = parse_text(event.get('text', ''), names)
                st.session_state[prefix + '::draft_version'] = str(event.get('event_id'))
                st.session_state[prefix + '::lookup_request'] = str(event.get('event_id'))
                st.rerun()
        text = st.text_area('辨識文字（也可貼上派工文字）', key=prefix + '::text', height=100)
        if st.button('整理為待確認清單', key=prefix + '::parse'):
            st.session_state[prefix + '::draft'] = parse_text(text, names)
            st.session_state[prefix + '::draft_version'] = str(time.time_ns())
            st.session_state[prefix + '::lookup_request'] = st.session_state[prefix + '::draft_version']
            st.rerun()
        st.button('🧹 清除輸入／重新開始', key=prefix + '::clear',
                  on_click=clear_draft, args=(st.session_state, prefix),
                  help='清除照片、辨識文字、待確認清單及車號搜尋結果；已加入的派工保留。')
        draft = st.session_state.get(prefix + '::draft')
        if draft:
            vehicle_mode = any(is_bike_number(x.get('equipment')) for x in draft)
            ready = sum(bool(x.get('location_confirmed')) for x in draft if is_bike_number(x.get('equipment')))
            if vehicle_mode:
                st.info(f'已自動定位 {ready} 台；其餘待查。場站由系統填入，無須手動選擇；查不到不會使用照片舊場站。')
                if st.button('🔄 重新自動查場站', key=prefix + '::retry_location'):
                    for row in draft:
                        if is_bike_number(row.get('equipment')):
                            row.update(station_name='', selected=False, location_confirmed=False)
                    st.session_state[prefix + '::lookup_request'] = str(time.time_ns())
                    st.session_state[prefix + '::draft_version'] = st.session_state[prefix + '::lookup_request']
                    st.rerun()
                st.caption('車號辨識有誤可在上方修正辨識文字，再按「整理為待確認清單」自動重查。')
            rows = [{'加入': x['selected'], **{LABELS[k]: x[k] for k in FIELDS}} for x in draft]
            edited = st.data_editor(pd.DataFrame(rows), hide_index=True, num_rows='fixed' if vehicle_mode else 'dynamic',
                disabled=['場站', '設備編號'] if vehicle_mode else [],
                column_config={'場站': (st.column_config.TextColumn('目前場站（自動查詢）') if vehicle_mode
                                      else st.column_config.SelectboxColumn(options=names, required=True)),
                               '設備編號': st.column_config.TextColumn(), '加入': st.column_config.CheckboxColumn()},
                key=prefix + '::editor::' + st.session_state.get(prefix + '::draft_version', '0'))
            if st.button('確認加入勾選派工', type='primary', key=prefix + '::add'):
                selected = edited[edited['加入'].fillna(False)]
                if selected.empty:
                    st.warning('請至少勾選一筆。')
                elif any(is_station_code(x) for x in selected['設備編號']):
                    st.error('辨識內容含場站編號，不能當車號加入。請按「整理為待確認清單」重新解析。')
                elif any(is_bike_number(row['設備編號']) and not any(
                    x.get('location_confirmed') and x.get('equipment') == row['設備編號']
                    and x.get('station_name') == row['場站'] for x in draft
                ) for row in selected.to_dict('records')):
                    st.warning('勾選項目尚未查到目前場站，請等待自動查詢或按「重新自動查場站」。不需要手動選站。')
                elif any(str(x) not in names for x in selected['場站']):
                    st.error('尚未確認場站，請重新查詢後再加入。')
                elif selected['派工類型'].fillna('').str.strip().eq('').any():
                    st.error('請填寫每筆勾選項目的派工類型。')
                else:
                    existing = {x['id']: x for x in orders(cache)}
                    added = 0
                    for row in selected.fillna('').to_dict('records'):
                        item = {k: str(row[LABELS[k]]).strip() for k in FIELDS}
                        item_id = identity(item)
                        if item_id not in existing or existing[item_id].get('state') == 'deleted':
                            item.update(id=item_id, state='pending', updated_at=time.time())
                            if item_id in existing:
                                existing[item_id].update(item)
                            else:
                                orders(cache).append(item)
                                existing[item_id] = item
                            added += 1
                    save()
                    st.session_state[prefix + '::notice'] = f'已新增 {added} 筆，略過 {len(selected)-added} 筆重複派工。'
                    remaining = edited[~edited['加入'].fillna(False).astype(bool)]
                    if not remaining.empty:
                        st.session_state[prefix + '::draft'] = [
                            dict(selected=False, location_confirmed=any(
                                x.get('location_confirmed') and x.get('equipment') == row['設備編號']
                                and x.get('station_name') == row['場站'] for x in draft),
                                **{k: str(row[LABELS[k]]) for k in FIELDS})
                            for row in remaining.fillna('').to_dict('records')
                        ]
                        st.session_state[prefix + '::draft_version'] = str(time.time_ns())
                    else:
                        st.session_state.pop(prefix + '::draft', None)
                        st.session_state.pop(prefix + '::lookup_request', None)
                    st.rerun()
        elif draft == []:
            st.info('尚未找到可解析的派工列。請保留每列的場站、派工類型、設備編號，或修正辨識文字。')
        notice = st.session_state.pop(prefix + '::notice', None)
        if notice:
            st.success(notice)
        show_done = st.checkbox('顯示已完成派工／復原', key=prefix + '::show_done')
        for item in list(orders(cache)):
            if item['state'] == 'deleted':
                continue
            if item['state'] == 'done' and not show_done:
                continue
            with st.container(border=True):
                marker = {'done': '✅', 'blocked': '⏸', 'pending': '🟣'}[item['state']]
                st.write(f"{marker} **{item['station_name']}**｜{item['kind']}｜{item['equipment'] or '未指定設備'}")
                st.caption(f"{item['issued_at']}｜{item['description']}")
                if item['station_name'] not in names:
                    st.warning('此站不在目前配置範圍，待辦仍保留；請切換配置後安排路線。')
                left, right = st.columns(2)
                if left.button('復原' if item['state'] == 'done' else '✓ 完成', key=prefix + item['id'] + 'done'):
                    item.update(state='pending' if item['state'] == 'done' else 'done', updated_at=time.time())
                    save()
                    st.rerun()
                if item['state'] != 'done' and right.button('恢復待辦' if item['state'] == 'blocked' else '暫時無法處理', key=prefix + item['id'] + 'block'):
                    item.update(state='pending' if item['state'] == 'blocked' else 'blocked', updated_at=time.time())
                    save()
                    st.rerun()
                if st.button('移除誤加派工', key=prefix + item['id'] + 'delete'):
                    item.update(state='deleted', updated_at=time.time())
                    save()
                    st.rerun()
        st.download_button('匯出派工紀錄', json.dumps(orders(cache), ensure_ascii=False, indent=2),
                           'work-orders.json', 'application/json', key=prefix + 'export')
        backup = st.file_uploader('還原派工紀錄 JSON', type=['json'], key=prefix + 'backup')
        if backup and st.button('合併還原紀錄', key=prefix + 'restore'):
            try:
                data = json.loads(backup.getvalue())
                if not isinstance(data, list):
                    raise ValueError('須為派工紀錄清單')
                merge(cache, data)
                save()
                st.rerun()
            except (ValueError, TypeError, KeyError):
                st.error('無法讀取此派工紀錄檔。')
    if pending(cache):
        st.caption(f'🟣 尚有 {len(pending(cache))} 件派工／{len(pending_names(cache))} 站，請在派工待辦逐筆完成。')
    if page_mode == '一般分析':
        st.checkbox('🟣 只看未完成派工', key=prefix + '::only')

