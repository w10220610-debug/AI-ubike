"""Work orders: explicit review, separate completion, cross-shift persistence."""
from __future__ import annotations
import hashlib
import json
import re
import time
import unicodedata
from pathlib import Path

KEY = '__work_orders_v1__'
DRAFT_SCHEMA = 'vehicle-description-v4'

def is_station_code(value):
    return bool(re.fullmatch(r'50\d{7}', str(value or '').strip()))

def is_bike_number(value):
    value = str(value or '').strip()
    return bool(re.fullmatch(r'[0-9]{6,10}', value)) and not is_station_code(value)

FIELDS = ('station_name', 'kind', 'equipment', 'issued_at', 'description', 'source_station', 'lookup_detail')
LABELS = dict(zip(FIELDS, ('場站', '派工類型', '設備編號', '派工時間', '問題描述', '照片場站（參考）', '定位資訊')))


def station_key(value):
    return ''.join(unicodedata.normalize('NFKC', str(value or '')).replace('臺', '台').split()).casefold()


def orders(cache):
    return cache.setdefault('metadata', {}).setdefault(KEY, [])


def pending(cache, station=None):
    return [x for x in orders(cache) if x.get('state') not in ('done', 'deleted') and
            (station is None or station_key(x.get('station_name')) == station_key(station))]


def pending_names(cache):
    return list(dict.fromkeys(x['station_name'] for x in pending(cache)))


def clear_added_orders(cache):
    """Hide every saved order (including completed) without dropping sync tombstones."""
    records = [item for item in orders(cache) if item.get('state') != 'deleted']
    if not records:
        return 0
    now = time.time()
    for item in records:
        try:
            previous = float(item.get('updated_at') or 0)
        except (TypeError, ValueError):
            previous = 0
        # A newer deletion survives older localStorage/backup snapshots.
        item.update(state='deleted', updated_at=max(now, previous + 0.001))
    return len(records)


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
    """Group OCR lines by vehicle, retain complete problem narratives for review.

    Original image station is only a reference. Current location must still be
    independently verified by the vehicle lookup before adding the order.
    """
    catalog = {}
    for name in station_names:
        catalog.setdefault(station_key(name), []).append(name)

    def exact_station(segment):
        compact = station_key(segment)
        hits = [key for key in catalog if key and key in compact]
        maximal = [key for key in hits if not any(key != other and key in other for other in hits)]
        return catalog[maximal[0]][0] if len(maximal) == 1 and len(catalog[maximal[0]]) == 1 else ''

    def type_from(before, narrative):
        before = re.sub(r'\s+', '', before)
        narrative = re.sub(r'\s+', '', narrative)
        if re.search(r'車輛電池|換電池|更換電池|換電', before):
            return '車輛電池'
        if '自行車' in before:
            return '自行車'
        for label in ('車輛故障', '補車', '收車', '綁車', '清潔', '巡檢'):
            if label in before:
                return label
        if ('電量' in narrative and ('低於' in narrative or '%' in narrative)):
            return '車輛電池'
        if any(word in narrative for word in ('同站借還', '交易紀錄', '交易記錄', '在站超過', '更版失敗', '禁用')):
            return '自行車'
        if any(word in narrative for word in ('故障', '異常', '無法借', '無法還')):
            return '車輛故障'
        return ''

    def attach(row, fragment):
        fragment = str(fragment).strip(' \t|｜')
        if not fragment or fragment in ('派工列表', '故障資訊', '維護描述'):
            return
        original = row.get('description', '')
        if not original:
            row['description'] = fragment
        elif fragment not in original:
            # Chinese OCR often breaks a sentence across several image rows.
            spacer = ' ' if original[-1:].isascii() and fragment[:1].isascii() else ''
            row['description'] = original + spacer + fragment
        if not row.get('kind'):
            row['kind'] = type_from('', row['description'])

    rows, context, current = [], [], None
    issued = ''
    timestamp_re = re.compile(r'20\d{2}[./-]\d{1,2}[./-]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?')
    vehicle_re = re.compile(r'(?<![A-Za-z0-9])\d{6,10}(?!\d)')
    for raw in str(text or '').splitlines():
        line = unicodedata.normalize('NFKC', raw).strip()
        if not line:
            continue
        compact = re.sub(r'\s+', '', line)
        if ('車輛號碼' in compact or '設備編號' in compact) and (
                '場站' in compact or '故障資訊' in compact or '派工類型' in compact):
            continue
        stamp = timestamp_re.search(line)
        if stamp:
            issued = stamp.group().replace('.', '-').replace('/', '-')
            current = None
            context = []
            line = timestamp_re.sub('', line).strip()
            if not line:
                continue
        matches = [m for m in vehicle_re.finditer(line) if is_bike_number(m.group())]
        if matches:
            for index, match in enumerate(matches):
                before = line[:match.start()] if index == 0 else ''
                if context and index == 0:
                    before = ' '.join(context + [before])
                tail_end = matches[index + 1].start() if index + 1 < len(matches) else len(line)
                narrative = line[match.end():tail_end].strip(' \t|｜')
                bike = match.group()
                source_station = exact_station(before)
                row = dict(
                    station_name='', kind=type_from(before, narrative),
                    equipment=bike, issued_at=issued, description=narrative,
                    source_station=source_station, detected_name=before.strip(),
                    selected=False, location_confirmed=False,
                )
                rows.append(row)
                current = row
            issued = ''
            context = []
            continue
        if compact in ('場站名稱', '派工類型', '設備編號', '維護描述', '故障資訊', '派工列表'):
            continue
        # Standalone station/type fragments preceding the vehicle ID should
        # become its prefix, not be appended to the previous vehicle's problem.
        standalone_station = exact_station(line)
        if standalone_station and (station_key(line) == station_key(standalone_station)
                                   or compact.endswith('車輛電池')
                                   or compact.endswith('自行車')):
            current = None
            context.append(line)
            continue
        if current is not None:
            attach(current, line)
        else:
            context.append(line)

    if rows:
        return rows

    # Preserve the existing manual, station-only workflow when no vehicle
    # number was recognized; never turn a date or station code into a bike.
    fallback = []
    for raw in str(text or '').splitlines():
        line = unicodedata.normalize('NFKC', raw).strip()
        name = exact_station(line)
        if not name:
            continue
        kind = type_from(line, line)
        if kind:
            fallback.append(dict(station_name=name, kind=kind, equipment='',
                                 issued_at='', description=line, source_station=name,
                                 selected=False))
    return fallback

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
                   prefix + '::notice', prefix + '::lookup_request', prefix + '::legacy_draft'} or key.startswith((prefix + '::editor::', prefix + '::select::')):
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
                           description=description, lookup_detail=description,
                           selected=True, location_confirmed=True))
    return output


def apply_lookup_to_draft(draft, matches, names, kind, searched_ids=None):
    """Fill current parsed rows by exact bike number; append vehicle-only rows."""
    base = [dict(row) for row in draft] if isinstance(draft, list) else []
    searched = {str(x) for x in (searched_ids or []) if is_bike_number(x)}
    for row in base:
        if str(row.get('equipment') or '') in searched:
            row.update(station_name='', selected=False, location_confirmed=False)
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
                # Query metadata is separate: never alter the OCR problem narrative.
                row['lookup_detail'] = found['description']
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
    if state.get(prefix + '::text') or any(is_station_code(row.get('equipment')) for row in draft):
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
        st.caption('自動擷取車號、派工類型及完整敘述，再查詢車輛目前場站；核對後加入。')
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
                          key=prefix + '::ocr::' + DRAFT_SCHEMA + '::partial-v1', default=None)
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
                    event.get('searched_ids'),
                )
                st.session_state[prefix + '::draft_version'] = str(event.get('event_id'))
            elif event.get('type') == 'ocr' and event.get('reset_epoch') == epoch:
                st.session_state[prefix + '::text'] = str(event.get('text') or '')[:100000]
                st.session_state[prefix + '::draft'] = parse_text(event.get('text', ''), names)
                st.session_state[prefix + '::draft_version'] = str(event.get('event_id'))
                st.session_state[prefix + '::lookup_request'] = str(event.get('event_id'))
                st.rerun()
        # The daily flow is photo -> auto location -> one-tap add.
        # Move manual editing/clean-up out of the primary mobile view.
        tool_col_1, tool_col_2 = st.columns(2, gap='small')
        with tool_col_1:
            with st.popover('✍️ 貼上派工文字'):
                text = st.text_area('貼上派工文字', key=prefix + '::text', height=110)
                if st.button('辨識文字並自動查站', key=prefix + '::parse', use_container_width=True):
                    st.session_state[prefix + '::draft'] = parse_text(text, names)
                    st.session_state[prefix + '::draft_version'] = str(time.time_ns())
                    st.session_state[prefix + '::lookup_request'] = st.session_state[prefix + '::draft_version']
                    st.rerun()
        with tool_col_2:
            with st.popover('⚙️ 其他操作'):
                st.button(
                    '🧹 清除本次辨識',
                    key=prefix + '::clear',
                    on_click=clear_draft,
                    args=(st.session_state, prefix),
                    help='清除照片、辨識文字、待確認清單及找車結果；已加入的派工不會被刪除。',
                    use_container_width=True,
                )
                st.caption('不影響已加入派工；需要時也可在下方管理紀錄。')

        draft = st.session_state.get(prefix + '::draft')
        if draft:
            vehicle_mode = any(is_bike_number(x.get('equipment')) for x in draft)
            draft_version = st.session_state.get(prefix + '::draft_version', '0')
            ready = sum(
                bool(x.get('location_confirmed')) and str(x.get('station_name') or '') in names
                for x in draft if is_bike_number(x.get('equipment'))
            )
            vehicle_count = sum(is_bike_number(x.get('equipment')) for x in draft)
            if vehicle_mode:
                st.markdown(f'**🔎 辨識結果｜已定位 {ready}／{vehicle_count} 台**')
                if ready < vehicle_count:
                    st.caption('⏳ 未定位的車會保留待查；不會直接使用截圖上的場站，也不必手動選站。')
                    if st.button('🔄 重新查詢車號位置', key=prefix + '::retry_location'):
                        for row in draft:
                            if is_bike_number(row.get('equipment')):
                                row.update(station_name='', selected=False, location_confirmed=False)
                        st.session_state[prefix + '::lookup_request'] = str(time.time_ns())
                        st.session_state[prefix + '::draft_version'] = st.session_state[prefix + '::lookup_request']
                        st.rerun()

                # Compact row selectors replace the full-width editable table on phones.
                rows = []
                for index, item in enumerate(draft):
                    bike = str(item.get('equipment') or '').strip()
                    station = str(item.get('station_name') or '').strip()
                    kind = str(item.get('kind') or '').strip()
                    located = (bool(item.get('location_confirmed')) and station in names
                               if is_bike_number(bike) else station in names)
                    label = (
                        f"{'✅' if located else '⏳'} {bike or '未指定車號'}｜"
                        f"{station if located else '目前場站待查'}｜{kind or '派工類型待補'}"
                    )
                    with st.container(border=True):
                        checked = st.checkbox(
                            label,
                            value=bool(item.get('selected')) and located,
                            disabled=not located,
                            key=f'{prefix}::select::{draft_version}::{index}',
                        )
                        st.markdown('**📝 派工完整敘述**')
                        st.write(item.get('description') or '⚠️ 未辨識到問題敘述，請在下方修正後再加入。')
                        if item.get('source_station'):
                            st.caption(f"📍 截圖原始場站（僅參考）：{item['source_station']}")
                        if item.get('lookup_detail'):
                            st.caption(f"🔎 {item['lookup_detail']}")
                        if not located:
                            st.caption('目前位置尚未完成車號查詢，不能直接依截圖場站加入。')
                    rows.append({'加入': checked, **{LABELS[k]: item.get(k, '') for k in FIELDS}})

                with st.popover('✏️ 修正派工類型／描述'):
                    st.caption('只需修正有誤的內容；車號和目前場站由系統查詢，不能在此手動指定。')
                    edited = st.data_editor(
                        pd.DataFrame(rows), hide_index=True, num_rows='fixed',
                        disabled=['加入', '場站', '設備編號', '照片場站（參考）', '定位資訊'],
                        column_config={
                            '場站': st.column_config.TextColumn('目前場站（自動查詢）'),
                            '設備編號': st.column_config.TextColumn('車號'),
                            '加入': st.column_config.CheckboxColumn(),
                        },
                        key=prefix + '::editor::' + draft_version,
                        use_container_width=True,
                    )
            else:
                st.info('尚未辨識到車號。可修改辨識文字重新解析；沒有車號的派工仍可依場站人工核對。')
                rows = [{'加入': x.get('selected', False),
                         **{LABELS[k]: x.get(k, '') for k in FIELDS}} for x in draft]
                edited = st.data_editor(
                    pd.DataFrame(rows), hide_index=True, num_rows='dynamic',
                    column_config={
                        '場站': st.column_config.SelectboxColumn(options=names, required=True),
                        '設備編號': st.column_config.TextColumn(),
                        '加入': st.column_config.CheckboxColumn(),
                    },
                    key=prefix + '::editor::' + draft_version,
                    use_container_width=True,
                )

            requested = edited['加入'].fillna(False).astype(bool)
            verified = {(x.get('equipment'), x.get('station_name'))
                        for x in draft if x.get('location_confirmed')}
            eligible = edited.apply(lambda row: (
                str(row['場站']) in names and pd.notna(row['派工類型'])
                and bool(str(row['派工類型']).strip())
                and (not is_bike_number(row['設備編號']) or bool(str(row['問題描述']).strip()))
                and not is_station_code(row['設備編號'])
                and (not is_bike_number(row['設備編號'])
                     or (row['設備編號'], row['場站']) in verified)
            ), axis=1)
            addable_count = int((requested & eligible).sum())
            if st.button(
                f'✅ 加入 {addable_count} 筆派工',
                type='primary',
                key=prefix + '::add',
                use_container_width=True,
            ):
                selected = edited[requested & eligible]
                if selected.empty:
                    if not eligible.any():
                        st.warning('目前沒有可加入的派工。請確認車號已定位、派工類型及完整敘述皆正確；查不到場站可重新查詢。')
                    else:
                        st.warning('請先勾選已定位的派工。')
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
                    st.session_state[prefix + '::notice'] = (
                        f'✅ 已新增 {added} 筆；{len(selected) - added} 筆重複略過；'
                        f'{len(edited) - len(selected)} 筆保留待查／待確認。'
                    )
                    remaining = edited.loc[~edited.index.isin(selected.index)]
                    if not remaining.empty:
                        st.session_state[prefix + '::draft'] = [
                            dict(
                                selected=False,
                                location_confirmed=any(
                                    x.get('location_confirmed') and x.get('equipment') == row['設備編號']
                                    and x.get('station_name') == row['場站'] for x in draft),
                                **{k: str(row[LABELS[k]]) for k in FIELDS},
                            )
                            for row in remaining.fillna('').to_dict('records')
                        ]
                        st.session_state[prefix + '::draft_version'] = str(time.time_ns())
                    else:
                        st.session_state.pop(prefix + '::draft', None)
                        st.session_state.pop(prefix + '::lookup_request', None)
                    st.rerun()
        elif draft == []:
            st.info('沒有辨識出可用的派工列。可換清晰截圖，或用「貼上派工文字」修正再試。')
        notice = st.session_state.pop(prefix + '::notice', None)
        if notice:
            st.success(notice)
        # Saved tasks and backup tools stay accessible without overwhelming photo import.
        if st.toggle(
            f'📋 查看／管理已加入派工（{len(pending(cache))} 件未完成）',
            value=False,
            key=prefix + '::show_tasks',
        ):
            active_orders = [x for x in orders(cache) if x.get('state') != 'deleted']
            if active_orders:
                with st.popover('🗑️ 一鍵清除已加入派工'):
                    st.warning(
                        f'確定要清除全部 {len(active_orders)} 筆已加入派工嗎？'
                        '包含未完成、暫緩與已完成紀錄；'
                        '不會清除上級 EMS 派工，也不會清除目前尚未加入的辨識清單。'
                    )
                    if st.button(
                        f'確認清除全部 {len(active_orders)} 筆',
                        key=prefix + '::clear_all_added',
                        type='primary',
                        use_container_width=True,
                    ):
                        removed = clear_added_orders(cache)
                        if removed:
                            save()
                            st.session_state[prefix + '::notice'] = (
                                f'✅ 已清除 {removed} 筆已加入派工；尚未加入的辨識清單仍保留。'
                            )
                        st.rerun()
            else:
                st.caption('目前沒有已加入的派工可清除。')
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
                    if item.get('source_station'):
                        st.caption(f"📍 截圖場站（僅參考）：{item['source_station']}")
                    if item.get('lookup_detail'):
                        st.caption(f"🔎 {item['lookup_detail']}")
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
        with st.popover('💾 匯出／還原派工紀錄'):
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

