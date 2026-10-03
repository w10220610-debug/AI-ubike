import ast
from pathlib import Path
import time

from work_orders import parse_text, identity, merge, orders, pending, pending_names
from test_performance_cache import prepared_app


def test_photo_rows_unknown_and_duplicate():
    text = '''2026.10.03 21:05:56
臺東航空站 車輛電池 1590155 正常營運中車輛，電量低於70%進行
派工
2026.10.03 21:06:47
康樂火車站 車輛電池 1560081 電量低於70%
找不到的站 車輛電池 1560082 電量低於70%'''
    rows = parse_text(text, ['臺東航空站', '康樂火車站'])
    assert len(rows) == 3
    assert rows[0]['equipment'] == '1590155'
    assert rows[0]['issued_at'] == '2026-10-03 21:05:56'
    assert rows[2]['station_name'] == '' and not rows[2]['selected']
    assert identity(rows[0]) == identity(dict(rows[0], station_name='台東航空站'))


def test_completion_and_restore_do_not_resurrect_stale_backup():
    cache = {}
    row = dict(station_name='A', equipment='1590155', kind='車輛電池', issued_at='2026-10-03 21:05:56', updated_at=1)
    assert merge(cache, [row, row]) == 1
    assert len(pending(cache)) == 1
    orders(cache)[0].update(state='done', updated_at=2)
    assert merge(cache, [row]) == 0
    assert not pending(cache)
    assert merge(cache, [dict(row, state='pending', updated_at=3)]) == 1
    assert pending_names(cache) == ['A']
    # A later work order for the same vehicle is a new assignment.
    assert merge(cache, [dict(row, issued_at='2026-10-04 21:05:56')]) == 1


def test_manual_priority_independent_and_cross_shift():
    ns = prepared_app()
    cache = {}
    merge(cache, [dict(station_name='A', kind='換電', equipment='1590155', updated_at=time.time())])
    bucket = ns['_priority_bucket'](cache, '大夜班')
    bucket['pending'].append({'station_name': 'A'})
    assert ns['_priority_pending_names'](cache, '大夜班') == ['A']
    orders(cache)[0]['state'] = 'done'
    assert ns['_priority_pending_names'](cache, '大夜班') == ['A']
    orders(cache)[0]['state'] = 'pending'
    assert ns['_priority_pending_names'](cache, '其他班') == ['A']


def test_zero_movement_work_order_plan_preserves_load_and_service_guards():
    ns = prepared_app()
    tree = ast.parse(ns['source'])
    wanted = {'build_dispatch_plan_for_station', 'normalize_current_status', 'safe_nonnegative_int',
              'station_total_bikes', 'station_empty_spaces', 'normalize_binding_requirement',
              'normalize_long_distance_zone', 'normalize_dispatch_zone'}
    funcs = [x for x in tree.body if isinstance(x, ast.FunctionDef) and x.name in wanted]
    import pandas as pd
    import re
    ns.update(pd=pd, re=re, status_cache={}, STATUS_UNAVAILABLE_TEXT='未取得', LONG_DISTANCE_ROUTE_ZONES=('D2', 'D3'), ALL_DISPATCH_ZONES=('D1','D2','D3'))
    exec(compile(ast.Module(body=funcs, type_ignores=[]), '<test>', 'exec'), ns)
    row = {'場站名稱': 'A', '2.0 現況': 5, '2.0E 現況': 5, '2.0 標準': 5, '2.0E 標準': 5, '服務狀態': 1}
    kwargs = dict(truck_bike=7, truck_ebike=7, max_capacity=14, global_bike_shortage=0, global_ebike_shortage=0)
    build = ns['build_dispatch_plan_for_station']
    assert build(row, **kwargs) is None
    merge(ns['status_cache'], [dict(station_name='A', kind='換電', equipment='1590155')])
    plan = build(row, **kwargs)
    assert plan['dispatch_count'] == 0
    assert plan['truck_after_bike'] + plan['truck_after_ebike'] == 14
    assert build(dict(row, **{'服務狀態': 0}), **kwargs) is None
    assert build(dict(row, **{'2.0 現況': None}), **kwargs) is None
