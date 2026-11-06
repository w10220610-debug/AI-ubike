import ast
import html
import importlib
import itertools
from pathlib import Path
import unittest
import pandas as pd
from cargo_order import (tail_unload, complete, loading_order, choose_loading,
                         order_context, scoped_order)

ROOT = Path(__file__).resolve().parents[1]


def load_plan():
    tree = ast.parse((ROOT / 'app.py').read_text())
    node = next(x for x in tree.body if isinstance(x, ast.FunctionDef) and x.name == 'build_dispatch_plan_for_station')
    env = dict(pd=pd, order_context=order_context, tail_unload=tail_unload,
               choose_loading=choose_loading, status_cache={},
               normalize_current_status=lambda x: None if x is None else int(x),
               safe_nonnegative_int=lambda x: max(0, int(x or 0)),
               work_order_pending=lambda *args: False,
               normalize_long_distance_zone=lambda x: x,
               normalize_binding_requirement=lambda x: x,
               station_total_bikes=lambda r: None, station_empty_spaces=lambda r: None)
    exec(compile(ast.Module(body=[node], type_ignores=[]), 'app.py', 'exec'), env)
    return env[node.name]


def row(name='A', b=0, e=0, bs=2, es=3):
    return {'場站名稱': name, '2.0 現況': b, '2.0E 現況': e,
            '2.0 標準': bs, '2.0E 標準': es, '服務狀態': 1}



class CargoStartupImports(unittest.TestCase):
    def test_app_imports_are_exported_by_cargo_modules(self):
        # Parse without running Streamlit so deployment-time name mismatches
        # fail CI (py_compile alone does not validate imported symbols).
        tree = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
        for stmt in tree.body:
            if isinstance(stmt, ast.ImportFrom) and stmt.module in (
                    'cargo_order', 'cargo_order_ui'):
                module = importlib.import_module(stmt.module)
                for name in stmt.names:
                    self.assertTrue(
                        hasattr(module, name.name),
                        f'app.py imports missing {stmt.module}.{name.name}'
                    )

    def test_app_local_locked_recovery_preserves_order(self):
        from cargo_order import validate
        tree = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
        fn = next(n for n in tree.body
                  if isinstance(n, ast.FunctionDef)
                  and n.name == 'recover_locked_cargo')
        env = {'validate_cargo': validate}
        exec(compile(ast.Module(body=[fn], type_ignores=[]), 'app.py', 'exec'), env)
        ui = type('FakeState', (), {'session_state': {
            'test::cargo_order': [],
            'test::cargo_ready': False,
            'test::cargo_draft': ['ebike'],
        }})()
        trip = {'cargo_before': ['bike', 'ebike', 'bike'],
                'truck_before_bike': 2, 'truck_before_ebike': 1}
        self.assertTrue(env['recover_locked_cargo'](ui, 'test', trip, 14))
        self.assertEqual(ui.session_state['test::cargo_order'],
                         ['bike', 'ebike', 'bike'])
        self.assertTrue(ui.session_state['test::cargo_ready'])
        self.assertNotIn('test::cargo_draft', ui.session_state)


class CargoTests(unittest.TestCase):
    def setUp(self):
        self.token = order_context.set(None)
        self.plan = load_plan()

    def tearDown(self):
        order_context.reset(self.token)

    def plan_for(self, r, b=2, e=3, rows=()):
        return self.plan(r, truck_bike=b, truck_ebike=e, max_capacity=14,
                         global_bike_shortage=5, global_ebike_shortage=5, cargo_rows=rows)

    def test_off_keeps_original_behavior(self):
        p = self.plan_for(row(bs=0))
        self.assertEqual(p['unload_ebike'], 3)
        self.assertNotIn('cargo_before', p)

    def test_blocked_electric_not_recommended(self):
        order_context.set(['ebike']*3 + ['bike']*2)
        self.assertIsNone(self.plan_for(row(bs=0)))

    def test_mixed_station_unloads_in_tail_order(self):
        order_context.set(['ebike']*3 + ['bike']*2)
        p = self.plan_for(row())
        self.assertEqual((p['unload_bike'], p['unload_ebike']), (2, 3))
        self.assertEqual(p['cargo_after'], [])

    def test_partial_unload_stops_at_first_blocker(self):
        order_context.set(['ebike']*3 + ['bike']*2)
        p = self.plan_for(row(bs=1))
        self.assertEqual((p['unload_bike'], p['unload_ebike']), (1, 0))

    def test_stale_counts_fail_closed(self):
        order_context.set(['bike'])
        self.assertIsNone(self.plan_for(row()))

    def test_example_two_stops(self):
        before = ['ebike', 'ebike', 'bike', 'bike']
        after = complete(before, 2, 0, ['ebike', 'ebike'], 14)
        self.assertEqual(complete(after, 0, 3, [], 14), ['ebike'])
        self.assertEqual(before, ['ebike', 'ebike', 'bike', 'bike'])

    def test_manual_interleave_preserved(self):
        self.assertEqual(complete(['bike'], 1, 0, ['bike','ebike','bike'], 14), ['bike','ebike','bike'])

    def test_blocked_actual_completion_rejected(self):
        with self.assertRaises(ValueError):
            complete(['ebike', 'bike'], 0, 1, [], 14)

    def test_capacity_and_invalid_types(self):
        with self.assertRaises(ValueError):
            complete(['bike'], 0, 0, ['ebike'], 1)
        with self.assertRaises(ValueError):
            complete([], 0, 0, ['unknown'], 14)

    def test_loading_considers_existing_cargo_and_next_shortage(self):
        pickup, hint = choose_loading(['ebike'], 2, 2, [row('B', bs=0, es=3)], 'A')
        self.assertEqual(pickup, ['bike','bike','ebike','ebike'])
        self.assertEqual(hint, 'B')
        pickup, _ = choose_loading([], 2, 2, [row('B', bs=2, es=0)], 'A')
        self.assertEqual(pickup, ['ebike','ebike','bike','bike'])

    def test_scope_restores_after_nested_simulation_and_error(self):
        order_context.set(['bike'])
        @scoped_order
        def simulate():
            order_context.set(['ebike'])
            raise ValueError('stop')
        with self.assertRaises(ValueError): simulate()
        self.assertEqual(order_context.get(), ['bike'])

    def test_exhaustive_small_loads_conserve_order_and_counts(self):
        for size in range(7):
            for order in itertools.product(('bike','ebike'), repeat=size):
                for b in range(4):
                    for e in range(4):
                        remaining, down = tail_unload(order, b, e)
                        n = down['bike'] + down['ebike']
                        self.assertEqual(remaining, list(order[:size-n]))
                        self.assertLessEqual(down['bike'], b)
                        self.assertLessEqual(down['ebike'], e)
                        self.assertEqual(complete(order, down['bike'], down['ebike'], [], 14), remaining)


if __name__ == '__main__': unittest.main()

class FakeUI:
    def __init__(self, enabled=False):
        self.session_state = {}
        self.enabled = enabled
        self.clicked = None
    def checkbox(self, *a, key, **kw):
        self.session_state[key] = self.enabled
        return self.enabled
    def expander(self, *a, **kw): return self
    def __enter__(self): return self
    def __exit__(self, *args): pass
    def container(self, *a, **kw): return self
    def markdown(self, *a, **kw): pass
    def caption(self, *a): pass
    def write(self, *a): pass
    def success(self, *a): pass
    def columns(self, widths): return [self] * (widths if isinstance(widths, int) else len(widths))
    def button(self, *a, key, disabled=False, **kw): return self.clicked == key and not disabled


class CargoUITests(unittest.TestCase):
    def settings(self, ui):
        from cargo_order_ui import settings
        return settings(ui, 'test', 14, False, lambda: None)

    def test_default_is_off(self):
        self.assertEqual(self.settings(FakeUI()), (None, True))

    def test_empty_truck_is_automatically_ready(self):
        ui = FakeUI(True)
        self.assertEqual(self.settings(ui), ([], True))
        self.assertTrue(ui.session_state['test::cargo_ready'])

    def test_unconfirmed_nonempty_draft_is_not_autoconfirmed(self):
        ui = FakeUI(True)
        ui.session_state.update({
            'test::cargo_was_enabled': True,
            'test::cargo_draft': ['bike'],
        })
        self.assertEqual(self.settings(ui), ([], False))

    def test_locked_snapshot_restores_confirmed_order_and_counts(self):
        from cargo_order_ui import recover_locked_cargo, settings
        ui = FakeUI(True)
        ui.session_state.update({
            'test::cargo_enabled': True, 'test::cargo_was_enabled': True,
            'test::cargo_ready': False, 'test::cargo_order': [],
            'test::cargo_draft': ['ebike', 'bike'],
            'test::truck_bike': 0, 'test::truck_ebike': 0,
        })
        trip = {'cargo_before': ['bike', 'ebike', 'bike'],
                'truck_before_bike': 2, 'truck_before_ebike': 1}
        self.assertTrue(recover_locked_cargo(ui, 'test', trip, 14))
        self.assertEqual(ui.session_state['test::cargo_order'], ['bike', 'ebike', 'bike'])
        self.assertEqual((ui.session_state['test::truck_bike'],
                          ui.session_state['test::truck_ebike']), (2, 1))
        self.assertNotIn('test::cargo_draft', ui.session_state)
        self.assertEqual(settings(ui, 'test', 14, True, lambda: None),
                         (['bike', 'ebike', 'bike'], True))

    def test_invalid_locked_snapshot_does_not_silently_confirm(self):
        from cargo_order_ui import recover_locked_cargo
        ui = FakeUI(True)
        trip = {'cargo_before': ['bike', 'ebike'],
                'truck_before_bike': 2, 'truck_before_ebike': 0}
        self.assertFalse(recover_locked_cargo(ui, 'test', trip, 14))
        self.assertNotIn('test::cargo_ready', ui.session_state)

    def test_ordinary_locked_trip_does_not_enable_cargo_order(self):
        from cargo_order_ui import recover_locked_cargo
        ui = FakeUI(True)
        ui.session_state['test::cargo_enabled'] = True
        self.assertFalse(recover_locked_cargo(ui, 'test', {'station_name': 'A'}, 14))
        self.assertFalse(ui.session_state['test::cargo_enabled'])

    def test_reenable_requires_confirmation_even_if_counts_same(self):
        ui = FakeUI(True)
        ui.session_state.update({'test::cargo_order':['bike','ebike'],
                                 'test::cargo_ready':True, 'test::cargo_was_enabled':True})
        self.assertTrue(self.settings(ui)[1])
        ui.enabled = False
        self.settings(ui)
        ui.enabled = True
        self.assertFalse(self.settings(ui)[1])

    def test_unconfirmed_draft_blocks_planning(self):
        ui = FakeUI(True)
        ui.session_state.update({'test::cargo_order':['bike','ebike'],
                                 'test::cargo_draft':['ebike','bike'],
                                 'test::cargo_ready':True, 'test::cargo_was_enabled':True})
        self.assertFalse(self.settings(ui)[1])

    def test_confirm_updates_counts_from_order(self):
        ui = FakeUI(True)
        ui.session_state.update({'test::cargo_order':[], 'test::cargo_draft':['bike','ebike','bike'],
                                 'test::cargo_was_enabled':True})
        ui.clicked = 'test::cargo_confirm'
        self.assertEqual(self.settings(ui), (['bike', 'ebike', 'bike'], True))
        self.assertEqual(ui.session_state['test::truck_bike'], 2)
        self.assertEqual(ui.session_state['test::truck_ebike'], 1)

    def test_confirm_fifteen_finishes_render_without_extra_rerun(self):
        from cargo_order_ui import settings
        from unittest.mock import Mock
        ui = FakeUI(True)
        order = ['bike'] * 4 + ['ebike'] * 11
        ui.session_state.update({'test::cargo_was_enabled': True,
                                 'test::cargo_draft': order,
                                 'test::max_capacity': 15,
                                 'test::truck_bike': 4,
                                 'test::truck_ebike': 11})
        ui.clicked = 'test::cargo_confirm'
        rerun = Mock()
        self.assertEqual(settings(ui, 'test', 15, False, rerun), (order, True))
        rerun.assert_not_called()
        self.assertEqual(ui.session_state['test::max_capacity'], 15)
        ui.clicked = None
        self.assertEqual(settings(ui, 'test', 15, False, rerun), (order, True))


class CargoCardHTMLTests(unittest.TestCase):
    """Regression for Streamlit card markup without importing a live Streamlit app."""

    def setUp(self):
        tree = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
        functions = [
            n for n in tree.body if isinstance(n, ast.FunctionDef)
            and n.name in {
                '_dispatch_cargo_chips',
                '_dispatch_truck_before_counts',
                'render_dispatch_plan_card',
            }
        ]
        self.assertEqual(len(functions), 3)
        self.html_fragments = []
        class FakeStreamlit:
            def __init__(self, rendered):
                self.rendered = rendered
            def markdown(self, markup, **kwargs):
                self.rendered.append(markup)
        env = {
            'st': FakeStreamlit(self.html_fragments),
            'html': html,
            'status_cache': {},
            'work_order_badge': lambda cache, name: '',
            'safe_nonnegative_int': lambda n: max(0, int(n or 0)),
            'binding_requirement_display_text': lambda value: '',
            'vehicle_balance_text': lambda current, standard: f'{current} / {standard}',
            'optional_count_text': lambda value: str(value if value is not None else '—'),
            'dispatch_action_text': lambda plan: '本站作業正常',
        }
        exec(compile(ast.Module(body=functions, type_ignores=[]), 'app.py', 'exec'), env)
        self.chips = env['_dispatch_cargo_chips']
        self.render = env['render_dispatch_plan_card']

    def plan(self):
        return {
            'station_name': '大同路與精誠路口', 'region': '台東市',
            'dispatch_count': 6, 'current_bike': 7, 'current_ebike': 7,
            'standard_bike': 3, 'standard_ebike': 5,
            'total_bikes': 14, 'empty_spaces': 4, 'parking_spaces': 18,
            'truck_before_bike': 0, 'truck_before_ebike': 0,
            'truck_after_bike': 4, 'truck_after_ebike': 2,
            'max_capacity': 15, 'routing_fallback': False,
            'road_detour_ratio': 1.0,
            'cargo_pickup': ['bike'] * 4 + ['ebike'] * 2,
            'cargo_after': ['bike'] * 4 + ['ebike'] * 2,
            'cargo_hint': '下一站',
        }

    def test_loading_advice_inside_station_card(self):
        self.render(self.plan(), title='已同意前往／目的地已鎖定')
        card = self.html_fragments[-1]
        self.assertEqual(card.count('dispatch-plan-card"'), 1)
        self.assertIn('本站作業', card)
        self.assertGreater(card.index('dispatch-cargo-panel'), card.index('本站作業'))
        self.assertIn('🍌', card)
        self.assertIn('🍊', card)
        self.assertIn('一般車 2.0 ×4', card)
        self.assertIn('電車 2.0E ×2', card)
        self.assertIn('<details class="dispatch-cargo-details">', card)
        self.assertNotIn('<div class="dispatch-plan-grid">', card)
        for removed in ('路網距離', '行車時間', '預估總時間', '綜合效益'):
            self.assertNotIn(removed, card)

    def test_non_order_mode_keeps_card_without_extra_panel(self):
        plan = self.plan()
        for key in ('cargo_pickup', 'cargo_after', 'cargo_hint'):
            del plan[key]
        self.render(plan, title='推薦場站')
        card = self.html_fragments[-1]
        self.assertIn('dispatch-plan-card', card)
        self.assertNotIn('dispatch-cargo-panel', card)
        self.assertIn('目前車上數量', card)

    def test_mixed_sequence_keeps_actual_interleaving(self):
        view = self.chips(['bike', 'ebike', 'bike', 'bike'])
        self.assertEqual(view.count('dispatch-cargo-chip dispatch-cargo-'), 3)
        self.assertEqual(view.count('dispatch-cargo-arrow'), 2)
        self.assertLess(view.index('🍌'), view.index('🍊'))
        self.assertLess(view.index('🍊'), view.rindex('🍌'))
        self.assertIn('一般車 2.0 ×2', view)
        self.assertIn('本次不上車', self.chips([]))
        self.assertIn('排列資料待確認', self.chips(['invalid']))

    def test_hint_escapes_untrusted_station_name(self):
        plan = self.plan()
        plan['cargo_hint'] = '<img src=x onerror=alert(1)>'
        self.render(plan, title='測試')
        card = self.html_fragments[-1]
        self.assertNotIn('<img', card)
        self.assertIn('&lt;img', card)
