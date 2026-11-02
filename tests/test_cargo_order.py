import ast
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
        ui.session_state['test::cargo_draft'] = ['bike']
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
        self.settings(ui)
        self.assertEqual(ui.session_state['test::truck_bike'], 2)
        self.assertEqual(ui.session_state['test::truck_ebike'], 1)
