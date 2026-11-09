"""Regression tests for V29 manual cargo lookahead."""
import ast
import unittest
from pathlib import Path

from cargo_lookahead import plan_loading_lookahead
from cargo_order import plan_unload

ROOT = Path(__file__).resolve().parents[1]


def row(name, b, e, bs, es, zone="D1"):
    return {
        "場站名稱": name, "2.0 現況": b, "2.0E 現況": e,
        "2.0 標準": bs, "2.0E 標準": es,
        "服務狀態": 1, "路線區域": zone,
    }


LOCATIONS = {
    "A": {"latitude": 22.750, "longitude": 121.145},
    "B": {"latitude": 22.751, "longitude": 121.146},
    "C": {"latitude": 22.753, "longitude": 121.150},
    "D": {"latitude": 22.756, "longitude": 121.154},
}


class LookaheadTests(unittest.TestCase):
    def plan(self, remaining=(), bike=2, ebike=2, rows=None, locations=None, priority=()):
        return plan_loading_lookahead(
            list(remaining), bike, ebike,
            {"station_name": "A", "latitude": 22.750,
             "longitude": 121.145, "route_zone": "D1"},
            rows if rows is not None else [
                row("A", 10, 10, 2, 2),
                row("B", 0, 0, 2, 1),
                row("C", 0, 0, 0, 1),
                row("D", 0, 0, 1, 0),
            ],
            locations if locations is not None else LOCATIONS,
            priority_names=priority,
        )

    def test_looks_ahead_without_changing_counts(self):
        loaded, stops = self.plan()
        self.assertEqual(loaded.count("bike"), 2)
        self.assertEqual(loaded.count("ebike"), 2)
        self.assertTrue(stops)
        self.assertLessEqual(len(stops), 3)
        self.assertEqual(stops[0]["station_name"], "B")
        self.assertLessEqual(stops[0]["temporary_bikes"], 3)

    def test_prefix_is_unchanged_and_reachable(self):
        original = ["ebike", "bike", "bike"]
        added, stops = self.plan(remaining=original, bike=2, ebike=1)
        self.assertEqual(original, ["ebike", "bike", "bike"])
        self.assertEqual((len(added), added.count("bike"), added.count("ebike")), (3, 2, 1))
        full = original + added
        self.assertEqual(len(full), 6)
        for stop in stops:
            full, unloaded, moves, held = plan_unload(
                full, stop["unload_bike"], stop["unload_ebike"]
            )
            self.assertLessEqual(held, 3)
            self.assertGreater(sum(unloaded.values()), 0)
            self.assertEqual(held, stop["temporary_bikes"])

    def test_zero_pickup_can_still_forecast_next_unload(self):
        loaded, stops = self.plan(remaining=["bike", "ebike"], bike=0, ebike=0)
        self.assertEqual(loaded, [])
        self.assertTrue(stops)

    def test_no_location_returns_safe_fallback(self):
        loaded, stops = self.plan(locations={})
        self.assertEqual(loaded.count("bike"), 2)
        self.assertEqual(loaded.count("ebike"), 2)
        self.assertEqual(stops, [])

    def test_priority_raises_probability_and_forecast_is_stable(self):
        rows = [row("A", 5, 5, 0, 0), row("B", 0, 0, 2, 0),
                row("C", 0, 0, 2, 0)]
        first = self.plan(bike=2, ebike=0, rows=rows, priority=("C",))
        second = self.plan(bike=2, ebike=0, rows=rows, priority=("C",))
        self.assertEqual(first, second)
        self.assertEqual(first[1][0]["station_name"], "C")

    def test_invalid_count_is_rejected(self):
        with self.assertRaises(ValueError):
            self.plan(bike=-1)

    def test_application_starts_with_real_import(self):
        tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
        names = {alias.name for stmt in tree.body if isinstance(stmt, ast.ImportFrom)
                 and stmt.module == "cargo_lookahead" for alias in stmt.names}
        self.assertIn("plan_loading_lookahead", names)


if __name__ == "__main__":
    unittest.main()
