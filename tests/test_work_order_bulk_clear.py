"""Bulk removal of saved V29 work orders must not resurrect browser snapshots."""
import time
import unittest

from work_orders import clear_added_orders, identity, merge, orders, pending


class WorkOrderBulkClearTests(unittest.TestCase):
    def make_order(self, state, equipment, updated_at):
        row = {
            "station_name": "臺東縣政府",
            "kind": "車輛電池",
            "equipment": equipment,
            "issued_at": "2026-10-27 21:24:24",
            "description": "正常營運中車輛，電量低於70%進行派工",
            "state": state,
            "updated_at": updated_at,
        }
        row["id"] = identity(row)
        return row

    def test_clears_every_saved_status_but_preserves_records(self):
        timestamp = time.time() - 60
        original = [
            self.make_order("pending", "1560104", timestamp),
            self.make_order("blocked", "1560296", timestamp),
            self.make_order("done", "1590106", timestamp),
            self.make_order("deleted", "1590142", timestamp),
        ]
        original_ids = [row["id"] for row in original]
        cache = {"metadata": {"__work_orders_v1__": original}}
        self.assertEqual(clear_added_orders(cache), 3)
        self.assertEqual(len(orders(cache)), 4)
        self.assertEqual([row["id"] for row in orders(cache)], original_ids)
        self.assertFalse(pending(cache))
        self.assertTrue(all(row["state"] == "deleted" for row in orders(cache)))

    def test_repeated_action_noop_and_old_browser_snapshot_cannot_restore(self):
        old_time = time.time() - 10
        row = self.make_order("pending", "1560104", old_time)
        cache = {"metadata": {"__work_orders_v1__": [row]}}
        stale_browser_copy = [dict(row)]
        self.assertEqual(clear_added_orders(cache), 1)
        self.assertEqual(clear_added_orders(cache), 0)
        self.assertEqual(merge(cache, stale_browser_copy), 0)
        self.assertEqual(orders(cache)[0]["state"], "deleted")
        self.assertGreater(orders(cache)[0]["updated_at"], old_time)

    def test_empty_workspace_is_noop(self):
        cache = {}
        self.assertEqual(clear_added_orders(cache), 0)
        self.assertEqual(orders(cache), [])


if __name__ == "__main__":
    unittest.main()
