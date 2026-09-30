from __future__ import annotations

import hashlib
import threading
import time
import unittest
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from background_refresh import BackgroundRefresh
import live_status_service as service


class BackgroundRefreshTests(unittest.TestCase):
    def setUp(self):
        self.now = 0.0
        self.manager = BackgroundRefresh(clock=lambda: self.now, max_workers=2)
        self.release = threading.Event()

    def tearDown(self):
        self.release.set()
        self.manager._pool.shutdown(wait=True, cancel_futures=True)

    def settle(self, key="scope"):
        self.manager._jobs[key].future.result(timeout=2)
        return self.manager.poll(key, lambda: self.fail("unexpected reload"))

    def test_slow_request_returns_immediately_and_is_single_flight(self):
        calls = []
        def load():
            calls.append(1)
            self.release.wait(2)
            return {"records": [1]}
        started = time.monotonic()
        for _ in range(50):
            self.assertTrue(self.manager.poll("scope", load, force=True).pending)
        self.assertLess(time.monotonic() - started, 0.5)
        self.release.set()
        result = self.settle()
        self.assertEqual(calls, [1])
        self.assertEqual(result.value, {"records": [1]})
        result.value["records"].append(2)
        self.assertEqual(self.manager.poll("scope", load).value, {"records": [1]})

    def test_timeout_keeps_old_data_and_discards_late_response(self):
        self.manager.poll("scope", lambda: {"records": ["old"]})
        self.settle()
        self.now = 61
        entered = threading.Event()
        def slow():
            entered.set()
            self.release.wait(2)
            return {"records": ["late"]}
        self.manager.poll("scope", slow)
        self.assertTrue(entered.wait(2))
        self.now = 87
        state = self.manager.poll("scope", slow)
        self.assertFalse(state.pending)
        self.assertIn("25", state.error)
        self.assertEqual(state.value, {"records": ["old"]})
        running = self.manager._jobs["scope"].future
        for _ in range(10):
            self.manager.poll("scope", slow, force=True)
            self.assertIs(self.manager._jobs["scope"].future, running)
        self.release.set()
        state = self.settle()
        self.assertEqual(state.value, {"records": ["old"]})
        self.assertEqual(state.revision, 2)

    def test_failed_refresh_backoff_and_stale_expiry(self):
        self.manager.poll("scope", lambda: {"records": ["old"]})
        self.settle()
        self.now = 61
        def fail():
            raise RuntimeError("offline")
        self.manager.poll("scope", fail)
        state = self.settle()
        self.assertEqual(state.error, "offline")
        self.assertEqual(state.value, {"records": ["old"]})
        self.now = 120
        self.assertFalse(self.manager.poll("scope", fail).pending)
        self.now = 301
        self.assertIsNone(self.manager.poll("scope", fail).value)

    def test_completion_before_deadline_survives_inactive_session(self):
        self.manager.poll("scope", lambda: "ready")
        self.manager._jobs["scope"].future.result(timeout=2)
        self.now = 40
        state = self.manager.poll("scope", lambda: self.fail("unexpected reload"))
        self.assertEqual(state.value, "ready")
        self.assertEqual(state.error, "")

    def test_active_jobs_are_not_evicted_when_capacity_reached(self):
        self.manager.max_entries = 1
        def slow():
            self.release.wait(2)
            return "ready"
        self.manager.poll("scope", slow)
        state = self.manager.poll("second", slow)
        self.assertFalse(state.pending)
        self.assertIn("忙碌", state.error)
        self.assertEqual(len(self.manager._jobs), 1)


class LiveDeadlineTests(unittest.TestCase):
    def test_whole_round_deadline_does_not_wait_for_slow_socket(self):
        release = threading.Event()
        finished = threading.Event()
        def slow(*args, **kwargs):
            try:
                release.wait(2)
                return [], 0, 1
            finally:
                finished.set()
        try:
            with patch.object(service, "LIVE_STATUS_TOTAL_TIMEOUT_SECONDS", 0.05), \
                 patch.object(service, "_request_batch", side_effect=slow):
                started = time.monotonic()
                with self.assertRaises(service.LiveStatusServiceError):
                    service._fetch_parking_records(["S001"])
                self.assertLess(time.monotonic() - started, 0.5)
        finally:
            release.set()
            finished.wait(2)

    def test_partial_results_survive_deadline(self):
        release = threading.Event()
        finished = threading.Event()
        def fetch(numbers, **kwargs):
            if numbers == ["S001"]:
                return [{"station_no": "S001", "available_spaces_detail": {"yb2": 3}}], 0, 1
            try:
                release.wait(2)
                return [], 0, 1
            finally:
                finished.set()
        try:
            with patch.object(service, "LIVE_STATUS_TOTAL_TIMEOUT_SECONDS", 0.05), \
                 patch.object(service, "LIVE_STATUS_BATCH_SIZE", 1), \
                 patch.object(service, "_request_batch", side_effect=fetch):
                records, meta = service._fetch_parking_records(["S001", "S002"])
                self.assertIn("S001", records)
                self.assertEqual(meta["missing_station_ids"], ["S002"])
        finally:
            release.set()
            finished.wait(2)


class AppNonblockingTests(unittest.TestCase):
    def test_configured_page_operates_while_live_and_catalog_requests_are_blocked(self):
        import os
        import pandas as pd
        import battery_upgrade
        import live_status_ui
        from streamlit.testing.v1 import AppTest
        rows = []
        for zone in ("D1", "D2", "D3"):
            rows.extend([[zone, "行政區", "場站名稱", "", "夜班配置", "2.0E", "2.0綁車", "早班配置", "2.0E", "2.0綁車", "晚班配置", "2.0E"],
                         [1, "台東市", zone + "測試站", "", 4, 2, 1, 4, 2, 1, 4, 2]])
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            pd.DataFrame(rows).to_excel(writer, sheet_name="平日", header=False, index=False)
        data = buffer.getvalue()
        token = "b" * 32
        path = Path(__file__).resolve().parents[1] / "app.py"
        gate = threading.Event()
        live = BackgroundRefresh()
        battery = BackgroundRefresh()
        def slow_live(*args, **kwargs):
            gate.wait(20)
            return {"ok": True, "event_id": "slow-event", "records": [], "fetched_at": "fixture"}
        def slow_catalog(*args, **kwargs):
            gate.wait(20)
            return ()
        previous = os.getcwd()
        try:
            with TemporaryDirectory() as temp, \
                 patch.object(live_status_ui, "live_refresh", live), \
                 patch.object(battery_upgrade, "station_map_refresh", battery), \
                 patch.object(service, "get_live_status_for_stations", side_effect=slow_live), \
                 patch.object(battery_upgrade, "get_station_catalog", side_effect=slow_catalog):
                os.chdir(temp)
                app = AppTest.from_file(str(path))
                app.query_params["base"] = token
                app.session_state[f"disk_base_cache::{token}"] = {
                    "token": token, "name": "fixture.xlsx", "bytes": data,
                    "sha256": hashlib.sha256(data).hexdigest(), "expires_at": None}
                gps = {"latitude": 22.75, "longitude": 121.15, "accuracy": 8,
                       "updated_at": time.time(), "source": "gps"}
                app.session_state[f"shared_geolocation::{token}::state"] = gps
                started = time.monotonic()
                app.run(timeout=8)
                self.assertFalse(app.exception)
                self.assertLess(time.monotonic() - started, 5)
                self.assertTrue(any("背景更新中" in x.value for x in app.caption))
                self.assertTrue(any("背景準備中" in x.value for x in app.caption))
                work_mode = next(x for x in app.radio if x.label == "工作模式")
                work_mode.set_value("智慧調度").run(timeout=8)
                self.assertFalse(app.exception)
                self.assertEqual(next(x.value for x in app.radio if x.label == "工作模式"), "智慧調度")
                self.assertEqual(app.session_state[f"disk_base_cache::{token}"]["bytes"], data)
                self.assertEqual(app.session_state[f"shared_geolocation::{token}::state"], gps)
                gate.set()
                live._pool.shutdown(wait=True)
                battery._pool.shutdown(wait=True)
                app.run(timeout=8)
                self.assertFalse(app.exception)
        finally:
            gate.set()
            live._pool.shutdown(wait=True)
            battery._pool.shutdown(wait=True)
            os.chdir(previous)


if __name__ == "__main__":
    unittest.main()
