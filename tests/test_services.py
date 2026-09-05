from __future__ import annotations

import threading
import time
import unittest
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import patch

import battery_service
import live_status_service


class BatteryServiceTests(unittest.TestCase):
    def setUp(self):
        battery_service.clear_battery_cache()

    def test_battery_payload_is_normalized_and_clamped(self):
        payload = {
            "data": [
                {"bike_no": "E001", "pillar_no": "03", "battery_power": "105"},
                {"bike_no": "E002", "pillar_no": None, "battery_power": -8},
                {"bike_no": "", "pillar_no": "09", "battery_power": 80},
                {"bike_no": "E004", "pillar_no": "10", "battery_power": "bad"},
            ]
        }
        self.assertEqual(
            battery_service._normalize_bikes(payload),
            (
                {"bike_no": "E001", "pillar_no": "03", "battery_power": 100},
                {"bike_no": "E002", "pillar_no": "", "battery_power": 0},
            ),
        )

    def test_same_station_requests_are_single_flight(self):
        call_count = 0
        call_guard = threading.Lock()

        def fake_request(_station_no, _station_name=""):
            nonlocal call_count
            with call_guard:
                call_count += 1
            time.sleep(0.08)
            return (({"bike_no": "E001", "pillar_no": "01", "battery_power": 80},), 80)

        with patch.object(battery_service, "_request_station_battery", side_effect=fake_request):
            with ThreadPoolExecutor(max_workers=100) as pool:
                results = list(pool.map(lambda _: battery_service.get_station_battery("S001"), range(100)))

        self.assertEqual(call_count, 1)
        self.assertTrue(all(result["bikes"][0]["bike_no"] == "E001" for result in results))

    def test_external_battery_requests_respect_concurrency_limit(self):
        active = 0
        maximum_active = 0
        guard = threading.Lock()

        class FakeResponse:
            def __enter__(self):
                nonlocal active, maximum_active
                with guard:
                    active += 1
                    maximum_active = max(maximum_active, active)
                time.sleep(0.04)
                return self

            def read(self):
                return b'{"data": []}'

            def __exit__(self, *_args):
                nonlocal active
                with guard:
                    active -= 1

        with patch.object(battery_service, "urlopen", side_effect=lambda *_args, **_kwargs: FakeResponse()):
            with ThreadPoolExecutor(max_workers=12) as pool:
                list(pool.map(lambda index: battery_service.get_station_battery(f"S{index:03d}"), range(12)))

        self.assertLessEqual(maximum_active, battery_service.BATTERY_MAX_CONCURRENCY)

    def test_recent_stale_data_is_returned_on_api_failure(self):
        battery_service._cache["S001"] = battery_service.BatterySnapshot(
            station_no="S001",
            station_name="測試站",
            bikes=({"bike_no": "E001", "pillar_no": "01", "battery_power": 70},),
            fetched_at=time.time() - battery_service.BATTERY_CACHE_TTL_SECONDS - 1,
        )
        with patch.object(
            battery_service,
            "_request_station_battery",
            side_effect=battery_service.BatteryServiceError("temporary"),
        ):
            result = battery_service.get_station_battery("S001")

        self.assertEqual(result["source"], "stale_cache")
        self.assertIn("temporary", result["error"])


class LiveStatusServiceTests(unittest.TestCase):
    def setUp(self):
        live_status_service.clear_live_status_cache()
        self.matched = [{
            "station_no": "S001",
            "station_name": "測試站",
            "station_key": "測試站",
            "service_status": 1,
            "latitude": 22.7,
            "longitude": 121.1,
        }]
        self.parking = {
            "S001": {
                "station_no": "S001",
                "available_spaces_detail": {"yb2": 3, "eyb": 2},
                "available_spaces": 5,
                "empty_spaces": 7,
                "parking_spaces": 12,
                "status": 1,
            }
        }

    def test_live_status_is_shared_across_concurrent_sessions(self):
        call_count = 0
        call_guard = threading.Lock()

        def fake_fetch(_station_numbers):
            nonlocal call_count
            with call_guard:
                call_count += 1
            time.sleep(0.08)
            return self.parking, {
                "request_count": 1,
                "failed_request_count": 0,
                "batch_round_count": 1,
                "missing_station_ids": [],
                "api_latency_ms": 80,
            }

        with patch.object(live_status_service, "_prepare_station_records", return_value=(self.matched, [])):
            with patch.object(live_status_service, "_fetch_parking_records", side_effect=fake_fetch):
                with ThreadPoolExecutor(max_workers=20) as pool:
                    results = list(
                        pool.map(
                            lambda _: live_status_service.get_live_status_for_stations([{"name": "測試站"}]),
                            range(20),
                        )
                    )

        self.assertEqual(call_count, 1)
        self.assertTrue(all(result["ok"] for result in results))
        self.assertTrue(all(result["records"][0]["electric_bikes"] == 2 for result in results))

    def test_live_status_uses_stale_data_on_failure(self):
        cache_key = ("S001",)
        live_status_service._cache[cache_key] = live_status_service.LiveStatusSnapshot(
            cache_key=cache_key,
            records=({"station_id": "S001", "station_name": "測試站"},),
            fetched_at_epoch=time.time() - live_status_service.LIVE_STATUS_CACHE_TTL_SECONDS - 1,
            event_id="old-event",
        )
        with patch.object(live_status_service, "_prepare_station_records", return_value=(self.matched, [])):
            with patch.object(
                live_status_service,
                "_fetch_parking_records",
                side_effect=live_status_service.LiveStatusServiceError("temporary"),
            ):
                result = live_status_service.get_live_status_for_stations([{"name": "測試站"}])

        self.assertEqual(result["cache_source"], "stale_cache")
        self.assertEqual(result["event_id"], "old-event")


if __name__ == "__main__":
    unittest.main()
