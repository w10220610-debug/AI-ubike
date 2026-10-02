"""Manual dispatch remains usable with mocked GPS, live counts and road data.

Run with: python -m unittest discover -s tests -p test_manual_dispatch.py
Requires the app's Streamlit and openpyxl dependencies; makes no network calls.
"""
import hashlib
import json
import re
import sys
import time
import unittest
import uuid
from io import BytesIO
from pathlib import Path
from unittest.mock import patch
from urllib.parse import parse_qs, urlparse

import streamlit.components.v1 as components
from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

REPO = Path(__file__).resolve().parents[1]
NAMES = ("臺東市公所", "微光市集", "臺東縣立體育場")


class ManualDispatchTests(unittest.TestCase):
    def setUp(self):
        sys.path.insert(0, str(REPO))
        self.addCleanup(lambda: sys.path.remove(str(REPO)))
        self.token = uuid.uuid4().hex
        self.cache = REPO / ".base_cache"
        self.cache.mkdir(exist_ok=True)
        self.addCleanup(self.remove_fixture)
        book = Workbook()
        sheet = book.active
        sheet.title = "平日配置"
        self.records = []
        for index, name in enumerate(NAMES, 1):
            sheet.append([f"D{index}", "", "場站名稱", "", "2.0", "2.0E", "", "2.0", "2.0E", "", "2.0", "2.0E"])
            sheet.append([index, "臺東市", name, "", 5, 5, "", 5, 5, "", 5, 5])
            self.records.append({
                "station_id": f"fixture-{index}", "station_name": name,
                "general_bikes": 8, "electric_bikes": 5,
                "available_spaces": 13, "empty_spaces": 17, "parking_spaces": 30,
                "service_status": 1, "latitude": 22.75 + index * .001,
                "longitude": 121.14 + index * .001,
            })
        excel = self.cache / f"{self.token}.xlsx"
        book.save(excel)
        (self.cache / f"{self.token}.json").write_text(json.dumps({
            "name": "調度測試配置.xlsx", "uploaded_at": time.time(),
            "sha256": hashlib.sha256(excel.read_bytes()).hexdigest(),
        }), encoding="utf-8")

        def factory(name, *args, **kwargs):
            def component(**options):
                if name == "youbike_browser_sync_v2":
                    return {"ok": True, "event_id": "fixture-counts", "records": self.records,
                            "station_count": 3, "requested_station_count": 3}
                if name == "dispatch_geolocation_v3":
                    return {"ok": True, "event_id": "fixture-gps", "latitude": 22.754,
                            "longitude": 121.145, "accuracy": 5}
                return options.get("default")
            return component

        self.factory_patch = patch.object(components, "declare_component", side_effect=factory)
        self.factory_patch.start()
        self.addCleanup(self.factory_patch.stop)
        self.inline_html = []
        original_html = components.html

        def capture_html(html, **options):
            if "const specs = " in html and "__ubikeInlineBatteryGeneration" in html:
                self.inline_html.append(html)
            return original_html(html, **options)

        self.html_patch = patch.object(components, "html", side_effect=capture_html)
        self.html_patch.start()
        self.addCleanup(self.html_patch.stop)
        self.network_patch = patch("urllib.request.urlopen", side_effect=self.road_response)
        self.network_patch.start()
        self.addCleanup(self.network_patch.stop)
        self.app = AppTest.from_file(str(REPO / "app.py"), default_timeout=30)
        self.app.query_params["base"] = self.token
        self.run_app()
        next(item for item in self.app.radio if item.label == "工作模式").set_value("智慧調度")
        self.run_app()

    def remove_fixture(self):
        for suffix in (".xlsx", ".json", ".status.json", ".runtime.json"):
            (self.cache / f"{self.token}{suffix}").unlink(missing_ok=True)

    def road_response(self, request, **kwargs):
        url = request.full_url if hasattr(request, "full_url") else str(request)
        self.assertIn("/table/v1/", url, "Unexpected network access")
        query = parse_qs(urlparse(url).query)
        sources = query["sources"][0].split(";")
        destinations = query["destinations"][0].split(";")
        return BytesIO(json.dumps({
            "code": "Ok", "durations": [[120] * len(destinations) for _ in sources],
            "distances": [[800] * len(destinations) for _ in sources],
        }).encode())

    def run_app(self):
        self.app.run()
        self.assertFalse(self.app.exception, [item.message for item in self.app.exception])

    def press(self, label):
        next(item for item in self.app.button if item.label == label).click()
        self.run_app()

    def state_value(self, suffix, default=None):
        matches = [value for key, value in self.app.session_state.filtered_state.items()
                   if key.startswith("smart_dispatch::") and key.endswith(suffix)]
        self.assertLessEqual(len(matches), 1)
        return matches[0] if matches else default

    def saved_station(self, station_name):
        saved = json.loads((self.cache / f"{self.token}.status.json").read_text(encoding="utf-8"))
        return next(row for rows in saved["contexts"].values() for row in rows
                    if row["場站名稱"] == station_name)

    def test_skip_accept_and_complete_with_manual_quantities(self):
        self.press("⏭️ 跳過並找下一站")
        self.assertEqual(self.state_value("::history")[-1]["action"], "rejected")
        skipped = self.state_value("::history")[-1]["station_name"]
        self.press("✅ 前往此站")
        trip = self.state_value("::active_trip")
        self.assertNotEqual(trip["station_name"], skipped)
        next(item for item in self.app.number_input if item.label == "實際上車 2.0").set_value(2)
        self.press("✅ 完成本站並安排下一站")
        self.assertIsNone(self.state_value("::active_trip"))
        self.assertEqual(self.state_value("::history")[-1]["action"], "completed")
        self.assertEqual(self.state_value("::truck_bike"), 2)
        self.assertEqual(self.saved_station(trip["station_name"])["2.0 現況"], 6)

    def test_inline_pillars_receive_official_ids_before_and_after_accepting(self):
        specs = json.loads(re.search(r"const specs = (.*);", self.inline_html[-1])[1])
        expected_ids = {record["station_name"]: record["station_id"] for record in self.records}
        self.assertTrue(specs)
        for spec in specs:
            self.assertEqual(spec["stationNo"], expected_ids[spec["name"]])
        self.press("✅ 前往此站")
        trip = self.state_value("::active_trip")
        specs = json.loads(re.search(r"const specs = (.*);", self.inline_html[-1])[1])
        self.assertEqual(len(specs), 1)
        self.assertEqual(specs[0]["stationNo"], expected_ids[trip["station_name"]])

    def test_cancel_keeps_station_and_truck_counts(self):
        self.press("✅ 前往此站")
        trip = self.state_value("::active_trip")
        next(item for item in self.app.number_input if item.label == "實際上車 2.0").set_value(2)
        self.press("❌ 取消配置")
        self.assertIsNone(self.state_value("::active_trip"))
        self.assertEqual(self.state_value("::history")[-1]["action"], "cancelled")
        self.assertEqual(self.state_value("::truck_bike"), 0)
        self.assertEqual(self.saved_station(trip["station_name"])["2.0 現況"], 8)

    def test_invalid_completion_keeps_trip_locked_and_counts_unchanged(self):
        self.press("✅ 前往此站")
        trip = self.state_value("::active_trip")
        next(item for item in self.app.number_input if item.label == "實際上車 2.0").set_value(10)
        self.press("✅ 完成本站並安排下一站")
        self.assertEqual(self.state_value("::active_trip")["trip_id"], trip["trip_id"])
        self.assertTrue(any("實際上車數量超過本站" in item.value for item in self.app.error))
        self.assertEqual(self.state_value("::truck_bike"), 0)
        self.assertEqual(self.saved_station(trip["station_name"])["2.0 現況"], 8)
        self.assertFalse(any(row["action"] == "completed" for row in self.state_value("::history", [])))


if __name__ == "__main__":
    unittest.main()
