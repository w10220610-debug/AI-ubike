"""Regression tests for multi-line V29 image work-order descriptions."""
import unittest

from work_orders import apply_lookup_to_draft, parse_text


class WorkOrderDescriptionTests(unittest.TestCase):
    def setUp(self):
        self.names = [
            "臺東縣政府", "臺東大學附屬體育高級中學",
            "國立臺東大學知本校區正門", "臺東市仁愛國民小學",
            "大橋部落聚會所", "臺東縣立壘球場", "臺東戶政事務所卑南辦公室",
        ]

    def test_example_screenshot_seven_orders_with_full_problems(self):
        captured = """2026.10.27 21:33:50
臺東縣政府 車輛電池 1560299 正常營運中車輛，電量低於70%進行
派工
2026.10.27 21:34:24
臺東大學附屬體育高級中學 車輛電池 1560104 正常營運中車輛，電量低於70%進行派工
2026.10.27 21:26:55
國立臺東大學知本校區正門 自行車 1560296 車輛在站超過168小時，未有交易紀
錄，且車輛狀態為在站/更版失敗禁
用，進行派工
2026.10.27 21:23:25
臺東市仁愛國民小學 車輛電池 1590076 正常營運中車輛，電量低於70%進行派工
2026.10.27 21:27:12
大橋部落聚會所 車輛電池 1590106 正常營運中車輛，電量低於70%進行派工
2026.10.27 21:27:59
臺東縣立壘球場 自行車 1590142 0D03040D-同站借還警告 9次。
2026.10.27 21:24:06
臺東戶政事務所卑南辦公室 車輛電池 1590082 正常營運中車輛，電量低於70%進行派工"""
        rows = parse_text(captured, self.names)
        self.assertEqual(len(rows), 7)
        self.assertEqual([x["equipment"] for x in rows],
                         ["1560299", "1560104", "1560296", "1590076",
                          "1590106", "1590142", "1590082"])
        by_number = {x["equipment"]: x for x in rows}
        self.assertEqual(by_number["1560299"]["description"],
                         "正常營運中車輛，電量低於70%進行派工")
        self.assertIn("更版失敗禁用", by_number["1560296"]["description"])
        self.assertIn("同站借還警告 9次", by_number["1590142"]["description"])
        self.assertEqual(by_number["1590142"]["kind"], "自行車")
        self.assertEqual(by_number["1590142"]["source_station"], "臺東縣立壘球場")
        self.assertEqual(by_number["1590082"]["kind"], "車輛電池")
        self.assertFalse(any(x.get("station_name") for x in rows))

    def test_description_survives_vehicle_location_lookup(self):
        rows = parse_text("臺東縣政府 自行車 1560296 車輛在站超過168小時，未有交易紀錄", self.names)
        result = apply_lookup_to_draft(
            rows,
            [{"bike_no": "1560296", "station_name": "臺東市仁愛國民小學",
              "pillar_no": "07", "queried_at": "2026-10-27 22:00"}],
            self.names,
            "換電",
            searched_ids=["1560296"],
        )
        self.assertEqual(result[0]["station_name"], "臺東市仁愛國民小學")
        self.assertEqual(result[0]["source_station"], "臺東縣政府")
        self.assertEqual(result[0]["kind"], "自行車")
        self.assertIn("在站超過168小時", result[0]["description"])
        self.assertTrue(result[0]["location_confirmed"])

    def test_station_code_is_not_bike_number(self):
        rows = parse_text("臺東縣政府 車輛電池 500012345\n1560299 電量低於70%進行派工", self.names)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["equipment"], "1560299")


if __name__ == "__main__":
    unittest.main()
