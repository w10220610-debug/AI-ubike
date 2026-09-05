from __future__ import annotations

import unittest

import pandas as pd

from excel_service import available_sources, discover_zones, parse_route


class ExcelServiceTests(unittest.TestCase):
    def setUp(self):
        self.sheet = pd.DataFrame(
            [
                ["A區", "行政區", "場站名稱", "", "夜班2.0", "夜班2.0E"],
                [1, "東區", "甲站", "", 4, 2],
                ["B區", "行政區", "場站名稱", "", "夜班2.0", "夜班2.0E"],
                [2, "西區", "乙站", "", 3, 1],
            ]
        )

    def test_zones_are_discovered_from_excel_instead_of_hard_coded(self):
        sheets = {"平日": self.sheet}
        self.assertEqual(available_sources(sheets), [("平日", "A區"), ("平日", "B區")])
        self.assertEqual(discover_zones(sheets), ["A區", "B區"])

    def test_each_dynamic_zone_is_parsed_independently(self):
        zone_a = parse_route(self.sheet, "A區", "夜班配置")
        zone_b = parse_route(self.sheet, "B區", "夜班配置")
        self.assertEqual(zone_a["場站名稱"].tolist(), ["甲站"])
        self.assertEqual(zone_b["場站名稱"].tolist(), ["乙站"])
        self.assertEqual(zone_a["2.0E 現況"].tolist(), [2])


if __name__ == "__main__":
    unittest.main()
