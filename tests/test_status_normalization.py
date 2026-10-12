"""Status normalization keeps missing data and safe copies during frequent reruns."""
import ast
import unittest
from pathlib import Path

import numpy as np
import pandas as pd
from pandas.testing import assert_frame_equal


def load_status_normalizers():
    """Load the app's pure functions without starting Streamlit or making requests."""
    source = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
    names = {"coerce_nullable_current_status", "coerce_nullable_station_status"}
    constants = {
        "CURRENT_STATUS_COLUMNS", "STATION_LIVE_COLUMNS", "PERSISTED_STATUS_COLUMNS",
        "STATUS_UNAVAILABLE_TEXT",
    }
    selected = []
    for node in ast.parse(source).body:
        if isinstance(node, ast.FunctionDef) and node.name in names:
            selected.append(node)
        elif isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id in constants for target in node.targets
        ):
            selected.append(node)
    namespace = {"pd": pd, "np": np}
    exec(compile(ast.Module(body=selected, type_ignores=[]), "app.py", "exec"), namespace)
    return namespace


NORMALIZERS = load_status_normalizers()
normalize_status = NORMALIZERS["coerce_nullable_station_status"]
STATUS_COLUMNS = NORMALIZERS["PERSISTED_STATUS_COLUMNS"]


class StatusNormalizationTests(unittest.TestCase):
    def test_nullable_integer_fast_path_clamps_negatives_and_returns_independent_copy(self):
        frame = pd.DataFrame({
            column: pd.Series([8, pd.NA, -2, 0], index=[5, 9, 12, 20], dtype="Int64")
            for column in STATUS_COLUMNS
        })
        frame["場站名稱"] = ["市公所", "轉運站", "圖書館", "體育場"]
        original = frame.copy(deep=True)
        result = normalize_status(frame)
        for column in STATUS_COLUMNS:
            self.assertEqual(result[column].dtype, pd.Int64Dtype())
            self.assertEqual(result.loc[5, column], 8)
            self.assertTrue(pd.isna(result.loc[9, column]))
            self.assertEqual(result.loc[12, column], 0)
        result.loc[5, "2.0 現況"] = 100
        result.loc[9, "空位數"] = 7
        result.loc[5, "場站名稱"] = "已修改"
        assert_frame_equal(frame, original)

    def test_raw_text_fraction_blank_and_boolean_conversion_is_preserved(self):
        raw = ["12.9", "-2.4", "資料未取得", "", None, "無資料", True, False, " 4 ", np.nan]
        frame = pd.DataFrame({column: raw for column in STATUS_COLUMNS})
        original = frame.copy(deep=True)
        result = normalize_status(frame)
        expected = pd.Series([12, 0, pd.NA, pd.NA, pd.NA, pd.NA, 1, 0, 4, pd.NA], dtype="Int64")
        for column in STATUS_COLUMNS:
            pd.testing.assert_series_equal(result[column], expected.rename(column))
        assert_frame_equal(frame, original)

    def test_missing_columns_and_all_missing_integer_columns_keep_index_and_nulls(self):
        for frame in (
            pd.DataFrame({"場站名稱": ["轉運站", "圖書館"]}, index=[3, 10]),
            pd.DataFrame({column: pd.Series(pd.NA, index=[3, 10], dtype="Int64")
                          for column in STATUS_COLUMNS}),
        ):
            with self.subTest(columns=list(frame.columns)):
                result = normalize_status(frame)
                self.assertEqual(result.index.tolist(), [3, 10])
                for column in STATUS_COLUMNS:
                    self.assertEqual(result[column].dtype, pd.Int64Dtype())
                    self.assertTrue(result[column].isna().all())
                assert_frame_equal(normalize_status(result), result)

    def test_current_only_normalization_does_not_change_station_live_fields(self):
        frame = pd.DataFrame({"2.0 現況": pd.Series([5, pd.NA], dtype="Int64"),
                              "空位數": ["未取得", "9.5"]})
        result = NORMALIZERS["coerce_nullable_current_status"](frame)
        self.assertTrue(result["2.0E 現況"].isna().all())
        pd.testing.assert_series_equal(result["空位數"], frame["空位數"])
        result.loc[0, "2.0 現況"] = 20
        self.assertEqual(frame.loc[0, "2.0 現況"], 5)


if __name__ == "__main__":
    unittest.main()
