from __future__ import annotations

import unittest
from datetime import datetime

import pandas as pd

from process import categorize_ages, insert_values, num_years, process_by_week


class TestProcessHelpers(unittest.TestCase):
    def test_insert_values_maps_values_without_manual_loops(self) -> None:
        dashboard = pd.DataFrame({"program": ["A", "B", "C"]})
        values = pd.DataFrame({"program": ["A", "C"], "values": [10, 30]})

        result = insert_values(dashboard, values, "program", "metric")

        self.assertListEqual(result.tolist(), [10, 0, 30])

    def test_process_by_week_builds_weekly_series(self) -> None:
        frame = pd.DataFrame(
            {
                "program": ["A", "A", "B"],
                "date": ["07.10.2025 10:00:00", "14.10.2025 10:00:00", "07.10.2025 10:00:00"],
            }
        )

        result = process_by_week(frame, "program", "date")

        self.assertSetEqual(set(result["program"]), {"A", "B"})
        self.assertTrue(result["count"].str.contains("1").any())

    def test_categorize_ages_counts_groups(self) -> None:
        ages = pd.Series([18, 19, 24, 42, 50])

        result = categorize_ages(ages)

        self.assertEqual(result, "0;2;1;0;0;1;1")

    def test_num_years_handles_exact_birthdays(self) -> None:
        begin = datetime(year=2000, month=3, day=25)
        end = datetime(year=2026, month=3, day=25)

        result = num_years(begin, end)

        self.assertEqual(result, 26)


if __name__ == "__main__":
    unittest.main()
