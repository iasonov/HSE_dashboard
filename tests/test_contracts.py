from __future__ import annotations

import unittest

from contracts import METRIC_SPECS


class TestContracts(unittest.TestCase):
    def test_registry_contains_core_metrics_only(self) -> None:
        metric_names = {metric.name for metric in METRIC_SPECS}

        self.assertSetEqual(metric_names, {"leads", "applications", "contracts", "payments", "enrollments"})
        self.assertNotIn("early_invitation", metric_names)

    def test_program_metrics_cover_both_sources(self) -> None:
        source_by_name = {metric.name: metric.source for metric in METRIC_SPECS}

        self.assertEqual(source_by_name["applications"], "asav_and_aispk")
        self.assertEqual(source_by_name["contracts"], "asav_and_aispk")
        self.assertEqual(source_by_name["payments"], "asav_and_aispk")
        self.assertEqual(source_by_name["enrollments"], "asav_and_aispk")


if __name__ == "__main__":
    unittest.main()
