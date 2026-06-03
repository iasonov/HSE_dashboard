from __future__ import annotations

import unittest

from contracts import BITRIX_ADMISSIONS_ENTITIES, BITRIX_DEALS


class TestContracts(unittest.TestCase):
    def test_registry_contains_bitrix_admissions_entities(self) -> None:
        entity_names = {entity.name for entity in BITRIX_ADMISSIONS_ENTITIES}

        self.assertSetEqual(
            entity_names,
            {"deals", "contacts", "educational_programs", "contracts", "exams", "portfolios"},
        )

    def test_deals_contract_uses_date_based_metrics(self) -> None:
        required_fields = set(BITRIX_DEALS.required_fields)

        self.assertIn("date_registrationaispk", required_fields)
        self.assertIn("date_dogovora", required_fields)
        self.assertIn("prikaz_zachislenya", required_fields)


if __name__ == "__main__":
    unittest.main()

