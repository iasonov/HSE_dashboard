from __future__ import annotations

import unittest

from bitrix import (
    BITRIX_BATCH_LIMIT,
    BITRIX_PAGE_SIZE,
    BitrixReadOnlyError,
    collect_portal_360_deals_dataframe,
    get_deal_category_id,
    _build_batch_command,
)


class FakeBitrixClient:
    def __init__(self) -> None:
        self.calls = []
        self.batch_calls = []

    def call(self, method, params=None):
        self.calls.append((method, params))
        if method == "crm.category.list":
            return {"result": {"categories": [{"id": 7, "name": "Поступление 360"}]}}
        if method == "crm.deal.list":
            return {
                "result": [{"ID": str(i), "CATEGORY_ID": "7"} for i in range(BITRIX_PAGE_SIZE)],
                "total": BITRIX_PAGE_SIZE + 3,
                "next": BITRIX_PAGE_SIZE,
            }
        raise AssertionError(f"Unexpected method {method}")

    def batch(self, commands):
        self.batch_calls.append(commands)
        return {
            "deals_50": [
                {"ID": str(BITRIX_PAGE_SIZE), "CATEGORY_ID": "7"},
                {"ID": str(BITRIX_PAGE_SIZE + 1), "CATEGORY_ID": "7"},
                {"ID": str(BITRIX_PAGE_SIZE + 2), "CATEGORY_ID": "7"},
            ]
        }


class TestBitrixHelpers(unittest.TestCase):
    def test_get_deal_category_id_finds_admissions_funnel(self) -> None:
        client = FakeBitrixClient()

        result = get_deal_category_id("Поступление 360", client)

        self.assertEqual(result, 7)
        self.assertEqual(client.calls[0][0], "crm.category.list")
        self.assertEqual(client.calls[0][1], {"entityTypeId": 2})

    def test_collect_360_deals_dataframe_uses_batch_read_only_pages(self) -> None:
        client = FakeBitrixClient()

        result = collect_360_deals_dataframe(client=client, select=("ID", "CATEGORY_ID"))

        self.assertEqual(len(result), BITRIX_PAGE_SIZE + 3)
        self.assertEqual(result.iloc[-1]["ID"], str(BITRIX_PAGE_SIZE + 2))
        self.assertEqual(client.calls[1][0], "crm.deal.list")
        self.assertEqual(client.calls[1][1]["filter"], {"CATEGORY_ID": 7})
        self.assertEqual(client.batch_calls[0]["items_50"][0], "crm.deal.list")

    def test_build_batch_command_rejects_writing_methods(self) -> None:
        with self.assertRaises(BitrixReadOnlyError):
            _build_batch_command("crm.deal.update", {"id": 1})

    def test_batch_command_encodes_nested_parameters(self) -> None:
        command = _build_batch_command(
            "crm.deal.list",
            {"select": ["ID"], "filter": {"CATEGORY_ID": 7}, "start": 50},
        )

        self.assertIn("crm.deal.list?", command)
        self.assertIn("select%5B%5D=ID", command)
        self.assertIn("filter%5BCATEGORY_ID%5D=7", command)
        self.assertIn("start=50", command)

    def test_batch_limit_constant_matches_bitrix_recommendation(self) -> None:
        self.assertEqual(BITRIX_BATCH_LIMIT, 50)


if __name__ == "__main__":
    unittest.main()
