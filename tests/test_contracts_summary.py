from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from contracts_summary import (
    BACHELOR_SECTION,
    MASTER_SECTION,
    PROGRAM_ROW,
    SUMMARY_CAPPED_CONTRACTS_COLUMN,
    SUMMARY_CAPPED_PAYMENTS_COLUMN,
    SUMMARY_CONTRACTS_COLUMN,
    SUMMARY_PAYMENTS_COLUMN,
    SUMMARY_PROGRAM_COLUMN,
    SUMMARY_ROW_TYPE_COLUMN,
    SUMMARY_SECTION_COLUMN,
    UNIQUE_ROW,
    build_asav_aispk_summary,
    find_asav_aispk_files,
)


def _write_asav_file(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(
        [
            "Рег. номер",
            "Конкурс в магистратуру",
            None,
            None,
            "Договор, скидки",
            None,
        ]
    )
    sheet.append(
        [
            None,
            "Кампус конкурса",
            "Конкурс в магистратуру",
            "Магистерская специализация",
            "Договор на обучение",
            "Оплата первого периода",
        ]
    )
    sheet.append(["m-1", "Москва", "Аналитика больших данных", None, "Договор 1", "Оплачено"])
    sheet.append(["m-2", "Москва", "Аналитика больших данных", None, "Договор 2", "Оплачено"])
    sheet.append(["m-3", "Москва", "Финансы", None, "Договор 3", None])
    sheet.append(["m-4", "Санкт-Петербург", "Финансы", None, "Договор 4", "Оплачено"])
    sheet.append(["m-5", "Москва", "Аналитика больших данных", "офлайн-трек", "Договор 5", "Оплачено"])
    sheet.append(["m-6", "Москва", "Не онлайн", None, "Договор 6", "Оплачено"])
    workbook.save(path)


def _write_aispk_file(path: Path) -> None:
    rows = pd.DataFrame(
        {
            "Заявление отозвано": [None, None, "Да", None, None, None, None],
            "Уникальный код поступающего": ["b-1", "b-2", "b-3", "b-4", "b-5", "b-5", "b-6"],
            "Образовательная программа": [
                "Глобальные цифровые коммуникации (онлайн)",
                "Глобальные цифровые коммуникации (онлайн) ",
                "Дизайн  (онлайн)",
                "Экономический анализ",
                "Экономический анализ",
                "Экономический анализ",
                "Не онлайн",
            ],
            "Статус оплаты": [
                "Оплачен по квитанциям",
                "Оплачен по квитанциям",
                "Оплачен по квитанциям",
                "Оплачен по квитанциям",
                None,
                "Оплачен по квитанциям",
                "Оплачен по квитанциям",
            ],
            "Статус": [None, None, None, "Аннулирован", None, None, None],
        }
    )
    rows.to_excel(path, index=False)


def _write_programs_file(path: Path) -> None:
    programs = pd.DataFrame(
        {
            "program": [
                "Аналитика больших данных",
                "Финансы",
                "Глобальные цифровые коммуникации - онлайн (О К)",
                "Глобальные цифровые коммуникации (Медиа) - онлайн (О К)",
                "Дизайн - онлайн (О К)",
                "Экономический анализ - онлайн (О К)",
                "Не онлайн",
            ],
            "level": ["master", "master", "bachelor", "bachelor", "bachelor", "bachelor", "master"],
            "plan_rus": [1, 1, 1, 2, 10, 1, 10],
            "format": ["online", "online", "online", "online", "online", "online", "offline"],
        }
    )
    programs.to_excel(path, index=False)


class TestContractsSummary(unittest.TestCase):
    def test_build_summary_applies_filters_aliases_uniques_and_limits(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            asav_file = folder / "2026_АСАВ.xlsx"
            aispk_file = folder / "2026_АИСПК.xlsx"
            programs_file = folder / "programs.xlsx"
            _write_asav_file(asav_file)
            _write_aispk_file(aispk_file)
            _write_programs_file(programs_file)

            summary = build_asav_aispk_summary(asav_file, aispk_file, programs_file)

        master_programs = summary.loc[
            summary[SUMMARY_SECTION_COLUMN].eq(MASTER_SECTION)
            & summary[SUMMARY_ROW_TYPE_COLUMN].eq(PROGRAM_ROW)
        ].set_index(SUMMARY_PROGRAM_COLUMN)
        self.assertEqual(master_programs.loc["Аналитика больших данных", SUMMARY_CONTRACTS_COLUMN], 2)
        self.assertEqual(master_programs.loc["Аналитика больших данных", SUMMARY_PAYMENTS_COLUMN], 2)
        self.assertEqual(master_programs.loc["Аналитика больших данных", SUMMARY_CAPPED_CONTRACTS_COLUMN], 1)
        self.assertEqual(master_programs.loc["Аналитика больших данных", SUMMARY_CAPPED_PAYMENTS_COLUMN], 1)
        self.assertEqual(master_programs.loc["Финансы", SUMMARY_CONTRACTS_COLUMN], 1)
        self.assertEqual(master_programs.loc["Финансы", SUMMARY_PAYMENTS_COLUMN], 0)

        bachelor_programs = summary.loc[
            summary[SUMMARY_SECTION_COLUMN].eq(BACHELOR_SECTION)
            & summary[SUMMARY_ROW_TYPE_COLUMN].eq(PROGRAM_ROW)
        ].set_index(SUMMARY_PROGRAM_COLUMN)
        self.assertEqual(bachelor_programs.loc["Глобальные цифровые коммуникации (онлайн)", SUMMARY_CONTRACTS_COLUMN], 2)
        self.assertEqual(bachelor_programs.loc["Глобальные цифровые коммуникации (онлайн)", SUMMARY_PAYMENTS_COLUMN], 2)
        self.assertEqual(bachelor_programs.loc["Экономический анализ", SUMMARY_CONTRACTS_COLUMN], 2)
        self.assertEqual(bachelor_programs.loc["Экономический анализ", SUMMARY_CAPPED_CONTRACTS_COLUMN], 1)
        self.assertEqual(bachelor_programs.loc["Экономический анализ", SUMMARY_CAPPED_PAYMENTS_COLUMN], 1)

        unique_rows = summary.loc[summary[SUMMARY_ROW_TYPE_COLUMN].eq(UNIQUE_ROW)].set_index(
            SUMMARY_SECTION_COLUMN
        )
        self.assertEqual(unique_rows.loc[MASTER_SECTION, SUMMARY_CONTRACTS_COLUMN], 3)
        self.assertEqual(unique_rows.loc[MASTER_SECTION, SUMMARY_PAYMENTS_COLUMN], 2)
        self.assertEqual(unique_rows.loc[BACHELOR_SECTION, SUMMARY_CONTRACTS_COLUMN], 3)
        self.assertEqual(unique_rows.loc[BACHELOR_SECTION, SUMMARY_PAYMENTS_COLUMN], 3)

    def test_find_files_ignores_combined_summary_and_temporary_files(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            asav_file = folder / "2026_АСАВ.xlsx"
            aispk_file = folder / "2026_АИСПК.xlsx"
            asav_file.touch()
            aispk_file.touch()
            (folder / "2026_АСАВ+АИСПК_сводная.xlsx").touch()
            (folder / "~$2026_АСАВ.xlsx").touch()

            selected_asav, selected_aispk = find_asav_aispk_files(folder)

        self.assertEqual(selected_asav, asav_file)
        self.assertEqual(selected_aispk, aispk_file)


if __name__ == "__main__":
    unittest.main()
