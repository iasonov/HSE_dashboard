from __future__ import annotations

import unittest
from datetime import datetime

import pandas as pd

from bitrix_pipeline import (
    BitrixRawTables,
    apply_bitrix_metrics_to_dashboard,
    normalize_bitrix_admissions_data,
)
from col_names import (
    col_applications,
    col_contracts,
    col_enrollments,
    col_female,
    col_male,
    col_payments,
    col_program,
)


class TestBitrixPipeline(unittest.TestCase):
    def test_normalize_bitrix_admissions_data_maps_core_fields(self) -> None:
        raw_tables = BitrixRawTables(
            deals=pd.DataFrame(
                {
                    "idaispk": ["deal-1"],
                    "idcontact": ["contact-1"],
                    "idop": ["program-1"],
                    "date_registrationaispk": ["01.06.2026"],
                    "date_dogovora": ["03.06.2026"],
                    "prikaz_zachislenya": ["Приказ 1"],
                }
            ),
            contacts=pd.DataFrame(
                {
                    "idaispk": ["contact-1"],
                    "pol": ["Муж."],
                    "birthdate": ["01.01.2000"],
                }
            ),
            educational_programs=pd.DataFrame(
                {
                    "idaispk": ["program-1"],
                    "name": ["Финансы"],
                    "shortname": ["Финансы"],
                    "uroven_obrazovanya": ["Магистратура"],
                    "campus": ["Москва"],
                    "forma_obuchenya": ["Онлайн"],
                }
            ),
            contracts=pd.DataFrame(
                {
                    "idaispk": ["contract-1"],
                    "iddeal": ["deal-1"],
                    "data_oplaty": ["05.06.2026"],
                }
            ),
            exams=pd.DataFrame(
                {
                    "idaispk": ["exam-1"],
                    "idcontact": ["contact-1"],
                    "iddeal": ["deal-1"],
                    "ball": [80],
                    "date_testirovanya": ["02.06.2026"],
                    "aktive": [True],
                }
            ),
            portfolios=pd.DataFrame(
                {
                    "idaispk": ["portfolio-1"],
                    "idcontact": ["contact-1"],
                    "iddeal": ["deal-1"],
                    "idtovar": ["product-1"],
                    "status_elementa_portfolio": [True],
                }
            ),
        )

        result = normalize_bitrix_admissions_data(raw_tables)

        application = result.applications.iloc[0]
        self.assertEqual(application["deal_id"], "deal-1")
        self.assertEqual(application["program"], "Финансы")
        self.assertEqual(application["program_campus"], "Москва")
        self.assertEqual(application["gender"], "Муж.")
        self.assertEqual(application["payment_date"], pd.Timestamp("2026-06-05"))
        self.assertEqual(result.exams.iloc[0]["program"], "Финансы")
        self.assertEqual(result.portfolios.iloc[0]["program"], "Финансы")

    def test_apply_bitrix_metrics_to_dashboard_counts_date_based_metrics(self) -> None:
        admissions = normalize_bitrix_admissions_data(
            BitrixRawTables(
                deals=pd.DataFrame(
                    {
                        "idaispk": ["deal-1", "deal-2"],
                        "idcontact": ["contact-1", "contact-2"],
                        "idop": ["program-1", "program-1"],
                        "date_registrationaispk": ["01.06.2026", "02.06.2026"],
                        "date_dogovora": ["03.06.2026", None],
                        "prikaz_zachislenya": ["Приказ 1", ""],
                    }
                ),
                contacts=pd.DataFrame(
                    {
                        "idaispk": ["contact-1", "contact-2"],
                        "pol": ["Муж.", "Жен."],
                        "birthdate": ["01.01.2000", "01.01.2001"],
                    }
                ),
                educational_programs=pd.DataFrame(
                    {
                        "idaispk": ["program-1"],
                        "name": ["Финансы"],
                        "uroven_obrazovanya": ["Магистратура"],
                        "campus": ["Москва"],
                    }
                ),
                contracts=pd.DataFrame(
                    {
                        "idaispk": ["contract-1"],
                        "iddeal": ["deal-1"],
                        "data_oplaty": ["05.06.2026"],
                    }
                ),
                exams=pd.DataFrame(
                    columns=["idaispk", "idcontact", "iddeal", "ball", "date_testirovanya", "aktive"]
                ),
                portfolios=pd.DataFrame(
                    columns=["idaispk", "idcontact", "iddeal", "idtovar", "status_elementa_portfolio"]
                ),
            )
        )
        dashboard = pd.DataFrame({col_program: ["Финансы"]})

        result = apply_bitrix_metrics_to_dashboard(dashboard, admissions, datetime(year=2026, month=6, day=10))

        self.assertEqual(result.loc[0, col_applications], 2)
        self.assertEqual(result.loc[0, col_contracts], 1)
        self.assertEqual(result.loc[0, col_payments], 1)
        self.assertEqual(result.loc[0, col_enrollments], 1)
        self.assertEqual(result.loc[0, col_male], 1)
        self.assertEqual(result.loc[0, col_female], 1)


if __name__ == "__main__":
    unittest.main()

