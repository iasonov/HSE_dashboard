"""Build the ASAV and AIS PK contracts summary from file exports."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from col_names import (
    bachelor_col_payments,
    bachelor_dict,
    col_id_asav,
    col_plan_rus,
    col_program,
    master_col_campus,
    master_col_contracts,
    master_col_payments,
    master_col_program_specialization,
    master_col_programs,
)


SUMMARY_SECTION_COLUMN = "Раздел"
SUMMARY_ROW_TYPE_COLUMN = "Тип строки"
SUMMARY_PROGRAM_COLUMN = "Программа"
SUMMARY_CONTRACTS_COLUMN = "Договоры"
SUMMARY_PAYMENTS_COLUMN = "Оплаты"
SUMMARY_CAPPED_CONTRACTS_COLUMN = "Договоры с учетом лимита"
SUMMARY_CAPPED_PAYMENTS_COLUMN = "Оплаты с учетом лимита"
MASTER_SECTION = "Магистратуры"
BACHELOR_SECTION = "Бакалавриаты"
TOTAL_SECTION = "Суммарно"
PROGRAM_ROW = "program"
TOTAL_ROW = "total"
UNIQUE_ROW = "unique"

SUMMARY_COLUMNS = [
    SUMMARY_SECTION_COLUMN,
    SUMMARY_ROW_TYPE_COLUMN,
    SUMMARY_PROGRAM_COLUMN,
    SUMMARY_CONTRACTS_COLUMN,
    SUMMARY_PAYMENTS_COLUMN,
    col_plan_rus,
    SUMMARY_CAPPED_CONTRACTS_COLUMN,
    SUMMARY_CAPPED_PAYMENTS_COLUMN,
]


def find_asav_aispk_files(data_folder: str | Path) -> tuple[Path, Path]:
    """Find one ASAV file and one AIS PK file in data_folder."""
    folder = Path(data_folder)
    if not folder.is_dir():
        raise FileNotFoundError(f"Data folder does not exist: '{folder}'")

    files = [
        file
        for file in folder.iterdir()
        if file.is_file()
        and file.suffix.casefold() in {".xls", ".xlsx", ".xlsm"}
        and not file.name.startswith("~$")
    ]
    asav_files = [
        file
        for file in files
        if "асав" in file.name.casefold() and "аиспк" not in file.name.casefold()
    ]
    aispk_files = [
        file
        for file in files
        if "аиспк" in file.name.casefold() and "асав" not in file.name.casefold()
    ]
    if len(asav_files) != 1 or len(aispk_files) != 1:
        raise ValueError(
            f"Expected one ASAV and one AIS PK file in '{folder}', "
            f"found ASAV={len(asav_files)}, AIS PK={len(aispk_files)}"
        )
    return asav_files[0], aispk_files[0]


def build_asav_aispk_summary(
    asav_file: str | Path,
    aispk_file: str | Path,
    programs_file: str | Path,
) -> pd.DataFrame:
    """Build contract, payment, unique-applicant, and plan-limited metrics."""
    programs = pd.read_excel(
        programs_file,
        usecols=[col_program, "level", col_plan_rus, "format"],
    )
    programs = programs.loc[programs["format"].ne("offline")].copy()
    programs[col_plan_rus] = pd.to_numeric(programs[col_plan_rus], errors="raise").astype(int)

    master_plans = (
        programs.loc[programs["level"].eq("master"), [col_program, col_plan_rus]]
        .rename(columns={col_program: SUMMARY_PROGRAM_COLUMN})
        .groupby(SUMMARY_PROGRAM_COLUMN, as_index=False)[col_plan_rus]
        .sum()
    )
    master_programs = set(master_plans[SUMMARY_PROGRAM_COLUMN])

    asav_columns = [
        col_id_asav,
        master_col_campus,
        master_col_programs,
        master_col_program_specialization,
        master_col_contracts,
        master_col_payments,
    ]
    asav = pd.read_excel(
        asav_file,
        skiprows=1,
        usecols=lambda column: column in asav_columns or column == "Unnamed: 0",
        dtype=object,
    ).rename(columns={"Unnamed: 0": col_id_asav})
    missing_asav_columns = [column for column in asav_columns if column not in asav.columns]
    if missing_asav_columns:
        raise ValueError(f"ASAV file is missing columns: {missing_asav_columns}")

    specialization = asav[master_col_program_specialization].fillna("").astype(str)
    campus = asav[master_col_campus].fillna("").astype(str).str.strip()
    asav = asav.loc[
        asav[master_col_programs].isin(master_programs)
        & ~specialization.str.contains("офлайн", case=False, regex=False)
        & ~(asav[master_col_programs].eq("Финансы") & campus.ne("Москва"))
    ].copy()
    master_contract_mask = (
        asav[master_col_contracts].notna()
        & asav[master_col_contracts].astype("string").str.strip().ne("").fillna(False)
    )
    master_payment_mask = asav[master_col_payments].fillna("").astype(str).str.strip().eq("Оплачено")
    master_counts = pd.concat(
        [
            asav.loc[master_contract_mask].groupby(master_col_programs).size().rename(SUMMARY_CONTRACTS_COLUMN),
            asav.loc[master_payment_mask].groupby(master_col_programs).size().rename(SUMMARY_PAYMENTS_COLUMN),
        ],
        axis=1,
    ).fillna(0).reset_index(names=SUMMARY_PROGRAM_COLUMN)
    master = master_plans.merge(master_counts, on=SUMMARY_PROGRAM_COLUMN, how="left").fillna(0)

    bachelor_programs = set(programs.loc[programs["level"].eq("bachelor"), col_program])
    display_names = {program: program for program in bachelor_programs}
    display_names.update(
        {
            canonical: source.strip()
            for source, canonical in bachelor_dict.items()
            if canonical in bachelor_programs
        }
    )
    bachelor_plans = programs.loc[
        programs["level"].eq("bachelor"),
        [col_program, col_plan_rus],
    ].copy()
    bachelor_plans[SUMMARY_PROGRAM_COLUMN] = bachelor_plans[col_program].map(display_names)
    bachelor_plans = bachelor_plans.groupby(SUMMARY_PROGRAM_COLUMN, as_index=False)[col_plan_rus].sum()

    aispk_columns = [
        "Заявление отозвано",
        "Уникальный код поступающего",
        "Образовательная программа",
        bachelor_col_payments,
        "Статус",
    ]
    aispk = pd.read_excel(aispk_file, usecols=aispk_columns, dtype=object)
    source_programs = aispk["Образовательная программа"].fillna("").astype(str)
    canonical_programs = source_programs.map(bachelor_dict).fillna(
        source_programs.where(source_programs.isin(bachelor_programs))
    )
    withdrawn = aispk["Заявление отозвано"].fillna("").astype(str).str.strip().str.casefold().eq("да")
    excluded_status = (
        aispk["Статус"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
        .isin({"аннулирован", "расторгнут"})
    )
    aispk = aispk.loc[canonical_programs.notna() & ~withdrawn & ~excluded_status].copy()
    aispk[SUMMARY_PROGRAM_COLUMN] = canonical_programs.loc[aispk.index].map(display_names)
    bachelor_payment_mask = (
        aispk[bachelor_col_payments]
        .fillna("")
        .astype(str)
        .str.strip()
        .eq("Оплачен по квитанциям")
    )
    bachelor_counts = pd.concat(
        [
            aispk.groupby(SUMMARY_PROGRAM_COLUMN).size().rename(SUMMARY_CONTRACTS_COLUMN),
            aispk.loc[bachelor_payment_mask].groupby(SUMMARY_PROGRAM_COLUMN).size().rename(SUMMARY_PAYMENTS_COLUMN),
        ],
        axis=1,
    ).fillna(0).reset_index()
    bachelor = bachelor_plans.merge(bachelor_counts, on=SUMMARY_PROGRAM_COLUMN, how="left").fillna(0)

    for section, frame in ((MASTER_SECTION, master), (BACHELOR_SECTION, bachelor)):
        frame[[SUMMARY_CONTRACTS_COLUMN, SUMMARY_PAYMENTS_COLUMN]] = frame[
            [SUMMARY_CONTRACTS_COLUMN, SUMMARY_PAYMENTS_COLUMN]
        ].astype(int)
        frame[SUMMARY_CAPPED_CONTRACTS_COLUMN] = frame[
            [SUMMARY_CONTRACTS_COLUMN, col_plan_rus]
        ].min(axis=1).astype(int)
        frame[SUMMARY_CAPPED_PAYMENTS_COLUMN] = frame[
            [SUMMARY_PAYMENTS_COLUMN, col_plan_rus]
        ].min(axis=1).astype(int)
        frame[SUMMARY_SECTION_COLUMN] = section
        frame[SUMMARY_ROW_TYPE_COLUMN] = PROGRAM_ROW

    master = master[SUMMARY_COLUMNS].sort_values(SUMMARY_PROGRAM_COLUMN)
    bachelor = bachelor[SUMMARY_COLUMNS].sort_values(SUMMARY_PROGRAM_COLUMN)
    program_rows = pd.concat([master, bachelor], ignore_index=True)
    totals = pd.DataFrame(
        [
            {
                SUMMARY_SECTION_COLUMN: section,
                SUMMARY_ROW_TYPE_COLUMN: TOTAL_ROW,
                SUMMARY_PROGRAM_COLUMN: label,
                **{column: int(frame[column].sum()) for column in SUMMARY_COLUMNS[3:]},
            }
            for section, label, frame in (
                (MASTER_SECTION, "Итого магистратуры", master),
                (BACHELOR_SECTION, "Итого бакалавриаты", bachelor),
                (TOTAL_SECTION, "Итого бакалавриаты и магистратуры", program_rows),
            )
        ]
    )

    master_contract_ids = asav.loc[master_contract_mask, col_id_asav].replace(r"^\s*$", pd.NA, regex=True)
    master_payment_ids = asav.loc[master_payment_mask, col_id_asav].replace(r"^\s*$", pd.NA, regex=True)
    bachelor_contract_ids = aispk["Уникальный код поступающего"].replace(r"^\s*$", pd.NA, regex=True)
    bachelor_payment_ids = aispk.loc[
        bachelor_payment_mask,
        "Уникальный код поступающего",
    ].replace(r"^\s*$", pd.NA, regex=True)
    unique_rows = pd.DataFrame(
        [
            (MASTER_SECTION, "Уникальных абитуриентов в магистратурах", master_contract_ids.nunique(), master_payment_ids.nunique()),
            (BACHELOR_SECTION, "Уникальных абитуриентов в бакалавриатах", bachelor_contract_ids.nunique(), bachelor_payment_ids.nunique()),
            (
                TOTAL_SECTION,
                "Уникальных абитуриентов в бакалавриатах и магистратурах",
                master_contract_ids.nunique() + bachelor_contract_ids.nunique(),
                master_payment_ids.nunique() + bachelor_payment_ids.nunique(),
            ),
        ],
        columns=[
            SUMMARY_SECTION_COLUMN,
            SUMMARY_PROGRAM_COLUMN,
            SUMMARY_CONTRACTS_COLUMN,
            SUMMARY_PAYMENTS_COLUMN,
        ],
    )
    unique_rows[SUMMARY_ROW_TYPE_COLUMN] = UNIQUE_ROW

    summary = pd.concat([program_rows, totals, unique_rows], ignore_index=True)
    summary = summary.reindex(columns=SUMMARY_COLUMNS)
    summary[SUMMARY_COLUMNS[3:]] = summary[SUMMARY_COLUMNS[3:]].astype("Int64")
    return summary
