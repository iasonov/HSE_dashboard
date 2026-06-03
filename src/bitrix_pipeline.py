"""Normalize Bitrix admissions tables and calculate dashboard-ready metrics."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime

import numpy as np
import pandas as pd

from col_names import (
    col_ages,
    col_ages_mean,
    col_applications,
    col_applications_by_week,
    col_contracts,
    col_contracts_by_week,
    col_enrollments,
    col_female,
    col_leads,
    col_male,
    col_payments,
    col_program,
)
from contracts import (
    BITRIX_CONTACTS,
    BITRIX_CONTRACTS,
    BITRIX_DEALS,
    BITRIX_EDUCATIONAL_PROGRAMS,
    BITRIX_EXAMS,
    BITRIX_PORTFOLIOS,
    BitrixEntity,
)
from process import categorize_ages, insert_values, num_years, process_by_week


@dataclass(frozen=True, slots=True)
class BitrixRawTables:
    """Raw Bitrix admissions tables exported from CRM."""

    deals: pd.DataFrame
    contacts: pd.DataFrame
    educational_programs: pd.DataFrame
    contracts: pd.DataFrame
    exams: pd.DataFrame
    portfolios: pd.DataFrame


@dataclass(frozen=True, slots=True)
class NormalizedBitrixAdmissionsData:
    """Normalized Bitrix admissions tables used by dashboard aggregations."""

    applications: pd.DataFrame
    exams: pd.DataFrame
    portfolios: pd.DataFrame


NORMALIZED_APPLICATION_COLUMNS: tuple[str, ...] = (
    "deal_id",
    "contact_id",
    "program_id",
    "program",
    "program_shortname",
    "program_campus",
    "program_level",
    "program_form",
    "application_date",
    "contract_date",
    "payment_date",
    "enrollment_order",
    "gender",
    "birthdate",
)


def _require_columns(frame: pd.DataFrame, entity: BitrixEntity) -> None:
    missing_columns = [column for column in entity.required_fields if column not in frame.columns]
    if missing_columns:
        raise ValueError(
            f"Bitrix table {entity.name!r} is missing required columns: {missing_columns}"
        )


def _as_text_series(series: pd.Series) -> pd.Series:
    return series.astype("string").str.strip()


def _as_datetime_series(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def _payment_dates_by_deal(contracts: pd.DataFrame) -> pd.DataFrame:
    _require_columns(contracts, BITRIX_CONTRACTS)
    prepared = contracts.loc[:, ["iddeal", "data_oplaty"]].copy()
    prepared["deal_id"] = _as_text_series(prepared["iddeal"])
    prepared["payment_date"] = _as_datetime_series(prepared["data_oplaty"])
    prepared = prepared.dropna(subset=["deal_id", "payment_date"])
    if prepared.empty:
        return pd.DataFrame(columns=["deal_id", "payment_date"])
    return prepared.groupby("deal_id", as_index=False)["payment_date"].min()


def _normalize_applications(raw_tables: BitrixRawTables) -> pd.DataFrame:
    _require_columns(raw_tables.deals, BITRIX_DEALS)
    _require_columns(raw_tables.contacts, BITRIX_CONTACTS)
    _require_columns(raw_tables.educational_programs, BITRIX_EDUCATIONAL_PROGRAMS)

    deals = raw_tables.deals.copy()
    deals["deal_id"] = _as_text_series(deals["idaispk"])
    deals["contact_id"] = _as_text_series(deals["idcontact"])
    deals["program_id"] = _as_text_series(deals["idop"])
    deals["application_date"] = _as_datetime_series(deals["date_registrationaispk"])
    deals["contract_date"] = _as_datetime_series(deals["date_dogovora"])
    deals["enrollment_order"] = _as_text_series(deals["prikaz_zachislenya"])

    contacts = raw_tables.contacts.copy()
    contacts["contact_id"] = _as_text_series(contacts["idaispk"])
    contacts["gender"] = _as_text_series(contacts["pol"])
    contacts["birthdate"] = _as_datetime_series(contacts["birthdate"])

    programs = raw_tables.educational_programs.copy()
    programs["program_id"] = _as_text_series(programs["idaispk"])
    programs["program"] = _as_text_series(programs["name"])
    programs["program_level"] = _as_text_series(programs["uroven_obrazovanya"])
    programs["program_campus"] = _as_text_series(programs["campus"])
    if "shortname" not in programs.columns:
        programs["shortname"] = programs["name"]
    if "forma_obuchenya" not in programs.columns:
        programs["forma_obuchenya"] = ""
    programs["program_shortname"] = _as_text_series(programs["shortname"])
    programs["program_form"] = _as_text_series(programs["forma_obuchenya"])

    payments = _payment_dates_by_deal(raw_tables.contracts)
    applications = (
        deals.merge(
            contacts.loc[:, ["contact_id", "gender", "birthdate"]],
            how="left",
            on="contact_id",
            validate="many_to_one",
        )
        .merge(
            programs.loc[
                :,
                [
                    "program_id",
                    "program",
                    "program_shortname",
                    "program_campus",
                    "program_level",
                    "program_form",
                ],
            ],
            how="left",
            on="program_id",
            validate="many_to_one",
        )
        .merge(payments, how="left", on="deal_id", validate="one_to_one")
    )
    return applications.loc[:, NORMALIZED_APPLICATION_COLUMNS]


def _normalize_exams(exams: pd.DataFrame, applications: pd.DataFrame) -> pd.DataFrame:
    _require_columns(exams, BITRIX_EXAMS)
    prepared = exams.copy()
    prepared["deal_id"] = _as_text_series(prepared["iddeal"])
    prepared["contact_id"] = _as_text_series(prepared["idcontact"])
    prepared["exam_score"] = pd.to_numeric(prepared["ball"], errors="coerce")
    prepared["exam_date"] = _as_datetime_series(prepared["date_testirovanya"])
    prepared["is_active"] = prepared["aktive"].astype("boolean")
    return prepared.merge(
        applications.loc[:, ["deal_id", "program", "program_campus", "program_level"]],
        how="left",
        on="deal_id",
        validate="many_to_one",
    )


def _normalize_portfolios(portfolios: pd.DataFrame, applications: pd.DataFrame) -> pd.DataFrame:
    _require_columns(portfolios, BITRIX_PORTFOLIOS)
    prepared = portfolios.copy()
    prepared["deal_id"] = _as_text_series(prepared["iddeal"])
    prepared["contact_id"] = _as_text_series(prepared["idcontact"])
    prepared["product_id"] = _as_text_series(prepared["idtovar"])
    prepared["is_active"] = prepared["status_elementa_portfolio"].astype("boolean")
    return prepared.merge(
        applications.loc[:, ["deal_id", "program", "program_campus", "program_level"]],
        how="left",
        on="deal_id",
        validate="many_to_one",
    )


def normalize_bitrix_admissions_data(raw_tables: BitrixRawTables) -> NormalizedBitrixAdmissionsData:
    """Normalize Bitrix admissions tables to dashboard-friendly tables."""

    applications = _normalize_applications(raw_tables)
    exams = _normalize_exams(raw_tables.exams, applications)
    portfolios = _normalize_portfolios(raw_tables.portfolios, applications)
    return NormalizedBitrixAdmissionsData(applications=applications, exams=exams, portfolios=portfolios)


def process_current_files_from_bitrix(
    raw_tables: BitrixRawTables,
    dashboard_template: pd.DataFrame,
    as_of: datetime,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build current dashboard data from Bitrix tables only."""

    admissions_data = normalize_bitrix_admissions_data(raw_tables)
    dashboard = apply_bitrix_metrics_to_dashboard(dashboard_template, admissions_data, as_of)
    history = pd.DataFrame()
    #     df_history, df_leads_prev, df_leads_after_april_prev, df_applications_prev, df_contracts_prev = process_history_files() 

    # masters_list = df_online_master_programs[col_program].unique()
    # master_2026_no_duplicates = df_master[df_master[master_col_programs].isin(masters_list)].drop_duplicates(subset=[master_col_reg_number])
    # try:
    #     bachelor_2026_no_duplicates = df_bachelor_app.drop_duplicates(subset=[bachelor_col_reg_number])
    # except:
    #     bachelor_2026_no_duplicates = pd.DataFrame(columns=[bachelor_col_programs])

    # df_history.loc[2026, 'applications_unique'] = master_2026_no_duplicates[master_col_programs].count() + bachelor_2026_no_duplicates[bachelor_col_programs].count()
    # df_history.loc[2026, 'early_invitations_unique'] = df_master_early[df_master_early[col_programs_names].isin(masters_list)].drop_duplicates(subset=[col_id_asav])[col_programs_names].count()


    # df_leads_prev = pd.DataFrame({col_program_bitrix:df_leads_prev.index, 'values':df_leads_prev.values})
    # df_leads_after_april_prev = pd.DataFrame({col_program_bitrix:df_leads_after_april_prev.index, 'values':df_leads_after_april_prev.values})

    # try: #TODO change to date comparison from try
    #     main_leads_after_april_prev = df_leads_after_april_prev[df_leads_after_april_prev[col_program_bitrix] == main_studyonline]['values'].values[0]
    # except:
    #     main_leads_after_april_prev = 0

    # try: #TODO change to date comparison from try
    #     main_leads_prev = df_leads_prev[df_leads_prev[col_program_bitrix] == main_studyonline]['values'].values[0]
    # except:
    #     main_leads_prev = 0

    # df_main_dashboard = pd.DataFrame(columns=df_master_dashboard.columns)
    # df_main_dashboard.loc[len(df_main_dashboard)] = {col_program: main_studyonline, 
    #                                                  col_program_bitrix: main_studyonline, 
    #                                                  col_leads: main_leads, 
    #                                                  col_leads_prev : main_leads_prev,
    #                                                  col_leads_after_april: main_leads_after_april, 
    #                                                  col_leads_after_april_prev: main_leads_after_april_prev}
    # df = pd.concat([df_main_dashboard, df_master_dashboard, df_bachelor_dashboard], ignore_index=True, sort=False)


    # df_applications_prev = pd.DataFrame({col_program:df_applications_prev.index, 'values':df_applications_prev.values})
    # df[col_applications_prev] = insert_values(df, df_applications_prev, col_program, col_applications_prev)

    # df_contracts_prev = pd.DataFrame({col_program:df_contracts_prev.index, 'values':df_contracts_prev.values})
    # df[col_contracts_prev] = insert_values(df, df_contracts_prev, col_program, col_contracts_prev)

    # # считываем тренды по неделям (заявки)
    # if now > DATE_01_04_26:
    #     df_bitrix_before_april.rename(columns={'leads_dates': bitrix_col_date}, inplace=True) 

    # df_bitrix = pd.concat([df_bitrix_before_april, df_bitrix_after_april])

    # df_leads_by_week = process_by_week(df_bitrix, col_programs_names, bitrix_col_date)
    # df_leads_by_week = pd.DataFrame({col_program_bitrix:df_leads_by_week[col_programs_names], 'values':df_leads_by_week['count']})
    # df[col_leads_by_week] = insert_values(df, df_leads_by_week, col_program_bitrix, col_leads_by_week)

    # df[col_leads_after_april_prev] = insert_values(df, df_leads_after_april_prev, col_program_bitrix, col_leads_after_april_prev)
    # df[col_leads_prev] = insert_values(df, df_leads_prev, col_program_bitrix, col_leads_prev)
    # df = df.drop(columns=['program_bitrix', 'tg_chat_id', 'campus', 'start_year'])

    # df.fillna(0, inplace=True)

    # # считаем зачисленных, если не посчитаны ранее
    # try:
    #     df_enr = pd.read_excel(enr_file)
    #     print("Данные по зачисленным из базы считаны")
    #     df[col_enrollments] = insert_values(df, df_enr[[col_program, col_enrollments]], col_program, col_enrollments)
    #     df[col_enrollments_foreign] = insert_values(df, df_enr[[col_program, col_enrollments_foreign]], col_program, col_enrollments_foreign)
    # except:
    #     df[col_enrollments] = 0
    #     df[col_enrollments_foreign] = 0
    #     print("Нет базы по зачисленным или она называется не:\n")
    #     print(enr_file)

    # # считаем второстепенные столбцы
    # df[col_leads_total]                            = df[col_leads_partners] + df[col_leads]
    # df[col_conversion_leads_to_contracts]          = df[col_contracts] / df[col_leads_total]
    # df[col_needed_applications]              = round(df[col_plan_rus]/ NEEDED_APPLICATIONS_RATIO)
    # df[col_conversion_applications_to_contracts]   = df[col_contracts] / df[col_applications]
    # df[col_conversion_contracts_to_payments]       = df[col_payments]  / df[col_contracts]
    # df[col_conversion_contracts_to_enrollments]    = df[col_enrollments]  / df[col_contracts]
    # df[col_payments_div_plan_rus]                  = df[col_payments]  / df[col_plan_rus]
    # df[col_payments_div_plan_foreign]              = df[col_payments_foreign]  / df[col_plan_foreign]
    # df[col_income_1year]                           = df['price'] * df[col_payments] / 1000 # from thousands to millions
    # df.loc[df['level'] == 'master', col_income_all]   = df[col_income_1year] * 2
    # df.loc[df['level'] == 'bachelor', col_income_all] = df[col_income_1year] * 4
    # # df[col_income_all       ] = df[col_income_1year]  * (2 if df['level'] == 'master' else 4) # TODO check later
    # df[col_income_1year_hse ] = df[col_income_1year] * df['income_percent'] / 100
    # df[col_income_all_hse   ] = df[col_income_all]   * df['income_percent'] / 100

    # df.replace(np.inf, 0, inplace=True)
    # df.fillna(0, inplace=True)

    # return df, df_history

    return dashboard, history


def _count_by_program(frame: pd.DataFrame, date_column: str) -> pd.DataFrame:
    filtered = frame.dropna(subset=["program", date_column])
    group_columns = ["program", "program_campus", "program_level", "program_form"]
    counts = filtered.groupby(group_columns, dropna=False)["program"].count().reset_index(name="values")
    return counts


def _count_present_by_program(frame: pd.DataFrame, value_column: str) -> pd.DataFrame:
    filtered = frame.dropna(subset=["program"])
    filtered = filtered[filtered[value_column].notna() & (filtered[value_column].astype("string").str.len() > 0)]
    group_columns = ["program", "program_campus", "program_level", "program_form"]
    return filtered.groupby(group_columns, dropna=False)["program"].count().reset_index(name="values")


def _gender_count_by_program(frame: pd.DataFrame, gender_value: str) -> pd.DataFrame:
    filtered = frame.dropna(subset=["program", "gender"])
    group_columns = ["program", "program_campus", "program_level", "program_form"]
    return (
        filtered[filtered["gender"] == gender_value]
        .groupby(group_columns, dropna=False)["program"]
        .count()
        .reset_index(name="values")
    )


def _age_bars_by_program(frame: pd.DataFrame, as_of: datetime) -> pd.DataFrame:
    filtered = frame.dropna(subset=["program", "birthdate"]).copy()
    filtered["age"] = filtered["birthdate"].apply(lambda birthdate: num_years(birthdate, as_of))
    group_columns = ["program", "program_campus", "program_level", "program_form"]
    return filtered.groupby(group_columns, dropna=False)["age"].apply(categorize_ages).reset_index(name="values")


def _age_mean_by_program(frame: pd.DataFrame, as_of: datetime) -> pd.DataFrame:
    filtered = frame.dropna(subset=["program", "birthdate"]).copy()
    filtered["age"] = filtered["birthdate"].apply(lambda birthdate: num_years(birthdate, as_of))
    group_columns = ["program", "program_campus", "program_level", "program_form"]
    return filtered.groupby(group_columns, dropna=False)["age"].mean().reset_index(name="values")


def _insert_metric_by_program_identity(
    dashboard: pd.DataFrame,
    metric_values: pd.DataFrame,
    metric_column: str,
) -> pd.Series:
    if metric_values.empty:
        return pd.Series([0] * len(dashboard), index=dashboard.index)

    left_keys = [col_program]
    right_keys = ["program"]
    optional_keys = (
        ("campus", "program_campus"),
        ("level", "program_level"),
        ("format", "program_form"),
    )
    for left_key, right_key in optional_keys:
        if left_key in dashboard.columns and right_key in metric_values.columns:
            left_keys.append(left_key)
            right_keys.append(right_key)

    prepared_dashboard = dashboard.reset_index(names="__dashboard_index")
    prepared_metrics = metric_values.copy()
    for left_key, right_key in zip(left_keys, right_keys):
        prepared_dashboard[left_key] = prepared_dashboard[left_key].astype("string").str.strip()
        prepared_metrics[right_key] = prepared_metrics[right_key].astype("string").str.strip()
    duplicated_metric_keys = prepared_metrics.duplicated(subset=right_keys, keep=False)
    if duplicated_metric_keys.any():
        duplicated_rows = prepared_metrics.loc[duplicated_metric_keys, right_keys].drop_duplicates().to_dict("records")
        raise ValueError(
            f"Dashboard keys {left_keys} do not disambiguate Bitrix programs for metric {metric_column!r}: "
            f"{duplicated_rows}"
        )

    merged = prepared_dashboard.merge(
        prepared_metrics.loc[:, [*right_keys, "values"]],
        how="left",
        left_on=left_keys,
        right_on=right_keys,
        validate="many_to_one",
    )
    merged = merged.sort_values("__dashboard_index")
    return merged["values"].fillna(0).rename(metric_column)


def apply_bitrix_metrics_to_dashboard(
    dashboard: pd.DataFrame,
    admissions_data: NormalizedBitrixAdmissionsData,
    as_of: datetime,
) -> pd.DataFrame:
    """Fill dashboard metric columns from normalized Bitrix admissions data."""

    result = dashboard.copy()
    applications = admissions_data.applications
    result[col_leads] = _insert_metric_by_program_identity(result, _count_by_program(applications, "application_date"), col_leads)
    result[col_applications] = _insert_metric_by_program_identity(
        result,
        _count_by_program(applications, "application_date"),
        col_applications,
    )
    result[col_contracts] = _insert_metric_by_program_identity(
        result,
        _count_by_program(applications, "contract_date"),
        col_contracts,
    )
    result[col_payments] = _insert_metric_by_program_identity(
        result,
        _count_by_program(applications, "payment_date"),
        col_payments,
    )
    result[col_enrollments] = _insert_metric_by_program_identity(
        result,
        _count_present_by_program(applications, "enrollment_order"),
        col_enrollments,
    )
    result[col_male] = _insert_metric_by_program_identity(result, _gender_count_by_program(applications, "Муж."), col_male)
    result[col_female] = _insert_metric_by_program_identity(result, _gender_count_by_program(applications, "Жен."), col_female)
    result[col_ages] = _insert_metric_by_program_identity(result, _age_bars_by_program(applications, as_of), col_ages)
    result[col_ages_mean] = _insert_metric_by_program_identity(result, _age_mean_by_program(applications, as_of), col_ages_mean)

    applications_by_week = process_by_week(applications, "program", "application_date", "count", "%Y-%m-%d")
    result[col_applications_by_week] = insert_values(
        result,
        pd.DataFrame({col_program: applications_by_week["program"], "values": applications_by_week["count"]}),
        col_program,
        col_applications_by_week,
    )
    contracts_by_week = process_by_week(applications, "program", "contract_date", "count", "%Y-%m-%d")
    result[col_contracts_by_week] = insert_values(
        result,
        pd.DataFrame({col_program: contracts_by_week["program"], "values": contracts_by_week["count"]}),
        col_program,
        col_contracts_by_week,
    )

    result.replace(np.inf, 0, inplace=True)
    result.fillna(0, inplace=True)
    return result
