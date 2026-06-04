"""Normalize Bitrix admissions tables and calculate dashboard-ready metrics."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime

import numpy as np
import pandas as pd

from bitrix import BitrixItemSource, BitrixRestClient, collect_bitrix_item_sources
from col_names import (
    col_ages,
    col_ages_mean,
    col_applications,
    col_applications_by_week,
    col_contracts,
    col_contracts_by_week,
    col_conversion_applications_to_contracts,
    col_conversion_contracts_to_enrollments,
    col_conversion_contracts_to_payments,
    col_conversion_leads_to_contracts,
    col_enrollments,
    col_enrollments_foreign,
    col_female,
    col_income_1year,
    col_income_1year_hse,
    col_income_all,
    col_income_all_hse,
    col_leads,
    col_leads_partners,
    col_leads_total,
    col_male,
    col_needed_applications,
    col_payments,
    col_payments_div_plan_foreign,
    col_payments_div_plan_rus,
    col_payments_foreign,
    col_plan_foreign,
    col_plan_rus,
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
REQUIRED_BITRIX_TABLE_NAMES = (
    "deals",
    "contacts",
    "educational_programs",
    "contracts",
    # "exams",
    # "portfolios",
)
OPTIONAL_ENTITY_SELECT_FIELDS: Mapping[str, tuple[str, ...]] = {
    "educational_programs": ("shortname", "forma_obuchenya"),
}
TECHNICAL_DASHBOARD_COLUMNS = ("program_bitrix", "tg_chat_id", "campus", "start_year", "format")
NEEDED_APPLICATIONS_RATIO = 45 / 100
MALE_VALUES = ("Муж.", "ÐœÑƒÐ¶.")
FEMALE_VALUES = ("Жен.", "Ð–ÐµÐ½.")


def _require_columns(frame: pd.DataFrame, entity: BitrixEntity) -> None:
    missing_columns = [column for column in entity.required_fields if column not in frame.columns]
    if missing_columns:
        raise ValueError(f"Bitrix table {entity.name!r} is missing required columns: {missing_columns}")


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


def create_bitrix_admissions_sources(entity_type_ids: Mapping[str, int]) -> tuple[BitrixItemSource, ...]:
    """Create read-only Bitrix item sources for all admissions tables."""

    missing_names = [name for name in REQUIRED_BITRIX_TABLE_NAMES if name not in entity_type_ids]
    if missing_names:
        raise ValueError(f"Missing Bitrix entity type IDs for admissions tables: {missing_names}")

    entity_by_name: Mapping[str, BitrixEntity] = {
        entity.name: entity
        for entity in (
            BITRIX_DEALS,
            BITRIX_CONTACTS,
            BITRIX_EDUCATIONAL_PROGRAMS,
            BITRIX_CONTRACTS,
            BITRIX_EXAMS,
            BITRIX_PORTFOLIOS,
        )
    }
    sources: list[BitrixItemSource] = []
    for table_name in REQUIRED_BITRIX_TABLE_NAMES:
        entity = entity_by_name[table_name]
        select = tuple(dict.fromkeys((*entity.required_fields, *OPTIONAL_ENTITY_SELECT_FIELDS.get(table_name, ()))))
        sources.append(
            BitrixItemSource(
                name=table_name,
                entity_type_id=int(entity_type_ids[table_name]),
                select=select,
                extra_filter={"CATEGORY_ID" : 4} if table_name == "deals" else {}, # TODO change to resolving CATEGORY_ID by name "Поступление 360"
            )
        )
    return tuple(sources)

# def collect_deals_dataframe(
#     *,
#     webhook_url: str = BITRIX_WEBHOOK_URL,
#     category_name: str = "Поступление 360",
#     category_id: int | None = None,
#     select: Sequence[str] = DEFAULT_DEAL_SELECT,
#     extra_filter: Mapping[str, Any] | None = None,
#     client: BitrixRestClient | None = None,
#     batch_size: int = BITRIX_BATCH_LIMIT,
# ) -> pd.DataFrame:
#     """Collect all deals from the Bitrix CRM funnel "Поступление 360" into one DataFrame.

#     The function uses only read-only REST methods: ``crm.category.list``,
#     ``crm.deal.list`` and read-only subcommands inside ``batch``. Pagination is
#     batched in groups of up to 50 commands, and the client throttles outbound
#     HTTP calls to two requests per second by default.
#     """

#     rest = client or BitrixRestClient(webhook_url)
#     resolved_category_id = (
#         category_id if category_id is not None else get_deal_category_id(category_name, client=rest)
#     )
#     deal_filter: dict[str, Any] = {"CATEGORY_ID": resolved_category_id}
#     if extra_filter:
#         deal_filter.update(extra_filter)
#     base_params: dict[str, Any] = {
#         "select": list(select),
#         "filter": deal_filter,
#         "order": {"ID": "ASC"},
#     }

#     first_response = rest.call("crm.deal.list", {**base_params, "start": 0})
#     deals = list(first_response.get("result", []))
#     total = int(first_response.get("total", len(deals)))
#     if total <= BITRIX_PAGE_SIZE:
#         return pd.DataFrame(deals)

#     max_batch_size = max(1, min(batch_size, BITRIX_BATCH_LIMIT))
#     starts = list(range(BITRIX_PAGE_SIZE, total, BITRIX_PAGE_SIZE)) #TODO test "-1"

#     total_batches = math.ceil(len(starts) / max_batch_size)
#     for batch_index in range(total_batches):
#         chunk_starts = starts[batch_index * max_batch_size : (batch_index + 1) * max_batch_size]
#         commands = {
#             f"deals_{start}": ("crm.deal.list", {**base_params, "start": start})
#             for start in chunk_starts
#         }
#         for page in rest.batch(commands).values():
#             deals.extend(page or [])

#     return pd.DataFrame(deals)


def collect_bitrix_raw_tables(
    client: BitrixRestClient,
    sources: Sequence[BitrixItemSource],
    batch_size: int,
) -> BitrixRawTables:
    """Collect Bitrix admissions item sources into the raw-table container."""

    tables = collect_bitrix_item_sources(client, sources, batch_size)
    missing_names = [name for name in REQUIRED_BITRIX_TABLE_NAMES if name not in tables]
    if missing_names:
        raise ValueError(f"Bitrix raw tables were not collected: {missing_names}")
    return BitrixRawTables(
        deals=tables["deals"],
        contacts=tables["contacts"],
        educational_programs=tables["educational_programs"],
        contracts=tables["contracts"],
        exams=tables["exams"],
        portfolios=tables["portfolios"],
    )


def _count_by_program(frame: pd.DataFrame, date_column: str) -> pd.DataFrame:
    filtered = frame.dropna(subset=["program", date_column])
    group_columns = ["program", "program_campus", "program_level", "program_form"]
    return filtered.groupby(group_columns, dropna=False)["program"].count().reset_index(name="values")


def _count_present_by_program(frame: pd.DataFrame, value_column: str) -> pd.DataFrame:
    filtered = frame.dropna(subset=["program"])
    filtered = filtered[filtered[value_column].notna() & (filtered[value_column].astype("string").str.len() > 0)]
    group_columns = ["program", "program_campus", "program_level", "program_form"]
    return filtered.groupby(group_columns, dropna=False)["program"].count().reset_index(name="values")


def _gender_count_by_program(frame: pd.DataFrame, gender_values: tuple[str, ...]) -> pd.DataFrame:
    filtered = frame.dropna(subset=["program", "gender"])
    group_columns = ["program", "program_campus", "program_level", "program_form"]
    return (
        filtered[filtered["gender"].isin(gender_values)]
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


def _finalize_dashboard_calculations(dashboard: pd.DataFrame) -> pd.DataFrame:
    result = dashboard.copy()
    for column in (col_leads_partners, col_payments_foreign, col_plan_rus, col_plan_foreign):
        if column not in result.columns:
            result[column] = 0
    result.fillna(0, inplace=True)

    result[col_leads_total] = result[col_leads_partners] + result[col_leads]
    result[col_conversion_leads_to_contracts] = result[col_contracts] / result[col_leads_total]
    result[col_needed_applications] = round(result[col_plan_rus] / NEEDED_APPLICATIONS_RATIO)
    result[col_conversion_applications_to_contracts] = result[col_contracts] / result[col_applications]
    result[col_conversion_contracts_to_payments] = result[col_payments] / result[col_contracts]
    result[col_conversion_contracts_to_enrollments] = result[col_enrollments] / result[col_contracts]
    result[col_payments_div_plan_rus] = result[col_payments] / result[col_plan_rus]
    result[col_payments_div_plan_foreign] = result[col_payments_foreign] / result[col_plan_foreign]
    result[col_income_1year] = result["price"] * result[col_payments] / 1000
    result.loc[result["level"] == "master", col_income_all] = result[col_income_1year] * 2
    result.loc[result["level"] == "bachelor", col_income_all] = result[col_income_1year] * 4
    result[col_income_1year_hse] = result[col_income_1year] * result["income_percent"] / 100
    result[col_income_all_hse] = result[col_income_all] * result["income_percent"] / 100
    if col_enrollments_foreign not in result.columns:
        result[col_enrollments_foreign] = 0

    result.replace(np.inf, 0, inplace=True)
    result.fillna(0, inplace=True)
    return result


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
    result[col_male] = _insert_metric_by_program_identity(result, _gender_count_by_program(applications, MALE_VALUES), col_male)
    result[col_female] = _insert_metric_by_program_identity(result, _gender_count_by_program(applications, FEMALE_VALUES), col_female)
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


def process_current_files_from_bitrix(
    client: BitrixRestClient,
    sources: Sequence[BitrixItemSource],
    dashboard_template: pd.DataFrame,
    as_of: datetime,
    batch_size: int,
    history_dataframes: list[pd.DataFrame],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build current dashboard data from Bitrix tables only."""

    raw_tables = collect_bitrix_raw_tables(client, sources, batch_size)
    admissions_data = normalize_bitrix_admissions_data(raw_tables)
    dashboard = apply_bitrix_metrics_to_dashboard(dashboard_template, admissions_data, as_of)
    dashboard = _finalize_dashboard_calculations(dashboard)
    dashboard = dashboard.drop(columns=[column for column in TECHNICAL_DASHBOARD_COLUMNS if column in dashboard.columns])
    
    #TODO add history_data processing and general columns in dashboard
    return dashboard, history_dataframes[0].copy()

