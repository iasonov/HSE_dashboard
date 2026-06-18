"""Normalize Bitrix admissions tables and calculate dashboard-ready metrics."""

from __future__ import annotations

import numpy as np
import pandas as pd

from collections.abc import Mapping, Sequence
from dataclasses import dataclass

from time_const import * 
from bitrix import BitrixRestClient, collect_bitrix_item_sources
from col_names import *
# \(
#     col_ages,
#     col_ages_mean,
#     col_applications,
#     col_applications_by_week,
#     col_contracts,
#     col_contracts_by_week,
#     col_conversion_applications_to_contracts,
#     col_conversion_contracts_to_enrollments,
#     col_conversion_contracts_to_payments,
#     col_conversion_leads_to_contracts,
#     col_enrollments,
#     col_enrollments_foreign,
#     col_female,
#     col_income_1year,
#     col_income_1year_hse,
#     col_income_all,
#     col_income_all_hse,
#     col_leads,
#     col_leads_partners,
#     col_leads_total,
#     col_male,
#     col_needed_applications,
#     col_payments,
#     col_payments_div_plan_foreign,
#     col_payments_div_plan_rus,
#     col_payments_foreign,
#     col_plan_foreign,
#     col_plan_rus,
#     col_program,
#     main_studyonline,
#     leads_dates,
#     applications_dates,
#     contracts_dates,
#     col_program_bitrix,
# )
from contracts import (
    BITRIX_CONTACTS,
    BITRIX_CRM_DEALS,
    BITRIX_PORTAL_DEALS,
    BITRIX_APPLICATIONS,
    BITRIX_EDUCATIONAL_PROGRAMS,
    BITRIX_ADMISSIONS_ENTITIES,
    # BITRIX_CONTRACTS,
    # BITRIX_EXAMS,
    # BITRIX_PORTFOLIOS,
    BitrixEntity,
)
from process import categorize_ages, insert_values, num_years, process_by_week


@dataclass(slots=True)
class BitrixRawTables:
    """Raw Bitrix admissions tables exported from CRM."""

    crm_deals: pd.DataFrame
    portal_deals: pd.DataFrame
    applications: pd.DataFrame
    # contacts: pd.DataFrame
    educational_programs: pd.DataFrame
    # TODO uncomment
    # contracts: pd.DataFrame
    # exams: pd.DataFrame
    # portfolios: pd.DataFrame


# @dataclass(frozen=True, slots=True)
# class NormalizedBitrixAdmissionsData:
#     """Normalized Bitrix admissions tables used by dashboard aggregations."""

#     applications: pd.DataFrame
#     exams: pd.DataFrame
#     portfolios: pd.DataFrame


# NORMALIZED_APPLICATION_COLUMNS: tuple[str, ...] = (
#     "deal_id",
#     "contact_id",
#     "program_id",
#     "program",
#     # "program_shortname", # TODO uncomment
#     # "program_campus",
#     # "program_level",
#     # "program_form",
#     "application_date",
#     "contract_date",
#     "payment_date",
#     "enrollment_order",
#     "gender",
#     "birthdate",
# )

REQUIRED_BITRIX_TABLE_NAMES = tuple(item.name for item in BITRIX_ADMISSIONS_ENTITIES) 
# (
#     "crm_deals",
#     "contacts",
#     "educational_programs",
#     "portal_deals",
#     "aispk_applications",
#     # "contracts",
#     # "exams",
#     # "portfolios",
# )
# OPTIONAL_ENTITY_SELECT_FIELDS: Mapping[str, tuple[str, ...]] = {
#     "educational_programs": ("shortname", "forma_obuchenya"),
# }
TECHNICAL_DASHBOARD_COLUMNS = ("program_bitrix", "tg_chat_id", "campus", "start_year", "format")
NEEDED_APPLICATIONS_RATIO = 45 / 100
MALE_VALUES = ("Муж.", "ÐœÑƒÐ¶.")
FEMALE_VALUES = ("Жен.", "Ð–ÐµÐ½.")


def _require_columns(frame: pd.DataFrame, entity: BitrixEntity) -> None:
    missing_columns = [column for column in entity.select if column not in frame.columns]
    if missing_columns:
        raise ValueError(f"Bitrix table {entity.name!r} is missing required columns: {missing_columns}")


# def _as_text_series(series: pd.Series) -> pd.Series:
#     return series.astype("string").str.strip()


# def _as_datetime_series(series: pd.Series) -> pd.Series:
#     return pd.to_datetime(series, errors="raise", dayfirst=True) # may be coerce?


# def _payment_dates_by_deal(contracts: pd.DataFrame) -> pd.DataFrame:
#     _require_columns(contracts, BITRIX_CONTRACTS)
#     prepared = contracts.loc[:, ["iddeal", "data_oplaty"]].copy()
#     prepared["deal_id"] = _as_text_series(prepared["iddeal"])
#     prepared["payment_date"] = _as_datetime_series(prepared["data_oplaty"])
#     prepared = prepared.dropna(subset=["deal_id", "payment_date"])
#     if prepared.empty:
#         return pd.DataFrame(columns=["deal_id", "payment_date"])
#     return prepared.groupby("deal_id", as_index=False)["payment_date"].min()


def _normalize_raw_data(raw_tables: BitrixRawTables) -> BitrixRawTables:
    # _require_columns(raw_tables.crm_deals, BITRIX_CRM_DEALS)
    # _require_columns(raw_tables.portal_deals, BITRIX_PORTAL_DEALS)
    # _require_columns(raw_tables.applications, BITRIX_APPLICATIONS)
    # _require_columns(raw_tables.contacts, BITRIX_CONTACTS)
    # _require_columns(raw_tables.educational_programs, BITRIX_EDUCATIONAL_PROGRAMS)

    # TODO перенести все строковые константы в col_names.py
    raw_tables.crm_deals['ufDealEducationProgram'] = raw_tables.crm_deals['ufDealEducationProgram'].fillna(0).astype(int, errors='raise') # TODO repair, пустые - преимущественно, но не только разводящий лендинг, но и другие программы (см. вкладку Пустоты в ufDealProgram в файле 2026_06_18_..xlsx)
    raw_tables.portal_deals['ufDealEducationProgram'] = raw_tables.portal_deals['ufDealEducationProgram'].fillna(-1).astype(int, errors='raise') # TODO check -1 - не наша программа
    raw_tables.applications['ufDealEducationProgram'] = raw_tables.applications['ufDealEducationProgram'].fillna(-2).astype(int, errors='raise') # TODO repair - там и наши ленды, и не наши ленды, см. Пустоты в ufDealProgram_APP

    raw_tables.educational_programs = raw_tables.educational_programs[['ID', 'NAME']] # убираем ненужные столбцы, альтернативно можно не забирать их с помощью SELECT
    raw_tables.educational_programs['ID'] = raw_tables.educational_programs['ID'].astype(int, errors='raise') # в целом тут пустых быть не должно
    raw_tables.educational_programs = raw_tables.educational_programs.rename(columns={'ID': 'ufDealEducationProgram', 'NAME': col_program_bitrix}) # переименовываем столбцы для удобства merge
    raw_tables.educational_programs = pd.concat([raw_tables.educational_programs, pd.DataFrame({
        'ufDealEducationProgram': [0,                -1,                                           -2],
        col_program_bitrix:       [main_studyonline, 'Не указана программа в воронке Портала ВШЭ', 'Не указана программа в воронке МАГ/БАК']
    })], ignore_index=True)
    # raw_tables.educational_programs = raw_tables.educational_programs.append({'ufDealEducationProgram': 0, col_program: main_studyonline})
    # raw_tables.educational_programs = raw_tables.educational_programs.append({'ufDealEducationProgram': -1, col_program: 'Не указана программа в воронке Портала ВШЭ'})
    # raw_tables.educational_programs = raw_tables.educational_programs.append({'ufDealEducationProgram': -2, col_program: 'Не указана программа в воронке МАГ/БАК'})

    raw_tables.crm_deals    = pd.merge(raw_tables.crm_deals,    raw_tables.educational_programs, on="ufDealEducationProgram", how="left").drop(columns=['ufDealEducationProgram'])
    raw_tables.portal_deals = pd.merge(raw_tables.portal_deals, raw_tables.educational_programs, on="ufDealEducationProgram", how="left").drop(columns=['ufDealEducationProgram'])
    raw_tables.applications = pd.merge(raw_tables.applications, raw_tables.educational_programs, on="ufDealEducationProgram", how="left").drop(columns=['ufDealEducationProgram'])
    
    raw_tables.crm_deals[leads_dates]           = pd.to_datetime(raw_tables.crm_deals['createdTime'], errors='raise').dt.tz_localize(None)
    raw_tables.portal_deals[leads_dates]        = pd.to_datetime(raw_tables.portal_deals['createdTime'], errors='raise').dt.tz_localize(None)
    raw_tables.applications[applications_dates] = pd.to_datetime(raw_tables.applications['createdTime'], errors='raise').dt.tz_localize(None)
    raw_tables.applications[contracts_dates]    = pd.to_datetime(raw_tables.applications['ufDealContractdate'], errors='raise').dt.tz_localize(None) # TODO check
    
    # deals = raw_tables.crm_deals.copy() # START strange things with convertion, NaN & NaT
    # deals["deal_id"] = _as_text_series(deals["id"]) # .astype("string").str.strip()
    # deals["contact_id"] = _as_text_series(deals["contactId"])
    # deals["program_id"] = _as_text_series(deals["ufDealEducationProgram"])
    # deals["application_date"] = _as_datetime_series(deals["ufDealDataRegistracii"]) # TODO check pd.to_datetime(series, errors="raise", dayfirst=True) # may be coerce?
    # deals["contract_date"] = _as_datetime_series(deals["ufDealContractdate"])
    # deals["enrollment_order"] = _as_text_series(deals["ufDealPrikazOZachislenii"])

    # contacts = raw_tables.contacts.copy()
    # contacts["contact_id"] = _as_text_series(contacts["id"])
    # # contacts["gender"] = _as_text_series(contacts["pol"]) # TODO complete gender
    # contacts["birthdate"] = _as_datetime_series(contacts["birthdate"])

    # programs = raw_tables.educational_programs.copy()
    # programs["program_id"] = _as_text_series(programs["ID"])
    # programs["program"] = _as_text_series(programs["NAME"])
    # # programs["program_level"] = _as_text_series(programs["uroven_obrazovanya"]) # TODO get from deal
    # # programs["program_campus"] = _as_text_series(programs["campus"]) # TODO get from deal
    # # if "shortname" not in programs.columns:
    # #     programs["shortname"] = programs["name"]
    # # if "forma_obuchenya" not in programs.columns:
    # #     programs["forma_obuchenya"] = ""
    # # programs["program_shortname"] = _as_text_series(programs["shortname"])
    # # programs["program_form"] = _as_text_series(programs["forma_obuchenya"])

    # # TODO complete contracts
    # # payments = _payment_dates_by_deal(raw_tables.contracts)
    # applications = (
    #     deals.merge(
    #         contacts.loc[:, ["contact_id", "gender", "birthdate"]],
    #         how="left",
    #         on="contact_id",
    #         validate="many_to_one",
    #     )
    #     .merge(
    #         programs.loc[
    #             :,
    #             [
    #                 "program_id",
    #                 "program",
    #                 # "program_shortname",
    #                 # "program_campus",
    #                 # "program_level",
    #                 # "program_form",
    #             ],
    #         ],
    #         how="left",
    #         on="program_id",
    #         validate="many_to_one",
    #     )
    #     #.merge(payments, how="left", on="deal_id", validate="one_to_one")
    # )
    return raw_tables #applications.loc[:, NORMALIZED_APPLICATION_COLUMNS]


# def _normalize_exams(exams: pd.DataFrame, applications: pd.DataFrame) -> pd.DataFrame:
#     _require_columns(exams, BITRIX_EXAMS)
#     prepared = exams.copy()
#     prepared["deal_id"] = _as_text_series(prepared["iddeal"])
#     prepared["contact_id"] = _as_text_series(prepared["idcontact"])
#     prepared["exam_score"] = pd.to_numeric(prepared["ball"], errors="coerce")
#     prepared["exam_date"] = _as_datetime_series(prepared["date_testirovanya"])
#     prepared["is_active"] = prepared["aktive"].astype("boolean")
#     return prepared.merge(
#         applications.loc[:, ["deal_id", "program", "program_campus", "program_level"]],
#         how="left",
#         on="deal_id",
#         validate="many_to_one",
#     )


# def _normalize_portfolios(portfolios: pd.DataFrame, applications: pd.DataFrame) -> pd.DataFrame:
#     _require_columns(portfolios, BITRIX_PORTFOLIOS)
#     prepared = portfolios.copy()
#     prepared["deal_id"] = _as_text_series(prepared["iddeal"])
#     prepared["contact_id"] = _as_text_series(prepared["idcontact"])
#     prepared["product_id"] = _as_text_series(prepared["idtovar"])
#     prepared["is_active"] = prepared["status_elementa_portfolio"].astype("boolean")
#     return prepared.merge(
#         applications.loc[:, ["deal_id", "program", "program_campus", "program_level"]],
#         how="left",
#         on="deal_id",
#         validate="many_to_one",
#     )


# def normalize_bitrix_admissions_data(raw_tables: BitrixRawTables) -> NormalizedBitrixAdmissionsData:
#     """Normalize Bitrix admissions tables to dashboard-friendly tables."""

#     applications = _normalize_applications(raw_tables)
#     exams = _normalize_exams(raw_tables.exams, applications)
#     portfolios = _normalize_portfolios(raw_tables.portfolios, applications)
#     return NormalizedBitrixAdmissionsData(applications=applications, exams=exams, portfolios=portfolios)


# def create_bitrix_admissions_sources() -> tuple[BitrixItemSource, ...]: #entity_type_ids: Mapping[str, int]
#     """Create read-only Bitrix item sources for all admissions tables."""

#     # missing_names = [name for name in REQUIRED_BITRIX_TABLE_NAMES if name not in entity_type_ids]
#     # if missing_names:
#     #     raise ValueError(f"Missing Bitrix entity type IDs for admissions tables: {missing_names}")

#     entity_by_name: Mapping[str, BitrixEntity] = {
#         entity.name: entity
#         for entity in (
#             BITRIX_CRM_DEALS,
#             BITRIX_PORTAL_DEALS,
#             BITRIX_AISPK_APPLICATIONS,
#             BITRIX_CONTACTS,
#             BITRIX_EDUCATIONAL_PROGRAMS,
#             # BITRIX_CONTRACTS,
#             # BITRIX_EXAMS,
#             # BITRIX_PORTFOLIOS,
#         )
#     }
#     sources: list[BitrixItemSource] = []
#     for table_name in REQUIRED_BITRIX_TABLE_NAMES:
#         entity = entity_by_name[table_name]
#         select = entity.required_fields
#         # select = tuple(dict.fromkeys((*entity.required_fields, *OPTIONAL_ENTITY_SELECT_FIELDS.get(table_name, ())))) # TODO move back if needed
#         sources.append(
#             BitrixItemSource(
#                 name=table_name,
#                 entity_type_id=entity.num_id, # int(entity_type_ids[table_name]),
#                 select=select,
#                 extra_filter={"CATEGORY_ID" : entity.num_id} if table_name == "crm_deals" else {}, # TODO change to resolving CATEGORY_ID by name "Поступление 360"
#             )
#         )
#     return tuple(sources)

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
    sources: tuple[BitrixEntity, ...],
    batch_size: int,
    debug: bool = False
) -> BitrixRawTables:
    """Collect Bitrix admissions item sources into the raw-table container."""

    tables = collect_bitrix_item_sources(client, sources, batch_size, debug)
    missing_names = [name for name in REQUIRED_BITRIX_TABLE_NAMES if name not in tables]
    if missing_names:
        raise ValueError(f"Bitrix raw tables were not collected: {missing_names}")
    return BitrixRawTables(
        crm_deals=tables["crm_deals"],
        portal_deals=tables["portal_deals"],
        applications=tables["applications"],
        #contacts=tables["contacts"],
        educational_programs=tables["educational_programs"],
        # TODO in future
        # contracts=tables["contracts"],
        # exams=tables["exams"],
        # portfolios=tables["portfolios"],
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


# def _insert_metric_by_program_identity(
#     dashboard: pd.DataFrame,
#     metric_values: pd.DataFrame,
#     metric_column: str,
# ) -> pd.Series:
#     if metric_values.empty:
#         return pd.Series([0] * len(dashboard), index=dashboard.index)

#     left_keys = [col_program]
#     right_keys = ["program"]
#     optional_keys = (
#         ("campus", "program_campus"),
#         ("level", "program_level"),
#         ("format", "program_form"),
#     )
#     for left_key, right_key in optional_keys:
#         if left_key in dashboard.columns and right_key in metric_values.columns:
#             left_keys.append(left_key)
#             right_keys.append(right_key)

#     prepared_dashboard = dashboard.reset_index(names="__dashboard_index")
#     prepared_metrics = metric_values.copy()
#     for left_key, right_key in zip(left_keys, right_keys):
#         prepared_dashboard[left_key] = prepared_dashboard[left_key].astype("string").str.strip()
#         prepared_metrics[right_key] = prepared_metrics[right_key].astype("string").str.strip()

#     duplicated_metric_keys = prepared_metrics.duplicated(subset=right_keys, keep=False)
#     if duplicated_metric_keys.any():
#         duplicated_rows = prepared_metrics.loc[duplicated_metric_keys, right_keys].drop_duplicates().to_dict("records")
#         raise ValueError(
#             f"Dashboard keys {left_keys} do not disambiguate Bitrix programs for metric {metric_column!r}: "
#             f"{duplicated_rows}"
#         )

#     merged = prepared_dashboard.merge(
#         prepared_metrics.loc[:, [*right_keys, "values"]],
#         how="left",
#         left_on=left_keys,
#         right_on=right_keys,
#         validate="many_to_one",
#     )
#     merged = merged.sort_values("__dashboard_index")
#     return merged["values"].fillna(0).rename(metric_column)


def _finalize_dashboard_calculations(dashboard: pd.DataFrame) -> pd.DataFrame:
    result = dashboard #.copy()
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
    admissions_data: pd.DataFrame, # NormalizedBitrixAdmissionsData,
    as_of: datetime,
) -> pd.DataFrame:
    """Fill dashboard metric columns from normalized Bitrix admissions data."""

    result = dashboard #.copy()

    # Previous version:
    # crm_leads_by_program = admissions_data.crm_deals.groupby(col_program_bitrix)[leads_dates].count()
    # crm_leads_by_program = pd.DataFrame({col_program_bitrix:crm_leads_by_program.index, 'values':crm_leads_by_program.values})
    # result[col_leads] = insert_values(result, crm_leads_by_program, col_program_bitrix, col_leads)

    leads_counts = admissions_data.crm_deals.groupby(col_program_bitrix)[leads_dates].size()
    result[col_leads] = result[col_program_bitrix].map(leads_counts).fillna(0).astype(int)

    leads_count_after_april = admissions_data.crm_deals[admissions_data.crm_deals[leads_dates] >= DATE_01_04_2026].groupby(col_program_bitrix)[leads_dates].size()
    result[col_leads_after_april] = result[col_program_bitrix].map(leads_count_after_april).fillna(0).astype(int)

    leads_delta = admissions_data.crm_deals[admissions_data.crm_deals[leads_dates] >= datetime.now() - timedelta(days=3, hours=12)].groupby(col_program_bitrix)[leads_dates].size()
    result[col_leads_delta] = result[col_program_bitrix].map(leads_delta).fillna(0).astype(int)

    portal_counts = admissions_data.portal_deals.groupby(col_program_bitrix)[leads_dates].size()
    result[col_leads_partners] = result[col_program_bitrix].map(portal_counts).fillna(0).astype(int) 

    result[col_leads_total] = result[col_leads] + result[col_leads_partners]

    applications_count = admissions_data.applications.groupby(col_program_bitrix)[applications_dates].size()
    result[col_applications] = result[col_program_bitrix].map(applications_count).fillna(0).astype(int)

    contracts_count = admissions_data.applications[admissions_data.applications[contracts_dates].notna()].groupby(col_program_bitrix)[contracts_dates].size()
    result[col_contracts] = result[col_program_bitrix].map(contracts_count).fillna(0).astype(int)

    # TODO лиды по общему ленду, оплаты, зачисление, иностранцы, исторические выгрузки, проверки всех полей и сверка с выгрузками; возраста, МЖ, даты оплат, бэклог

    # Пока не работает:
    # leads_by_week = process_by_week(admissions_data.crm_deals, col_program_bitrix, leads_dates, 'count')
    # result[col_leads_by_week] = result[col_program_bitrix].map(leads_by_week).fillna(0).astype(int)

    # applications_by_week = process_by_week(admissions_data.applications, col_program_bitrix, applications_dates, 'count') # , "%Y-%m-%d"
    # result[col_applications_by_week] = result[col_program_bitrix].map(applications_by_week).fillna(0).astype(int)
    
    # contracts_by_week = process_by_week(admissions_data.applications, col_program_bitrix, contracts_dates, 'count')
    # result[col_contracts_by_week] = result[col_program_bitrix].map(contracts_by_week).fillna(0).astype(int)

    result.replace(np.inf, 0, inplace=True)
    result.fillna(0, inplace=True)
    return result


def process_current_files_from_bitrix(
    client: BitrixRestClient,
    sources: tuple[BitrixEntity, ...],
    dashboard_template: pd.DataFrame,
    as_of: datetime,
    batch_size: int,
    history_dataframes: list[pd.DataFrame],
    debug: bool = False
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build current dashboard data from Bitrix tables only."""

    raw_tables = collect_bitrix_raw_tables(client, sources, batch_size, debug)
    data = _normalize_raw_data(raw_tables)
    # admissions_data = normalize_bitrix_admissions_data(raw_tables)
    dashboard = apply_bitrix_metrics_to_dashboard(dashboard_template, data, as_of)
    dashboard = _finalize_dashboard_calculations(dashboard)
    dashboard = dashboard.drop(columns=[column for column in TECHNICAL_DASHBOARD_COLUMNS if column in dashboard.columns])
    
    #TODO add history_data processing and general columns in dashboard
    return dashboard, history_dataframes[0].copy()

