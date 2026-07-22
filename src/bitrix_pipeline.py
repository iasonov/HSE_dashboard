'''Normalize Bitrix admissions tables and calculate dashboard-ready metrics.'''

from __future__ import annotations

import numpy as np
import pandas as pd

from dataclasses import dataclass

from time_const import *
from bitrix import BITRIX_BATCH_LIMIT, BitrixRestClient, collect_bitrix_item_sources
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
from bitrix_contracts import (
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
from general_pipeline import process_by_week, categorize_ages, num_years


@dataclass(slots=True)
class BitrixRawTables:
    '''Raw Bitrix admissions tables exported from CRM.'''

    crm_deals: pd.DataFrame
    portal_deals: pd.DataFrame
    applications: pd.DataFrame
    # contacts: pd.DataFrame
    educational_programs: pd.DataFrame
    # TODO uncomment
    # contracts: pd.DataFrame
    # exams: pd.DataFrame
    # portfolios: pd.DataFrame

REQUIRED_BITRIX_TABLE_NAMES = tuple(item.name for item in BITRIX_ADMISSIONS_ENTITIES)
TECHNICAL_DASHBOARD_COLUMNS = ('program_bitrix', 'tg_chat_id', 'campus', 'start_year', 'format')
NEEDED_APPLICATIONS_RATIO = 45 / 100
MALE_VALUES = ('Муж.', 'ÐœÑƒÐ¶.')
FEMALE_VALUES = ('Жен.', 'Ð–ÐµÐ½.')


def _require_columns(frame: pd.DataFrame, entity: BitrixEntity) -> None:
    missing_columns = [column for column in entity.select if column not in frame.columns]
    if missing_columns:
        raise ValueError(f'Bitrix table {entity.name!r} is missing required columns: {missing_columns}')


def _normalize_raw_data(raw_tables: BitrixRawTables) -> BitrixRawTables:
    # _require_columns(raw_tables.crm_deals, BITRIX_CRM_DEALS)
    # _require_columns(raw_tables.portal_deals, BITRIX_PORTAL_DEALS)
    # _require_columns(raw_tables.applications, BITRIX_APPLICATIONS)
    # _require_columns(raw_tables.contacts, BITRIX_CONTACTS)
    # _require_columns(raw_tables.educational_programs, BITRIX_EDUCATIONAL_PROGRAMS)

    # TODO перенести все строковые константы в col_names.py

    raw_tables.crm_deals['ufDealEducationProgram']    = raw_tables.crm_deals['ufDealEducationProgram'].fillna(0).astype(int, errors='raise') # TODO repair, пустые - преимущественно, но не только разводящий лендинг, но и другие программы (см. вкладку Пустоты в ufDealProgram в файле 2026_06_18_..xlsx)
    raw_tables.portal_deals['ufDealEducationProgram'] = raw_tables.portal_deals['ufDealEducationProgram'].fillna(-1).astype(int, errors='raise') # TODO check -1 - не наша программа
    raw_tables.applications['ufDealEducationProgram'] = raw_tables.applications['ufDealEducationProgram'].fillna(-2).astype(int, errors='raise') # TODO repair - там и наши ленды, и не наши ленды, см. Пустоты в ufDealProgram_APP

    raw_tables.educational_programs         = raw_tables.educational_programs[['ID', 'NAME']] # убираем ненужные столбцы, альтернативно можно не забирать их с помощью SELECT
    raw_tables.educational_programs['ID']   = raw_tables.educational_programs['ID'].astype(int, errors='raise') # в целом тут пустых быть не должно
    raw_tables.educational_programs         = raw_tables.educational_programs.rename(columns={'ID': 'ufDealEducationProgram', 'NAME': col_program_bitrix}) # переименовываем столбцы для удобства merge
    raw_tables.educational_programs         = pd.concat([raw_tables.educational_programs, pd.DataFrame({
        'ufDealEducationProgram': [0,                -1,                                           -2],
        col_program_bitrix:       [main_studyonline, 'Не указана программа в воронке Портала ВШЭ', 'Не указана программа в воронке МАГ/БАК']
    })], ignore_index=True)

    raw_tables.crm_deals    = pd.merge(raw_tables.crm_deals,    raw_tables.educational_programs, on='ufDealEducationProgram', how='left').drop(columns=['ufDealEducationProgram'])
    raw_tables.portal_deals = pd.merge(raw_tables.portal_deals, raw_tables.educational_programs, on='ufDealEducationProgram', how='left').drop(columns=['ufDealEducationProgram'])
    raw_tables.applications = pd.merge(raw_tables.applications, raw_tables.educational_programs, on='ufDealEducationProgram', how='left').drop(columns=['ufDealEducationProgram'])

    raw_tables.crm_deals[leads_dates]           = pd.to_datetime(raw_tables.crm_deals['createdTime'], errors='raise').dt.tz_localize(None)
    raw_tables.portal_deals[leads_dates]        = pd.to_datetime(raw_tables.portal_deals['createdTime'], errors='raise').dt.tz_localize(None)
    raw_tables.applications[applications_dates] = pd.to_datetime(raw_tables.applications['createdTime'], errors='raise').dt.tz_localize(None)
    raw_tables.applications[contracts_dates]    = pd.to_datetime(raw_tables.applications['ufDealContractdate'], errors='raise').dt.tz_localize(None) # TODO check

    raw_tables.applications['ufDealFinancing'] = raw_tables.applications['ufDealFinancing'].fillna(807).astype(int)

    return raw_tables


def _collect_bitrix_raw_tables(
    client: BitrixRestClient,
    sources: tuple[BitrixEntity, ...],
    batch_size: int,
    debug: bool = False
) -> BitrixRawTables:
    '''Collect Bitrix admissions item sources into the raw-table container.'''

    tables = collect_bitrix_item_sources(client, sources, batch_size, debug)
    missing_names = [name for name in REQUIRED_BITRIX_TABLE_NAMES if name not in tables]
    if missing_names:
        raise ValueError(f'Bitrix raw tables were not collected: {missing_names}')
    return BitrixRawTables(
        crm_deals=tables['crm_deals'],
        portal_deals=tables['portal_deals'],
        applications=tables['applications'],
        #contacts=tables['contacts'],
        educational_programs=tables['educational_programs'],
        # TODO in future
        # contracts=tables['contracts'],
        # exams=tables['exams'],
        # portfolios=tables['portfolios'],
    )


def _count_by_program(frame: pd.DataFrame, date_column: str) -> pd.DataFrame:
    filtered = frame.dropna(subset=['program', date_column])
    group_columns = ['program', 'program_campus', 'program_level', 'program_form']
    return filtered.groupby(group_columns, dropna=False)['program'].count().reset_index(name='values')


def _count_present_by_program(frame: pd.DataFrame, value_column: str) -> pd.DataFrame:
    filtered = frame.dropna(subset=['program'])
    filtered = filtered[filtered[value_column].notna() & (filtered[value_column].astype('string').str.len() > 0)]
    group_columns = ['program', 'program_campus', 'program_level', 'program_form']
    return filtered.groupby(group_columns, dropna=False)['program'].count().reset_index(name='values')


def _gender_count_by_program(frame: pd.DataFrame, gender_values: tuple[str, ...]) -> pd.DataFrame:
    filtered = frame.dropna(subset=['program', 'gender'])
    group_columns = ['program', 'program_campus', 'program_level', 'program_form']
    return (
        filtered[filtered['gender'].isin(gender_values)]
        .groupby(group_columns, dropna=False)['program']
        .count()
        .reset_index(name='values')
    )


def _age_bars_by_program(frame: pd.DataFrame, as_of: datetime) -> pd.DataFrame:
    filtered = frame.dropna(subset=['program', 'birthdate']).copy()
    filtered['age'] = filtered['birthdate'].apply(lambda birthdate: num_years(birthdate, as_of))
    group_columns = ['program', 'program_campus', 'program_level', 'program_form']
    return filtered.groupby(group_columns, dropna=False)['age'].apply(categorize_ages).reset_index(name='values')


def _age_mean_by_program(frame: pd.DataFrame, as_of: datetime) -> pd.DataFrame:
    filtered = frame.dropna(subset=['program', 'birthdate']).copy()
    filtered['age'] = filtered['birthdate'].apply(lambda birthdate: num_years(birthdate, as_of))
    group_columns = ['program', 'program_campus', 'program_level', 'program_form']
    return filtered.groupby(group_columns, dropna=False)['age'].mean().reset_index(name='values')


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
    result[col_income_1year] = result['price'] * result[col_payments] / 1000
    # TODO income calculations
    # result.loc[result['level'] == 'master', col_income_all] = result[col_income_1year] * 2
    # result.loc[result['level'] == 'bachelor', col_income_all] = result[col_income_1year] * 4
    # result[col_income_1year_hse] = result[col_income_1year] * result['income_percent'] / 100
    # result[col_income_all_hse] = result[col_income_all] * result['income_percent'] / 100
    if col_enrollments_foreign not in result.columns:
        result[col_enrollments_foreign] = 0

    result.replace(np.inf, 0, inplace=True)
    result.fillna(0, inplace=True)

    float_columns = result.select_dtypes('float64')
    result[float_columns.columns] = float_columns.astype(int, errors='raise')
    return result


def _apply_bitrix_metrics_to_dashboard(
    dashboard: pd.DataFrame,
    admissions_data: pd.DataFrame, # NormalizedBitrixAdmissionsData,
    as_of: datetime,
) -> pd.DataFrame:
    '''Fill dashboard metric columns from normalized Bitrix admissions data.'''

    result = dashboard #.copy()

    # Previous version:
    # crm_leads_by_program = admissions_data.crm_deals.groupby(col_program_bitrix)[leads_dates].count()
    # crm_leads_by_program = pd.DataFrame({col_program_bitrix:crm_leads_by_program.index, 'values':crm_leads_by_program.values})
    # result[col_leads] = insert_values(result, crm_leads_by_program, col_program_bitrix, col_leads)

    main_leads = admissions_data.crm_deals[admissions_data.crm_deals[col_program_bitrix] == main_studyonline][col_program_bitrix].count()
    main_leads_after_april = admissions_data.crm_deals[(admissions_data.crm_deals[col_program_bitrix] == main_studyonline) & (admissions_data.crm_deals[leads_dates] >= DATE_01_04_2026)][col_program_bitrix].count()

    main_leads_counts = pd.DataFrame(columns=result.columns, data=[{col_leads: main_leads, col_leads_after_april: main_leads_after_april, col_program_bitrix: main_studyonline, col_program: main_studyonline}])
    result = pd.concat([main_leads_counts, result], ignore_index=True)

    leads_counts = admissions_data.crm_deals.groupby(col_program_bitrix)[leads_dates].size()
    result[col_leads] = result[col_program_bitrix].map(leads_counts).fillna(0).astype(int)

    leads_count_after_april = admissions_data.crm_deals[admissions_data.crm_deals[leads_dates] >= DATE_01_04_2026].groupby(col_program_bitrix)[leads_dates].size()
    result[col_leads_after_april] = result[col_program_bitrix].map(leads_count_after_april).fillna(0).astype(int)

    leads_delta = admissions_data.crm_deals[admissions_data.crm_deals[leads_dates] >= datetime.now() - timedelta(days=3, hours=12)].groupby(col_program_bitrix)[leads_dates].size()
    result[col_leads_delta] = result[col_program_bitrix].map(leads_delta).fillna(0).astype(int)

    portal_counts = admissions_data.portal_deals.groupby(col_program_bitrix)[leads_dates].size()
    result[col_leads_partners] = result[col_program_bitrix].map(portal_counts).fillna(0).astype(int)

    result[col_leads_total] = result[col_leads] + result[col_leads_partners]

    applications_count = admissions_data.applications[admissions_data.applications['ufDealFinancing'] == 808].groupby(col_program_bitrix)[applications_dates].size()
    result[col_applications] = result[col_program_bitrix].map(applications_count).fillna(0).astype(int)

    applications_count_budget = admissions_data.applications[admissions_data.applications['ufDealFinancing'] == 807].groupby(col_program_bitrix)[applications_dates].size()
    result[col_applications_budget] = result[col_program_bitrix].map(applications_count_budget).fillna(0).astype(int)

    applications_delta = admissions_data.applications[(admissions_data.applications['ufDealFinancing'] == 808) & (admissions_data.applications[applications_dates] >= datetime.now() - timedelta(days=3, hours=12))].groupby(col_program_bitrix)[applications_dates].size()
    result[col_applications_delta] = result[col_program_bitrix].map(applications_delta).fillna(0).astype(int)

    contracts_count = admissions_data.applications[admissions_data.applications['ufDealNomerDogovora'].notna()].groupby(col_program_bitrix)['ufDealNomerDogovora'].size()
    result[col_contracts] = result[col_program_bitrix].map(contracts_count).fillna(0).astype(int)

    payments_count = admissions_data.applications[admissions_data.applications['ufDealDogovorOplachen'] == 'Y'].groupby(col_program_bitrix)['ufDealDogovorOplachen'].size()
    result[col_payments] = result[col_program_bitrix].map(payments_count).fillna(0).astype(int)

    # TODO зачисление, иностранцы, исторические выгрузки, проверки всех полей и сверка с выгрузками; возраста, МЖ, даты оплат, бэклог

    leads_by_week = process_by_week(admissions_data.crm_deals, col_program_bitrix, leads_dates)
    result[col_leads_by_week] = result[col_program_bitrix].map(leads_by_week).fillna('').astype(str)

    applications_by_week = process_by_week(admissions_data.applications[admissions_data.applications['ufDealFinancing'] == 808], col_program_bitrix, applications_dates, pd.Timestamp(year=2026, month=6, day=15, hour=0, minute=0, second=0)) # , "%Y-%m-%d"
    result[col_applications_by_week] = result[col_program_bitrix].map(applications_by_week).fillna("").astype(str)

    contracts_by_week = process_by_week(admissions_data.applications[admissions_data.applications['ufDealFinancing'] == 808], col_program_bitrix, contracts_dates, pd.Timestamp(year=2026, month=6, day=15, hour=0, minute=0, second=0)) # pd.Series() # TODO no contracts_dates, unfortunatly 
    result[col_contracts_by_week] = result[col_program_bitrix].map(contracts_by_week).fillna("").astype(str)


    result.replace(np.inf, 0, inplace=True)
    result.fillna(0, inplace=True)
    return result

def _apply_history_data_to_dashboard(
    dashboard: pd.DataFrame,
    data: BitrixRawTables,
    history_data: pd.DataFrame,
    leads_prev: pd.DataFrame,
    leads_after_april_prev: pd.DataFrame,
    applications_prev: pd.DataFrame,
    contracts_prev: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    '''Обновляет исторические метрики и подготавливает строку main_studyonline для дашборда.'''

    # TODO check
    online_masters_ids = dashboard[dashboard['level'] == 'master'][col_program_bitrix]
    online_bachelors_ids = dashboard[dashboard['level'] == 'bachelor'][col_program_bitrix]
    online_no_rossokhins_ids = dashboard[~dashboard[col_program].str.startswith('Психоанализ и')][col_program_bitrix]
    online_masters_no_rossokhins_ids = dashboard[(~dashboard[col_program].str.startswith('Психоанализ и'))&(dashboard['level'] == 'master')][col_program_bitrix]
    budget_ids = data.applications[data.applications['ufDealFinancing'] == 807]['contactId'].unique()


    # Filter applications to only include those with financing type 808 (contract) and have no budget applications
    
    applications_contracts_no_budget = data.applications[(data.applications['ufDealFinancing'] == 808) & (~data.applications['contactId'].isin(budget_ids))].copy()

    history_data.loc[2026, 'applications_unique_no_budget'] = data.applications[~data.applications['contactId'].isin(budget_ids)]['contactId'].drop_duplicates().count()
    history_data.loc[2026, 'applications_masters_unique_no_budget'] = applications_contracts_no_budget[applications_contracts_no_budget[col_program_bitrix].isin(online_masters_ids)]['contactId'].drop_duplicates().count()
    history_data.loc[2026, 'applications_bachelors_unique_no_budget'] = applications_contracts_no_budget[applications_contracts_no_budget[col_program_bitrix].isin(online_bachelors_ids)]['contactId'].drop_duplicates().count()
    history_data.loc[2026, 'applications_masters_no_rossokhins_unique_no_budget'] = applications_contracts_no_budget[applications_contracts_no_budget[col_program_bitrix].isin(online_masters_no_rossokhins_ids)]['contactId'].drop_duplicates().count()
    history_data.loc[2026, 'applications_no_rossokhins_unique_no_budget'] = applications_contracts_no_budget[applications_contracts_no_budget[col_program_bitrix].isin(online_no_rossokhins_ids)]['contactId'].drop_duplicates().count()

    # Filter applications to only include those with financing type 808 (contract)

    history_data.loc[2026, 'applications_unique'] = data.applications[data.applications['ufDealFinancing'] == 808]['contactId'].drop_duplicates().count() # TODO check

    applications_contract = data.applications[data.applications['ufDealFinancing'] == 808].copy()
    history_data.loc[2026, 'applications_masters_unique'] = applications_contract[applications_contract[col_program_bitrix].isin(online_masters_ids)]['contactId'].drop_duplicates().count()
    history_data.loc[2026, 'applications_bachelors_unique'] = applications_contract[applications_contract[col_program_bitrix].isin(online_bachelors_ids)]['contactId'].drop_duplicates().count()

    history_data.loc[2026, 'applications_no_rossokhins_unique'] = applications_contract[applications_contract[col_program_bitrix].isin(online_no_rossokhins_ids)]['contactId'].drop_duplicates().count()

    history_data.loc[2026, 'applications_masters_no_rossokhins_unique'] = applications_contract[applications_contract[col_program_bitrix].isin(online_masters_no_rossokhins_ids)]['contactId'].drop_duplicates().count()

    # list(set(online_masters_ids) - set([219, 220])) # две психологии Россохина


    history_data.loc[2026, 'early_invitations_unique'] = 1658 # TODO исправить на расчет


    # leads_prev_df = pd.Series(leads_prev.values, index=# ({col_program_bitrix: leads_prev.index, 'values': leads_prev.values})
    # leads_after_april_prev_df = pd.DataFrame({col_program_bitrix: leads_after_april_prev.index, 'values': leads_after_april_prev.values})

    dashboard[col_leads_prev] = dashboard[col_program_bitrix].map(leads_prev).fillna(0).astype(int)
    dashboard[col_leads_after_april_prev] = dashboard[col_program_bitrix].map(leads_after_april_prev).fillna(0).astype(int)

    # result[col_leads] = result[col_program_bitrix].map(leads_counts).fillna(0).astype(int)
    try:
        main_leads_after_april_prev = leads_after_april_prev[main_studyonline]
    except:
        main_leads_after_april_prev = 0

    dashboard.loc[dashboard[col_program_bitrix] == main_studyonline, col_leads_after_april_prev] = main_leads_after_april_prev

    try:
        main_leads_prev = leads_prev[main_studyonline]
    except:
        main_leads_prev = 0

    dashboard.loc[dashboard[col_program_bitrix] == main_studyonline, col_leads_prev] = main_leads_prev


    dashboard[col_applications_prev] = dashboard[col_program].map(applications_prev).fillna(0).astype(int) # TODO check map
    dashboard[col_contracts_prev] = dashboard[col_program].map(contracts_prev).fillna(0).astype(int)

    # applications_prev_df = pd.DataFrame({col_program: applications_prev.index, 'values': applications_prev.values})
    # contracts_prev_df = pd.DataFrame({col_program: contracts_prev.index, 'values': contracts_prev.values})

    # # # 6. Создание строки для main_studyonline в дашборде (аналог df_main_dashboard)
    # main_dashboard_row = pd.DataFrame([{
    #     col_program: main_studyonline,
    #     col_program_bitrix: main_studyonline,
    #     col_leads_prev: main_leads_prev,
    #     col_leads_after_april_prev: main_leads_after_april_prev
    # }])

    # Сохраняем рассчитанные метрики в history_data для последующей вставки в дашборд
    # (slots=True в BitrixRawTables не позволяет хранить их в data напрямую)
    # history_data.loc[2026, col_leads_prev] = main_leads_prev
    # history_data.loc[2026, col_leads_after_april_prev] = main_leads_after_april_prev
    # # history_data['main_dashboard_row'] = main_dashboard_row
    # history_data[col_applications_prev] = applications_prev_df
    # history_data[col_contracts_prev] = contracts_prev_df

    return dashboard, history_data

def process_from_bitrix(debug: bool = False) -> tuple[pd.DataFrame, pd.DataFrame]:
    '''Build current dashboard data from Bitrix tables only.'''
    from bitrix import BITRIX_BATCH_LIMIT, BITRIX_WEBHOOK_URL, create_bitrix_client
    from files_pipeline import load_dashboard_template, process_history_files

    templates_folder = 'templates/'
    dashboard_template = load_dashboard_template(templates_folder)
    client = create_bitrix_client(BITRIX_WEBHOOK_URL)
    sources = BITRIX_ADMISSIONS_ENTITIES # create_bitrix_admissions_sources() # bitrix_entity_type_ids
    if not debug:
        history_data, leads_prev, leads_after_april_prev, applications_prev, contracts_prev = process_history_files() # TODO df_pivot, df_leads_all_prev, df_leads_after_april_prev, df_applications_prev, df_contracts_prev
    else:
        history_data, leads_prev, leads_after_april_prev, applications_prev, contracts_prev = [pd.DataFrame() for _ in range(5)]

    raw_tables = _collect_bitrix_raw_tables(client, sources, BITRIX_BATCH_LIMIT, debug)
    data = _normalize_raw_data(raw_tables)
    dashboard = _apply_bitrix_metrics_to_dashboard(dashboard_template, data, datetime.now())
    dashboard, history_data = _apply_history_data_to_dashboard(dashboard, data, history_data, leads_prev, leads_after_april_prev, applications_prev, contracts_prev)
    dashboard = _finalize_dashboard_calculations(dashboard)
    dashboard = dashboard.drop(columns=[column for column in TECHNICAL_DASHBOARD_COLUMNS if column in dashboard.columns])

    #TODO add general columns in dashboard
    return dashboard, history_data

