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


def get_deal_category_id(category_name: str, client: BitrixRestClient) -> int:
    '''Return Bitrix deal category ID by its visible funnel name.'''

    try:
        response = client.call('crm.category.list', {'entityTypeId': BITRIX_DEAL_ENTITY_TYPE_ID})
        category_result = response.get('result', [])
        categories = category_result.get('categories', []) if isinstance(category_result, Mapping) else category_result
    except BitrixAPIError:
        response = client.call(
            'crm.dealcategory.list',
            {'order': {'SORT': 'ASC'}, 'select': ['ID', 'NAME', 'SORT']},
        )
        categories = response.get('result', [])

    for category in categories:
        category_title = str(category.get('name') or category.get('NAME') or '').strip().casefold()
        if category_title == category_name.strip().casefold():
            return int(category.get('id') or category.get('ID'))
    raise ValueError(f'Deal funnel {category_name!r} was not found')

# Get all lists with 
# base_params: dict[str, Any] = {
#     'IBLOCK_TYPE_ID' : 'lists',
# }
# return _list_dataframe(client, 'lists.get', 'items', base_params, batch_size)
# 
# [{key: d[key]} for d in first_result for key in ['ID', 'CODE', 'API_CODE', 'NAME']]: 
# [{'ID': '43'}, {'CODE': None}, {'API_CODE': None}, {'NAME': 'Сферы интересов'}, {'ID': '39'}, {'CODE': None}, {'API_CODE': None}, {'NAME': 'Категория пользователя'}, {'ID': '38'}, {'CODE': None}, {'API_CODE': None}, {'NAME': 'Совокупность конкурсных групп'}, {'ID': '35'}, {'CODE': 'rannee priglashenie'}, {'API_CODE': 'ranneepriglashenie'}, {'NAME': 'Раннее приглашение'}, {'ID': '26'}, {'CODE': 'formy oprosa na portale'}, {'API_CODE': 'formyoprosanaportale'}, {'NAME': 'Формы опроса на портале'}, {'ID': '24'}, {'CODE': 'forma obucheniya'}, {'API_CODE': 'formaobucheniya'}, {'NAME': 'Форма обучения'}, {'ID': '23'}, {'CODE': 'kampusy'}, {'API_CODE': 'kampusy'}, {'NAME': 'Кампусы'}, {'ID': '22'}, {'CODE': 'fakultety'}, {'API_CODE': 'fakultety'}, {'NAME': 'Факультеты'}, {'ID': '21'}, {'CODE': 'obrazovatelnye programmy'}, {'API_CODE': 'obrazovatelnyeprogrammy'}, {'NAME': 'Образовательные программы'}, {'ID': '20'}, {'CODE': 'urovni obrazovaniya'}, {'API_CODE': 'urovniobrazovaniya'}, {'NAME': 'Уровни образования'}, {'ID': '18'}, {'CODE': 'strany'}, {'API_CODE': 'strany'}, {'NAME': 'Страны'}, {'ID': '17'}, {'CODE': 'goroda'}, {'API_CODE': 'goroda'}, {'NAME': 'Города'}, {'ID': '16'}, {'CODE': 'nabor na uchebnyj god'}, {'API_CODE': 'nabornauchebnyjgod'}, {'NAME': 'Набор на учебный год'}, {'ID': '5'}, {'CODE': 'clients_s1'}, {'API_CODE': None}, {'NAME': 'Клиенты'}]



def collect_deals_dataframe(
    client: BitrixRestClient,
    category_name: str,
    category_id: int,
    select: Sequence[str],
    extra_filter: Mapping[str, Any] | None,
    batch_size: int,
) -> pd.DataFrame:
    '''Collect all deals from a Bitrix CRM funnel into one DataFrame.'''

    if category_id is None:
        raise('No category ID') # get_deal_category_id(category_name, client)

    deal_filter: dict[str, Any] = {'CATEGORY_ID': category_id}
    if extra_filter:
        deal_filter.update(extra_filter)
    base_params: dict[str, Any] = {
        'select': list(select),
        'filter': deal_filter,
        'order': {'ID': 'ASC'},
    }
    return _list_dataframe(client, 'crm.deal.list', 'result', base_params, batch_size)


# TODO unused? only for testing?
def collect_360_deals_dataframe(
    client: BitrixRestClient | None = None,
    category_name: str = 'Поступление 360',
    category_id: int | None = None,
    select: Sequence[str] = DEFAULT_DEAL_SELECT,
    extra_filter: Mapping[str, Any] | None = None,
    batch_size: int = BITRIX_BATCH_LIMIT,
) -> pd.DataFrame:
    '''Compatibility helper for the admissions funnel deal export.'''

    rest = client if client is not None else create_bitrix_client(BITRIX_WEBHOOK_URL)
    return collect_deals_dataframe(rest, category_name, category_id, select, extra_filter, batch_size)


# TODO move if to BitrixEntity setting
def collect_crm_items_dataframe(
    client: BitrixRestClient,
    entity_type_id: int,
    select: Sequence[str],
    extra_filter: Mapping[str, Any],
    batch_size: int,
    debug: bool = False
) -> pd.DataFrame:
    '''Collect Bitrix dynamic CRM items by entity type ID.'''

    if entity_type_id in [1, 2, 3, 4, 5, 31, 7, 8, 36, 39]:
        if debug:
            return pd.DataFrame() # for debug
        rest_request = 'crm.item.list'
        base_params: dict[str, Any] = {
            'entityTypeId': entity_type_id,
            'select': ['*'], # TODO list(select), now for the case of problem
            'filter': dict(extra_filter),
            'order': {'id': 'ASC'},
        }
        
    else: 
        rest_request = 'lists.element.get'
        base_params: dict[str, Any] = {
            'IBLOCK_TYPE_ID' : 'lists',
            'IBLOCK_ID': entity_type_id,
            # 'SELECT' : ['*'],
            # 'FILTER' : dict(extra_filter), 
            'ELEMENT_ORDER': {'id': 'ASC'},
        }

        
    return _list_dataframe(client, rest_request, 'items', base_params, batch_size, debug)
# # TODO make list+get (faster version) https://habr.com/ru/articles/537694/

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

def _as_text_series(series: pd.Series) -> pd.Series:
    return series.astype("string").str.strip()


def _as_datetime_series(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="raise", dayfirst=True) # may be coerce?


def _payment_dates_by_deal(contracts: pd.DataFrame) -> pd.DataFrame:
    _require_columns(contracts, BITRIX_CONTRACTS)
    prepared = contracts.loc[:, ["iddeal", "data_oplaty"]].copy()
    prepared["deal_id"] = _as_text_series(prepared["iddeal"])
    prepared["payment_date"] = _as_datetime_series(prepared["data_oplaty"])
    prepared = prepared.dropna(subset=["deal_id", "payment_date"])
    if prepared.empty:
        return pd.DataFrame(columns=["deal_id", "payment_date"])
    return prepared.groupby("deal_id", as_index=False)["payment_date"].min()

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


def create_bitrix_admissions_sources() -> tuple[BitrixItemSource, ...]: #entity_type_ids: Mapping[str, int]
    """Create read-only Bitrix item sources for all admissions tables."""

    # missing_names = [name for name in REQUIRED_BITRIX_TABLE_NAMES if name not in entity_type_ids]
    # if missing_names:
    #     raise ValueError(f"Missing Bitrix entity type IDs for admissions tables: {missing_names}")

    entity_by_name: Mapping[str, BitrixEntity] = {
        entity.name: entity
        for entity in (
            BITRIX_CRM_DEALS,
            BITRIX_PORTAL_DEALS,
            BITRIX_AISPK_APPLICATIONS,
            BITRIX_CONTACTS,
            BITRIX_EDUCATIONAL_PROGRAMS,
            # BITRIX_CONTRACTS,
            # BITRIX_EXAMS,
            # BITRIX_PORTFOLIOS,
        )
    }
    sources: list[BitrixItemSource] = []
    for table_name in REQUIRED_BITRIX_TABLE_NAMES:
        entity = entity_by_name[table_name]
        select = entity.required_fields
        # select = tuple(dict.fromkeys((*entity.required_fields, *OPTIONAL_ENTITY_SELECT_FIELDS.get(table_name, ())))) # TODO move back if needed
        sources.append(
            BitrixItemSource(
                name=table_name,
                entity_type_id=entity.num_id, # int(entity_type_ids[table_name]),
                select=select,
                extra_filter={"CATEGORY_ID" : entity.num_id} if table_name == "crm_deals" else {}, # TODO change to resolving CATEGORY_ID by name "Поступление 360"
            )
        )
    return tuple(sources)

def collect_deals_dataframe(
    *,
    webhook_url: str = BITRIX_WEBHOOK_URL,
    category_name: str = "Поступление 360",
    category_id: int | None = None,
    select: Sequence[str] = DEFAULT_DEAL_SELECT,
    extra_filter: Mapping[str, Any] | None = None,
    client: BitrixRestClient | None = None,
    batch_size: int = BITRIX_BATCH_LIMIT,
) -> pd.DataFrame:
    """Collect all deals from the Bitrix CRM funnel "Поступление 360" into one DataFrame.

    The function uses only read-only REST methods: ``crm.category.list``,
    ``crm.deal.list`` and read-only subcommands inside ``batch``. Pagination is
    batched in groups of up to 50 commands, and the client throttles outbound
    HTTP calls to two requests per second by default.
    """

    rest = client or BitrixRestClient(webhook_url)
    resolved_category_id = (
        category_id if category_id is not None else get_deal_category_id(category_name, client=rest)
    )
    deal_filter: dict[str, Any] = {"CATEGORY_ID": resolved_category_id}
    if extra_filter:
        deal_filter.update(extra_filter)
    base_params: dict[str, Any] = {
        "select": list(select),
        "filter": deal_filter,
        "order": {"ID": "ASC"},
    }

    first_response = rest.call("crm.deal.list", {**base_params, "start": 0})
    deals = list(first_response.get("result", []))
    total = int(first_response.get("total", len(deals)))
    if total <= BITRIX_PAGE_SIZE:
        return pd.DataFrame(deals)

    max_batch_size = max(1, min(batch_size, BITRIX_BATCH_LIMIT))
    starts = list(range(BITRIX_PAGE_SIZE, total, BITRIX_PAGE_SIZE)) #TODO test "-1"

    total_batches = math.ceil(len(starts) / max_batch_size)
    for batch_index in range(total_batches):
        chunk_starts = starts[batch_index * max_batch_size : (batch_index + 1) * max_batch_size]
        commands = {
            f"deals_{start}": ("crm.deal.list", {**base_params, "start": start})
            for start in chunk_starts
        }
        for page in rest.batch(commands).values():
            deals.extend(page or [])

    return pd.DataFrame(deals)
