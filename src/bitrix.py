"""Read-only helpers for loading Bitrix24 CRM data into pandas."""

from __future__ import annotations

import math
import time
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Any
from urllib.parse import urlencode

import pandas as pd

try:
    from my_secrets import secrets
except ImportError:
    secrets = {}

BITRIX_WEBHOOK_URL = str(secrets.get("BITRIX_WEBHOOK_URL", ""))
BITRIX_DEAL_ENTITY_TYPE_ID = 2
BITRIX_PAGE_SIZE = 50
BITRIX_BATCH_LIMIT = 50
READ_ONLY_METHOD_SUFFIXES = (".get", ".list", ".fields")
READ_ONLY_METHODS = {"batch"}
DEFAULT_DEAL_SELECT = (
    "ID",
    "TITLE",
    "TYPE_ID",
    "CATEGORY_ID",
    "STAGE_ID",
    "STAGE_SEMANTIC_ID",
    "IS_NEW",
    "IS_RECURRING",
    "IS_RETURN_CUSTOMER",
    "IS_REPEATED_APPROACH",
    "OPPORTUNITY",
    "CURRENCY_ID",
    "ASSIGNED_BY_ID",
    "CONTACT_ID",
    "COMPANY_ID",
    "DATE_CREATE",
    "DATE_MODIFY",
    "BEGINDATE",
    "CLOSEDATE",
    "SOURCE_ID",
    "SOURCE_DESCRIPTION",
    "COMMENTS",
    "UF_*",
)


@dataclass(frozen=True, slots=True)
class BitrixItemSource:
    """Describe one Bitrix dynamic CRM item source."""

    name: str
    entity_type_id: int
    select: tuple[str, ...]
    extra_filter: Mapping[str, Any]


class BitrixReadOnlyError(ValueError):
    """Raised when a non-read-only REST method is requested."""


class BitrixAPIError(RuntimeError):
    """Raised when Bitrix REST returns an API-level error."""


def _is_read_only_method(method: str) -> bool:
    return method in READ_ONLY_METHODS or method.endswith(READ_ONLY_METHOD_SUFFIXES)


def _assert_read_only_method(method: str) -> None:
    if not _is_read_only_method(method):
        raise BitrixReadOnlyError(
            f"Method {method!r} is not allowed: only read-only Bitrix REST methods are used"
        )


def _flatten_params(params: Mapping[str, Any], prefix: str | None) -> list[tuple[str, Any]]:
    """Convert nested Bitrix REST parameters to bracket-notation pairs."""

    pairs: list[tuple[str, Any]] = []
    for key, value in params.items():
        nested_key = f"{prefix}[{key}]" if prefix else str(key)
        if isinstance(value, Mapping):
            pairs.extend(_flatten_params(value, nested_key))
        elif isinstance(value, (list, tuple)):
            for item in value:
                if isinstance(item, Mapping):
                    pairs.extend(_flatten_params(item, f"{nested_key}[]"))
                else:
                    pairs.append((f"{nested_key}[]", item))
        elif value is not None:
            pairs.append((nested_key, value))
    return pairs


def _build_batch_command(method: str, params: Mapping[str, Any]) -> str:
    _assert_read_only_method(method)
    query = urlencode(_flatten_params(params, None), doseq=True)
    return f"{method}?{query}" if query else method


class BitrixRestClient:
    """Minimal read-only Bitrix REST client with rate limiting and retries."""

    def __init__(
        self,
        webhook_url: str,
        *,
        requests_per_second: float,
        timeout: int,
        max_retries: int,
        retry_backoff: float,
    ) -> None:
        self.webhook_url = webhook_url.rstrip("/") + "/"
        self.timeout = timeout
        self.max_retries = max_retries
        self.retry_backoff = retry_backoff
        self._min_interval = 1 / requests_per_second if requests_per_second > 0 else 0
        self._last_request_at = 0.0
        try:
            import requests
        except ImportError as error:
            raise RuntimeError("Install the project dependencies before using BitrixRestClient") from error
        self.session = requests.Session()

    def _wait_for_rate_limit(self) -> None:
        if self._min_interval <= 0:
            return
        elapsed = time.monotonic() - self._last_request_at
        if elapsed < self._min_interval:
            time.sleep(self._min_interval - elapsed)

    def call(self, method: str, params: Mapping[str, Any] | None) -> dict[str, Any]:
        """Call a read-only Bitrix REST method and return its decoded JSON response."""

        _assert_read_only_method(method)
        payload = dict(params or {})
        if method == "batch":
            for command in payload.get("cmd", {}).values():
                _assert_read_only_method(command.split("?", 1)[0])

        url = f"{self.webhook_url}{method}.json"
        for attempt in range(self.max_retries + 1):
            self._wait_for_rate_limit()
            self._last_request_at = time.monotonic()
            response = self.session.post(url, json=payload, timeout=self.timeout)
            if response.status_code == 503 and attempt < self.max_retries:
                time.sleep(self.retry_backoff * (attempt + 1))
                continue
            response.raise_for_status()
            data = response.json()
            error = data.get("error")
            if error == "QUERY_LIMIT_EXCEEDED" and attempt < self.max_retries:
                time.sleep(self.retry_backoff * (attempt + 1))
                continue
            if error:
                description = data.get("error_description", "")
                raise BitrixAPIError(f"Bitrix method {method!r} failed: {error} {description}".strip())
            return data
        raise BitrixAPIError(f"Bitrix method {method!r} failed after {self.max_retries + 1} attempts")

    def batch(self, commands: Mapping[str, tuple[str, Mapping[str, Any]]]) -> dict[str, Any]:
        """Execute up to 50 read-only commands through Bitrix batch."""

        if len(commands) > BITRIX_BATCH_LIMIT:
            raise ValueError(f"Bitrix batch supports at most {BITRIX_BATCH_LIMIT} commands")
        cmd = {key: _build_batch_command(method, params) for key, (method, params) in commands.items()}
        response = self.call("batch", {"halt": 1, "cmd": cmd})
        result = response.get("result", {})
        if result.get("result_error"):
            raise BitrixAPIError(f"Bitrix batch failed: {result['result_error']}")
        return result.get("result", {})


def create_bitrix_client(webhook_url: str) -> BitrixRestClient:
    """Create the default read-only Bitrix REST client."""

    if not webhook_url:
        raise ValueError("Bitrix webhook URL is empty; set secrets['BITRIX_WEBHOOK_URL']")
    return BitrixRestClient(
        webhook_url,
        requests_per_second=2.0,
        timeout=60,
        max_retries=5,
        retry_backoff=2.0,
    )


def get_deal_category_id(category_name: str, client: BitrixRestClient) -> int:
    """Return Bitrix deal category ID by its visible funnel name."""

    try:
        response = client.call("crm.category.list", {"entityTypeId": BITRIX_DEAL_ENTITY_TYPE_ID})
        category_result = response.get("result", [])
        categories = category_result.get("categories", []) if isinstance(category_result, Mapping) else category_result
    except BitrixAPIError:
        response = client.call(
            "crm.dealcategory.list",
            {"order": {"SORT": "ASC"}, "select": ["ID", "NAME", "SORT"]},
        )
        categories = response.get("result", [])

    for category in categories:
        category_title = str(category.get("name") or category.get("NAME") or "").strip().casefold()
        if category_title == category_name.strip().casefold():
            return int(category.get("id") or category.get("ID"))
    raise ValueError(f"Deal funnel {category_name!r} was not found")


def _list_dataframe(
    client: BitrixRestClient,
    method: str,
    result_key: str,
    base_params: Mapping[str, Any],
    batch_size: int,
) -> pd.DataFrame:
    first_response = client.call(method, {**base_params, "start": 0})
    first_result = first_response.get("result", [])
    rows = list(first_result.get(result_key, []) if isinstance(first_result, Mapping) else first_result)
    total = int(first_response.get("total", len(rows)))
    if total <= BITRIX_PAGE_SIZE:
        return pd.DataFrame(rows)

    max_batch_size = max(1, min(batch_size, BITRIX_BATCH_LIMIT))
    starts = list(range(BITRIX_PAGE_SIZE, total, BITRIX_PAGE_SIZE))
    total_batches = math.ceil(len(starts) / max_batch_size)
    for batch_index in range(total_batches):
        chunk_starts = starts[batch_index * max_batch_size : (batch_index + 1) * max_batch_size]
        commands = {
            f"items_{start}": (method, {**base_params, "start": start})
            for start in chunk_starts
        }
        for page in client.batch(commands).values():
            if isinstance(page, Mapping):
                rows.extend(page.get(result_key, []))
            else:
                rows.extend(page or [])

    return pd.DataFrame(rows)


def collect_deals_dataframe(
    client: BitrixRestClient,
    category_name: str,
    category_id: int | None,
    select: Sequence[str],
    extra_filter: Mapping[str, Any] | None,
    batch_size: int,
) -> pd.DataFrame:
    """Collect all deals from a Bitrix CRM funnel into one DataFrame."""

    resolved_category_id = category_id if category_id is not None else get_deal_category_id(category_name, client)
    deal_filter: dict[str, Any] = {"CATEGORY_ID": resolved_category_id}
    if extra_filter:
        deal_filter.update(extra_filter)
    base_params: dict[str, Any] = {
        "select": list(select),
        "filter": deal_filter,
        "order": {"ID": "ASC"},
    }
    return _list_dataframe(client, "crm.deal.list", "result", base_params, batch_size)


def collect_portal_360_deals_dataframe(
    client: BitrixRestClient | None = None,
    category_name: str = "Поступление 360",
    category_id: int | None = None,
    select: Sequence[str] = DEFAULT_DEAL_SELECT,
    extra_filter: Mapping[str, Any] | None = None,
    batch_size: int = BITRIX_BATCH_LIMIT,
) -> pd.DataFrame:
    """Compatibility helper for the admissions funnel deal export."""

    rest = client if client is not None else create_bitrix_client(BITRIX_WEBHOOK_URL)
    return collect_deals_dataframe(rest, category_name, category_id, select, extra_filter, batch_size)


def collect_crm_items_dataframe(
    client: BitrixRestClient,
    entity_type_id: int,
    select: Sequence[str],
    extra_filter: Mapping[str, Any],
    batch_size: int,
) -> pd.DataFrame:
    """Collect Bitrix dynamic CRM items by entity type ID."""

    base_params: dict[str, Any] = {
        "entityTypeId": entity_type_id,
        "select": list(select),
        "filter": dict(extra_filter),
        "order": {"id": "ASC"},
    }
    return _list_dataframe(client, "crm.item.list", "items", base_params, batch_size)


def collect_bitrix_item_sources(
    client: BitrixRestClient,
    sources: Sequence[BitrixItemSource],
    batch_size: int,
) -> dict[str, pd.DataFrame]:
    """Collect configured Bitrix dynamic CRM item sources."""

    tables: dict[str, pd.DataFrame] = {}
    for source in sources:
        tables[source.name] = collect_crm_items_dataframe(
            client,
            source.entity_type_id,
            source.select,
            source.extra_filter,
            batch_size,
        )
    return tables
