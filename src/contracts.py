"""Static registry for dashboard sources and metric definitions."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class SourceSpec:
    """Describe one data source used by the dashboard."""

    name: str
    folder: str
    primary_filename: str
    file_patterns: tuple[str, ...]
    extensions: tuple[str, ...]
    description: str


@dataclass(frozen=True, slots=True)
class MetricSpec:
    """Describe one dashboard metric."""

    name: str
    source: str
    grain: str
    formula: str
    dependencies: tuple[str, ...]


DATA_LAYOUT: tuple[tuple[str, str], ...] = (
    ("data/raw", "Source exports"),
    ("data/processed", "Normalized intermediate tables"),
    ("data/archive", "Historical snapshots"),
    ("data/dashboards", "Generated dashboard workbooks"),
    ("templates", "Templates and mappings"),
)


SOURCE_SPECS: tuple[SourceSpec, ...] = (
    SourceSpec(
        name="bitrix",
        folder="data/raw",
        primary_filename="DEAL_*.xls*",
        file_patterns=("*DEAL*.xls*",),
        extensions=("xls", "xlsx", "csv"),
        description="Leads export from Bitrix or the future API adapter.",
    ),
    SourceSpec(
        name="portal",
        folder="data/raw",
        primary_filename="portal.xls",
        file_patterns=("*port*.xls*",),
        extensions=("xls", "xlsx"),
        description="Partner portal lead export.",
    ),
    SourceSpec(
        name="asav_master",
        folder="data/raw",
        primary_filename="asav.xlsx",
        file_patterns=("*asav*.xls*",),
        extensions=("xls", "xlsx"),
        description="Master's ASAV export.",
    ),
    SourceSpec(
        name="asav_foreign",
        folder="data/raw",
        primary_filename="asav_foreign.xlsx",
        file_patterns=("*foreign*.xls*",),
        extensions=("xls", "xlsx"),
        description="Foreign-track ASAV export.",
    ),
    SourceSpec(
        name="aispk_applications",
        folder="data/raw",
        primary_filename="bac_applications.xlsx",
        file_patterns=("*app*.xls*",),
        extensions=("xls", "xlsx"),
        description="Bachelor applications export from AIS PK.",
    ),
    SourceSpec(
        name="aispk_contracts",
        folder="data/raw",
        primary_filename="bac_contracts.xlsx",
        file_patterns=("*con*.xls*",),
        extensions=("xls", "xlsx"),
        description="Bachelor contracts export from AIS PK.",
    ),
    SourceSpec(
        name="aispk_enrollments",
        folder="data/raw",
        primary_filename="bac_enrolled.xlsx",
        file_patterns=("*enroll*.xls*",),
        extensions=("xls", "xlsx"),
        description="Bachelor enrollments export from AIS PK.",
    ),
)


METRIC_SPECS: tuple[MetricSpec, ...] = (
    MetricSpec(
        name="leads",
        source="bitrix_and_portal",
        grain="program",
        formula="count rows by program",
        dependencies=("lead_date", "program"),
    ),
    MetricSpec(
        name="applications",
        source="asav_and_aispk",
        grain="program",
        formula="count applications by program across master's ASAV and bachelor's AIS PK",
        dependencies=("applications_dates", "program"),
    ),
    MetricSpec(
        name="contracts",
        source="asav_and_aispk",
        grain="program",
        formula="count contracts by program across master's ASAV and bachelor's AIS PK",
        dependencies=("contracts_dates", "program"),
    ),
    MetricSpec(
        name="payments",
        source="asav_and_aispk",
        grain="program",
        formula="count paid contracts by program across master's ASAV and bachelor's AIS PK",
        dependencies=("payment_status", "program"),
    ),
    MetricSpec(
        name="enrollments",
        source="asav_and_aispk",
        grain="program",
        formula="count enrollments by program across master's ASAV and bachelor's AIS PK",
        dependencies=("enrollment_status", "program"),
    ),
)

