"""Write dashboard data to Google Sheets."""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import gspread
import numpy as np
import pandas as pd
from gspread.utils import ValueRenderOption
from oauth2client.service_account import ServiceAccountCredentials

try:
    from .col_names import *
except ImportError:  # pragma: no cover - direct script execution fallback
    from col_names import *

ROOT_DIR = Path(__file__).resolve().parents[1]


def _credentials_path() -> Path:
    return ROOT_DIR / "service_credentials.json"


def _read_previous_delta_values(dashboard: gspread.Worksheet, prev_file: Path) -> tuple[np.ndarray, np.ndarray]:
    try:
        df_prev = pd.read_csv(prev_file)
        return df_prev[col_leads].to_numpy(), df_prev[col_applications].to_numpy()
    except FileNotFoundError:
        prev_leads = np.array(dashboard.get("J2:J49", value_render_option=ValueRenderOption.unformatted))[:, 0]
        prev_applications = np.array(dashboard.get("O2:O49", value_render_option=ValueRenderOption.unformatted))[:, 0]
        return prev_leads, prev_applications


def _write_history_cells(dashboard: gspread.Worksheet, history_data: pd.DataFrame) -> None:
    updates: tuple[tuple[str, str], ...] = (
        ("B53", "2025"),
        ("B55", "2024"),
        ("B57", "2023"),
        ("AM52", str(history_data.loc[2026, "early_invitations_unique"])),
        ("K53", str(history_data.loc[2025, "leads"][0])),
        ("K55", str(history_data.loc[2024, "leads"][0])),
        ("K57", str(history_data.loc[2023, "leads"][0])),
        ("O53", str(history_data.loc[2025, "applications"])),
        ("O55", str(history_data.loc[2024, "applications"])),
        ("O57", str(history_data.loc[2023, "applications"])),
        ("S53", str(history_data.loc[2025, "contracts"])),
        ("S55", str(history_data.loc[2024, "contracts"])),
        ("S57", str(history_data.loc[2023, "contracts"])),
        ("O52", str(history_data.loc[2026, "applications_unique"])),
        ("O54", str(history_data.loc[2025, "applications_unique"])),
        ("O56", str(history_data.loc[2024, "applications_unique"])),
        ("O58", str(history_data.loc[2023, "applications_unique"])),
    )
    for cell, value in updates:
        dashboard.update_acell(cell, value)


def update_sheet(aggregated_data: pd.DataFrame, update_delta: bool = False, history_data: pd.DataFrame | None = None) -> None:
    prev_file = ROOT_DIR / "templates" / "prev_data.csv"
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    if sys.platform not in {"win32", "darwin"}:
        raise ValueError(f"Unsupported platform: {sys.platform}")

    creds = ServiceAccountCredentials.from_json_keyfile_name(str(_credentials_path()), scope)
    client = gspread.authorize(creds)
    sheet = client.open("Еженедельный отчет 2026_общий")
    dashboard = sheet.get_worksheet(0)

    print("Google dashboard opened")

    str_time = datetime.now().strftime("%H:%M")
    str_date = datetime.now().strftime("%d.%m")

    if update_delta:
        prev_leads, prev_applications = _read_previous_delta_values(dashboard, prev_file)
        aggregated_data[col_leads_delta] = aggregated_data[col_leads] - prev_leads
        aggregated_data[col_applications_delta] = aggregated_data[col_applications] - prev_applications
        aggregated_data[[col_leads, col_applications]].to_csv(prev_file, index=False)
        print("Lead and application deltas updated")
    else:
        aggregated_data[col_leads_delta] = np.array(dashboard.get("L2:L49", value_render_option=ValueRenderOption.unformatted))[:, 0]
        aggregated_data[col_applications_delta] = np.array(dashboard.get("P2:P49", value_render_option=ValueRenderOption.unformatted))[:, 0]
        print("Weekly delta values reused from the sheet")

    dashboard.update([aggregated_data.columns.values.tolist()] + aggregated_data.values.tolist())
    print("Dashboard data written")

    dashboard.update_acell("B50", f"{str_time}, {str_date}.2026")
    if history_data is not None:
        _write_history_cells(dashboard, history_data)
