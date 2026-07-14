'''Write dashboard data to Google Sheets.'''

from __future__ import annotations

import sys
import time
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
    return ROOT_DIR / 'service_credentials.json'


def _read_previous_delta_values(dashboard: gspread.Worksheet, prev_file: Path) -> tuple[np.ndarray, np.ndarray]:
    try:
        df_prev = pd.read_csv(prev_file)
        return df_prev[col_leads].to_numpy(), df_prev[col_applications].to_numpy()
    except FileNotFoundError:
        prev_leads = np.array(dashboard.get('J2:J49', value_render_option=ValueRenderOption.unformatted))[:, 0]
        prev_applications = np.array(dashboard.get('O2:O49', value_render_option=ValueRenderOption.unformatted))[:, 0]
        return prev_leads, prev_applications


def _write_history_cells(dashboard: gspread.Worksheet, history_data: pd.DataFrame) -> None:
    updates: tuple[tuple[str, str], ...] = (
        ('B53', '2025'),
        ('B55', '2024'),
        ('B57', '2023'),
        ('AM52', str(history_data.loc[2026, 'early_invitations_unique'])),
        ('K53', str(history_data.loc[2025, 'leads'][0])),
        ('K55', str(history_data.loc[2024, 'leads'][0])),
        ('K57', str(history_data.loc[2023, 'leads'][0])),
        ('O53', str(history_data.loc[2025, 'applications'])),
        ('O55', str(history_data.loc[2024, 'applications'])),
        ('O57', str(history_data.loc[2023, 'applications'])),
        ('S53', str(history_data.loc[2025, 'contracts'])),
        ('S55', str(history_data.loc[2024, 'contracts'])),
        ('S57', str(history_data.loc[2023, 'contracts'])),
        ('O52', str(history_data.loc[2026, 'applications_unique'])),
        ('P52', str(history_data.loc[2026, 'applications_no_rossokhins_unique'])),
        ('Q52', str(history_data.loc[2026, 'applications_bachelors_unique'])),
        ('R52', str(history_data.loc[2026, 'applications_masters_unique'])),
        ('S52', str(history_data.loc[2026, 'applications_masters_no_rossokhins_unique'])),
        ('O54', str(history_data.loc[2025, 'applications_unique'])),
        ('O56', str(history_data.loc[2024, 'applications_unique'])),
        ('O58', str(history_data.loc[2023, 'applications_unique'])),
    )
    for cell, value in updates:
        dashboard.update_acell(cell, value)


def update_sheet(aggregated_data: pd.DataFrame, update_delta: bool = False, history_data: pd.DataFrame | None = None) -> None:
    print('Starting Google Sheets update')
    prev_file = ROOT_DIR / 'templates' / 'prev_data.csv'
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    if sys.platform not in {'win32', 'darwin'}:
        raise ValueError(f'Unsupported platform: {sys.platform}')

    creds = ServiceAccountCredentials.from_json_keyfile_name(str(_credentials_path()), scope)
    client = gspread.authorize(creds)
    sheet = client.open('Еженедельный отчет 2026_общий')
    dashboard = sheet.get_worksheet(0)

    print('Google dashboard opened')

    str_time = datetime.now().strftime('%H:%M')
    str_date = datetime.now().strftime('%d.%m')

    # TODO make if for legacy mode
    # if update_delta:
    #     prev_leads, prev_applications = _read_previous_delta_values(dashboard, prev_file)
    #     aggregated_data[col_leads_delta] = aggregated_data[col_leads] - prev_leads
    #     aggregated_data[col_applications_delta] = aggregated_data[col_applications] - prev_applications
    #     aggregated_data[[col_leads, col_applications]].to_csv(prev_file, index=False)
    #     print('Lead and application deltas updated')
    # else:
    #     aggregated_data[col_leads_delta] = np.array(dashboard.get('L2:L49', value_render_option=ValueRenderOption.unformatted))[:, 0]
    #     aggregated_data[col_applications_delta] = np.array(dashboard.get('P2:P49', value_render_option=ValueRenderOption.unformatted))[:, 0]
    #     print('Weekly delta values reused from the sheet')

    aggregated_data = aggregated_data.fillna("")
    dashboard.update([aggregated_data.columns.values.tolist()] + aggregated_data.values.tolist())
    print('Dashboard data written')

    dashboard.update_acell('B50', f'{str_time}, {str_date}.2026')
    if history_data is not None:
        _write_history_cells(dashboard, history_data)


def update_exams_dates_sheet(exams_data: dict[str, pd.DataFrame], online_exams_dict: dict[str, str], delay_sec: float = 1.0) -> None:
    '''Выгружает таблицу с датами ВИ в Google Sheets "Даты ВИ 2026".

    Для каждого предмета (ключа словаря) проверяет наличие вкладки:
    если вкладка есть — обновляет данные, если нет — создаёт и записывает.

    Args:
        exams_data: словарь {название_предмета: DataFrame} от process_exams_dates_file
        online_exams_dict: словарь {название_предмета: название_онлайн_программы}
    '''
    print('Начинаем выгрузку дат ВИ в Google Sheets')
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    if sys.platform not in {'win32', 'darwin'}:
        raise ValueError(f'Unsupported platform: {sys.platform}')

    creds = ServiceAccountCredentials.from_json_keyfile_name(str(_credentials_path()), scope)
    client = gspread.authorize(creds)

    try:
        sheet = client.open('Даты ВИ 2026')
        print('Google таблица "Даты ВИ 2026" открыта')
    except gspread.SpreadsheetNotFound:
        sheet = client.create('Даты ВИ 2026')
        print('Google таблица "Даты ВИ 2026" создана')
        # Удаляем дефолтную вкладку
        time.sleep(delay_sec)
        default_ws = sheet.get_worksheet(0)
        if default_ws.title == 'Sheet1':
            time.sleep(delay_sec)
            sheet.del_worksheet(default_ws)

    time.sleep(delay_sec)
    existing_titles = {ws.title for ws in sheet.worksheets()}

    for subject, df in exams_data.items():
        # Название вкладки ограничено 100 символами
        title = (online_exams_dict.get(subject, "NOT FOUND") + ': ' + subject.split(':')[0])[:100]
        values = [df.columns.tolist()] + df.fillna('').values.tolist()

        if title in existing_titles:
            worksheet = sheet.worksheet(title)
            time.sleep(delay_sec)
            worksheet.clear()
            n_rows = len(values)
            n_cols = len(values[0]) if values else 0
            if n_rows > 0 and n_cols > 0:
                time.sleep(delay_sec)
                worksheet.resize(n_rows, n_cols)
                
            print(f'Обновлена вкладка: {title}')
        else:
            n_rows = max(len(values), 1)
            n_cols = max(len(values[0]), 1) if values else 1
            time.sleep(delay_sec)
            worksheet = sheet.add_worksheet(title=title, rows=n_rows, cols=n_cols)
            print(f'Создана вкладка: {title}')
        time.sleep(delay_sec)
        worksheet.update(values)
        print(f'Данные на вкладку "{title}" загружены')

    print('Выгрузка дат ВИ завершена')
