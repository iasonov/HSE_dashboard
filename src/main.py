'''Executable entrypoint for dashboard generation.'''

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import pandas as pd

from general_pipeline import calculate_dashboard
from update_pipeline import update_sheet


def run_dashboard(count_delta:bool = True, update_dashboard:bool = True, legacy:bool = False, debug:bool = False) -> Path:
    '''Run the dashboard pipeline and return the exported workbook path.'''
    # legacy: True - ASAV & AIS PK, False - only Bitrix

    output_dir = Path('data') / 'dashboards'
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime('%Y.%m.%d-%H.%M.%S')
    output_path = output_dir / f'dashboard{timestamp}.xlsx'

    current_data, history_data = calculate_dashboard(debug, legacy)
    current_data.to_excel(output_path)
    if update_dashboard:
        update_sheet(pd.read_excel(output_path), count_delta, history_data)

    return output_path


if __name__ == '__main__':

    bitrix_mode = True
    if not bitrix_mode: # ASAV & AIS PK xlsx
        count_delta = False
        update_dashboard = False
        legacy = True
        debug = False
    else: #bitrix
        count_delta = True
        update_dashboard = True
        legacy = False
        debug = False

    run_dashboard(count_delta=count_delta, update_dashboard=update_dashboard, legacy=legacy, debug=debug)
