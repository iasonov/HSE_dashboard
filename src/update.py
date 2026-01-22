import sys
import numpy as np
import gspread
from gspread.utils import ValueRenderOption
import pandas as pd
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from col_names import *

def update_sheet(aggregated_data, update_delta=False, history_data=None):

    prev_file = "templates/prev_data.csv"
    # define the scope
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    # add credentials to the account
    if sys.platform == 'win32':
        path_credentials = 'C:\\Users\\tijuanap\\Programs\\HSE_dashboard\\service_credentials.json'
    elif sys.platform == 'darwin':
        path_credentials = '/Users/iasonov/Programming/Python/HSE_dashboard/service_credentials.json'
    else:
        raise ValueError(f"Unsupported platform: {sys.platform}")
        
    creds = ServiceAccountCredentials.from_json_keyfile_name(path_credentials, scope)

    # authorize the clientsheet
    client = gspread.authorize(creds)

    # get the instance of the Spreadsheet
    sheet = client.open('Еженедельный отчет 2025_общий')

    # get the first sheet of the Spreadsheet
    dashboard = sheet.get_worksheet(0)

    print("Гугл-дашборд открыт")

    str_time = datetime.now().strftime("%H:%M")
    str_date = datetime.now().strftime("%d.%m")

    if update_delta:
        # TODO - replace using formulae (not absolute range)
        try:
            df_prev = pd.read_csv(prev_file)
            prev_leads = df_prev[col_leads]
            prev_applications = df_prev[col_applications]
            print("Данные о лидах с прошлого обновления считаны")
        except:
            prev_leads        = np.array(dashboard.get("J2:J42", value_render_option=ValueRenderOption.unformatted))[:,0]
            prev_applications = np.array(dashboard.get("O2:O42", value_render_option=ValueRenderOption.unformatted))[:,0]
            print("Данные в " + prev_file + " на локальном диске не найдены, считаю дельту относительно гугл-дашборда")
        aggregated_data[col_leads_delta]        = aggregated_data[col_leads]        - prev_leads
        aggregated_data[col_applications_delta] = aggregated_data[col_applications] - prev_applications
        aggregated_data[[col_leads, col_applications]].to_csv(prev_file)
        print("Сведения о лидах и регистрациях в ЛК за полнедели обновлены")
    else:
        prev_leads_delta        = np.array(dashboard.get("L2:L42", value_render_option=ValueRenderOption.unformatted))
        prev_applications_delta = np.array(dashboard.get("P2:P42", value_render_option=ValueRenderOption.unformatted))
        aggregated_data[col_leads_delta]        = prev_leads_delta[:,0]
        aggregated_data[col_applications_delta] = prev_applications_delta[:,0]
        print("Данные за полнедели не обновляются")

    dashboard.update([aggregated_data.columns.values.tolist()] + aggregated_data.values.tolist())
    print("Данные в гугл-дашборд записаны")

    # TODO replace without absolute cell indexes
    dashboard.update_acell('B43', str_time + ", " + str_date + ".2025") # 2025
    if history_data is not None:
        dashboard.update_acell('B46', str_date + ".2024") # 2024
        dashboard.update_acell('B48', str_date + ".2023") # 2024
        dashboard.update_acell('K46', str(history_data.loc[2024, 'leads'][0]))
        dashboard.update_acell('K48', str(history_data.loc[2023, 'leads'][0]))
        dashboard.update_acell('O46', str(history_data.loc[2024, 'applications']))
        dashboard.update_acell('O48', str(history_data.loc[2023, 'applications']))
        dashboard.update_acell('S46', str(history_data.loc[2024, 'contracts']))
        dashboard.update_acell('S48', str(history_data.loc[2023, 'contracts']))
        dashboard.update_acell('O45', str(history_data.loc[2025, 'applications_unique']))
        dashboard.update_acell('O47', str(history_data.loc[2024, 'applications_unique']))
        dashboard.update_acell('O49', str(history_data.loc[2023, 'applications_unique']))

    return