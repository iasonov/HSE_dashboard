import numpy as np
import gspread
from gspread.utils import ValueRenderOption
import pandas as pd
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
from col_names import *

def update_sheet(aggregated_data, update_delta=False, history_data=None):


    # define the scope
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    # add credentials to the account
    creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/iasonov/Programming/Python/HSE_dashboard/service_credentials.json', scope)

    # authorize the clientsheet
    client = gspread.authorize(creds)

    # get the instance of the Spreadsheet
    sheet = client.open('Еженедельный отчет 2025_общий')

    # get the first sheet of the Spreadsheet
    dashboard_sales = sheet.get_worksheet(0)
    # dashboard_simple = sheet.get_worksheet(1)


    if update_delta:
        # sheet_prev = sheet.worksheet('prev')
        # sheet.del_worksheet(sheet_prev)
        # TODO - replace using formulae (not absolute range)
        prev_leads        = np.array(dashboard_sales.get("J2:J42", value_render_option=ValueRenderOption.unformatted))
        prev_applications = np.array(dashboard_sales.get("O2:O42", value_render_option=ValueRenderOption.unformatted))
        aggregated_data[col_leads_delta]        = aggregated_data[col_leads]        - prev_leads[:,0]
        aggregated_data[col_applications_delta] = aggregated_data[col_applications] - prev_applications[:,0]
    else:
        prev_leads_delta        = np.array(dashboard_sales.get("L2:L42", value_render_option=ValueRenderOption.unformatted))
        prev_applications_delta = np.array(dashboard_sales.get("P2:P42", value_render_option=ValueRenderOption.unformatted))
        aggregated_data[col_leads_delta]        = prev_leads_delta[:,0]
        aggregated_data[col_applications_delta] = prev_applications_delta[:,0]
#БАКЭКАН. Экономический анализ / Москва / 380301 Экономика / факультет экономических наук / Бакалавриат

        # sheet_instance.duplicate(1, sheet_instance.id+1, "prev")

    # Clear the worksheet before updating
    # worksheet.clear()

    # Update the worksheet with the aggregated data
    # try:
    # sheet_instance.set_dataframe(aggregated_data, 'A1')
    dashboard_sales.update([aggregated_data.columns.values.tolist()] + aggregated_data.values.tolist())
    #     #sheet_instance.update([aggregated_data.columns.values.tolist()] + aggregated_data.values.tolist())
    #     print("Google Spreadsheet updated successfully.")
    # except:
    #     aggregated_data.to_excel("dashboard.xlsx")
    #     print("Error update")
    str_time = datetime.now().strftime("%H:%M")
    str_date = datetime.now().strftime("%d.%m")

    # TODO replace without absolute cell indexes
    dashboard_sales.update_acell('B43', str_time + ", " + str_date + ".2025") # 2025
    if history_data is not None:
        dashboard_sales.update_acell('B45', str_date + ".2024") # 2024
        dashboard_sales.update_acell('B46', str_date + ".2023") # 2024
        dashboard_sales.update_acell('O45', str(history_data.loc[2024, 'applications'][0]))
        dashboard_sales.update_acell('O46', str(history_data.loc[2023, 'applications'][0]))
        dashboard_sales.update_acell('S45', str(history_data.loc[2024, 'contracts'][0]))
        dashboard_sales.update_acell('S46', str(history_data.loc[2023, 'contracts'][0]))

    return