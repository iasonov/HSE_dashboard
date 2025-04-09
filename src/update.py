from col_names import *
import numpy as np

def update_sheet(aggregated_data, update_delta=False):
    import gspread
    import pandas as pd
    from oauth2client.service_account import ServiceAccountCredentials

    # define the scope
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    # add credentials to the account
    creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/iasonov/Programming/Python/HSE_dashboard/service_credentials.json', scope)

    # authorize the clientsheet
    client = gspread.authorize(creds)

    # get the instance of the Spreadsheet
    sheet = client.open('Еженедельный отчет 2025_общий')

    # get the first sheet of the Spreadsheet
    sheet_instance = sheet.get_worksheet(0)

    if update_delta:
        from gspread.utils import ValueRenderOption
        # sheet_prev = sheet.worksheet('prev')
        # sheet.del_worksheet(sheet_prev)
        # TODO - replace using formulae (not absolute range)
        prev_leads        = np.array(sheet_instance.get("J2:J42", value_render_option=ValueRenderOption.unformatted))
        prev_applications = np.array(sheet_instance.get("N2:N42", value_render_option=ValueRenderOption.unformatted))
        aggregated_data[col_leads_delta]        = aggregated_data[col_leads]        - prev_leads[:,0]
        aggregated_data[col_applications_delta] = aggregated_data[col_applications] - prev_applications[:,0]

        # sheet_instance.duplicate(1, sheet_instance.id+1, "prev")

    # Clear the worksheet before updating
    # worksheet.clear()

    # Update the worksheet with the aggregated data
    # try:
    # sheet_instance.set_dataframe(aggregated_data, 'A1')
    sheet_instance.update([aggregated_data.columns.values.tolist()] + aggregated_data.values.tolist())
    #     #sheet_instance.update([aggregated_data.columns.values.tolist()] + aggregated_data.values.tolist())
    #     print("Google Spreadsheet updated successfully.")
    # except:
    #     aggregated_data.to_excel("dashboard.xlsx")
    #     print("Error update")

    return