def update_sheet(aggregated_data):
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

    # Clear the worksheet before updating
    # worksheet.clear()

    # Update the worksheet with the aggregated data
    sheet_instance.update([aggregated_data.columns.values.tolist()] + aggregated_data.values.tolist())

    print("Google Spreadsheet updated successfully.")
    return