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

sheet_instance.update('A1', [[1, 2], [3, 4]])
# # get all the records of the data
# records_data = sheet_instance.get_all_records()

# # view the data
# print(records_data)