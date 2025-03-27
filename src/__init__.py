# This is the __init__.py file for the my_python_project package
from process import process_excel_files
from update import update_sheet
import datetime
import pandas as pd

str = datetime.date.today().strftime("%Y.%m.%d-%H.%M.%S")
file_name = "dashboard" + str + ".xlsx"

df = process_excel_files()
df.to_excel(file_name)

update_sheet(pd.read_excel(file_name))
# update_sheet(process_excel_files())