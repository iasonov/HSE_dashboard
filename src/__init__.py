# This is the __init__.py file for the my_python_project package
from process import process_history_files, process_current_files
from update import update_sheet
from datetime import datetime
import pandas as pd

str = datetime.now().strftime("%Y.%m.%d-%H.%M.%S")
file_name = "dashboard" + str + ".xlsx"

df_current = process_current_files()
df_current.to_excel(file_name)


df_history = process_history_files()

# update_sheet(pd.read_excel(file_name), True, df_history)


# update_sheet(process_excel_files())