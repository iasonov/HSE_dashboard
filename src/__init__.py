from process import process_current_files
from update import update_sheet
from datetime import datetime
import pandas as pd

str = datetime.now().strftime("%Y.%m.%d-%H.%M.%S")
file_name = "data/dashboards/dashboard" + str + ".xlsx"

df_current, df_history = process_current_files()
df_current.to_excel(file_name)

# df_history = process_history_files()

update_deltas = False # each monday and thursday
update_google_dashboard = True

if update_google_dashboard:
  update_sheet(pd.read_excel(file_name), update_deltas, df_history)