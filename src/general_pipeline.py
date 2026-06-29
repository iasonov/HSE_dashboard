import pandas as pd
import numpy as np
from datetime import datetime

def categorize_ages(age_column):
    # Определяем диапазоны
    bins = [0, 17, 23, 29, 35, 41, 47, float('inf')]
    labels = ['0-17', '18-23', '24-29', '30-35', '36-41', '42-47', '48+']

    # Используем pd.cut для разбиения на интервалы
    categories = pd.cut(age_column, bins=bins, labels=labels, right=True, include_lowest=True)

    # Считаем количество в каждом диапазоне
    counts = categories.value_counts().sort_index()

    return np.array2string(counts.values, separator=';')[1:-1]

def years_ago(years, from_date=None):
    if from_date is None:
        from_date = datetime.now()
    try:
        return from_date.replace(year=from_date.year - years)
    except ValueError:
        # Must be 2/29!
        assert from_date.month == 2 and from_date.day == 29 # can be removed
        return from_date.replace(month=2, day=28,
                                 year=from_date.year-years)

def num_years(begin, end=None):
    if end is None:
        end = datetime.now()
    if pd.isna(begin):
        begin = datetime.now()
    years = end.year - begin.year
    if begin > years_ago(years, end):
        return years - 1
    return years

def insert_values(df_dashboard, df_values, col_join, col_values): # df_values should have 'values' column
    for i, row in df_dashboard.iterrows():
        if row[col_join] in df_values[col_join].values:
            df_dashboard.loc[i, col_values] = df_values[df_values[col_join] == row[col_join]].values[0][1] # row[col_values]
            #row[col_values] = df_values[(row[col_join], 'values')]
        else:
            df_dashboard.loc[i, col_values] = 0
        # print(row)
    # df_dashboard.loc[df_dashboard[col_join].isin(df_values[col_join]), col_values] = df_values.loc[df_values[col_join].isin(df_dashboard[col_join]), 'values'].values
    # df_dashboard[col_values] = df_dashboard[col_values].fillna(0).astype(int)
    return df_dashboard[col_values]


def process_by_week(df, col_program, col_date, start_date=pd.Timestamp(year=2025, month=9, day=29, hour=0, minute=0, second=0), col_values='count', format='%d.%m.%Y %H:%M:%S'):
    df_temp = df.copy().dropna(subset=[col_date])
    df_temp[col_date] = pd.to_datetime(df_temp[col_date], format=format)

    # Вычисляем номер недели (можно также использовать понедельник недели как якорь)
    df_temp['week_start'] = df_temp[col_date].dt.to_period('W-SUN').apply(lambda r: r.start_time) # немного магии - тут надо начинать с пн; df_temp['week_start'] = df_temp[col_date].dt.to_period('W-SUN').dt.start_time

    # Группируем по программе и неделе
    weekly_counts = df_temp.groupby([col_program, 'week_start']).size().reset_index(name=col_values)

    # Получим все уникальные программы и все недели
    all_programs = weekly_counts[col_program].unique()
    if all_programs.size == 0:
        weekly_counts.loc[0, 'week_start'] = datetime.now()

    all_weeks = pd.date_range(start=start_date,
                            end=weekly_counts['week_start'].max(),
                            freq='W-MON')  # каждую неделю по вторникам TODO: check различия MON & TUE

    # Создаем полную сетку: программа × неделя
    full_index = pd.MultiIndex.from_product([all_programs, all_weeks], names=[col_program, 'week_start'])
    full_df = pd.DataFrame(index=full_index).reset_index()

    # Объединяем с посчитанными заявками
    merged = pd.merge(full_df, weekly_counts, how='left', on=[col_program, 'week_start'])
    merged[col_values] = merged[col_values].fillna(0).astype(int)

    # Группируем по программе и объединяем значения в строку через ';'
    result = merged.groupby(col_program)[col_values].apply(lambda x: ';'.join(map(str, x))).reset_index(name=col_values)
    return result.set_index(col_program, verify_integrity=True, drop=True)[col_values]



def calculate_dashboard(debug:bool = None, legacy:bool = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    if legacy:
        from files_pipeline import process_from_current_files
        return process_from_current_files(debug)
    else:
        from bitrix_pipeline import process_from_bitrix
        return process_from_bitrix(debug)
