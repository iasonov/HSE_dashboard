import glob
import pandas as pd
import numpy as np
from col_names import *
from datetime import datetime, timedelta

def categorize_ages(age_column):
    # Определяем диапазоны
    bins = [0, 17, 23, 29, 35, 41, 47, float('inf')]
    labels = ['0-17', '18-23', '24-29', '30-35', '36-41', '42-47', '48+']

    # Используем pd.cut для разбиения на интервалы
    categories = pd.cut(age_column, bins=bins, labels=labels, right=True, include_lowest=True)

    # Считаем количество в каждом диапазоне
    counts = categories.value_counts().sort_index()

    return np.array2string(counts.values, separator=";")[1:-1]

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
    num_years = int((end - begin).days / 365.2425)
    if begin > years_ago(num_years, end):
        return num_years - 1
    else:
        return num_years

def insert_values(df_dashboard, df_values, col_join, col_values): # df_values should have "values" column
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

def process_history_files():
    templates_folder = "templates/"
    master_leads_file_2023        = "bitrix_2023-04-01_2023-09-15.csv"
    master_leads_file_2024        = "bitrix_2024-04-01_2024-09-15.csv"
    bitrix_file_2024              = "bitrix_2024-04-01_2024-09-15.xlsx"
    # master_applications_file_2023 = "asav_2023_applications.csv"
    # master_contracts_file_2023    = "asav_2023_contracts.csv"
    # master_applications_file_2024 = "asav_2024_applications.csv"
    # master_contracts_file_2024    = "asav_2024_contracts.csv"
    asav_file_2023                = "asav_2023.xlsx"
    asav_file_2024                = "asav_2024.xlsx"
    bachelor_file_2024            = "bachelor_2024.xls"


    # TODO вписать сводные расчеты

    print("Начинаем считывать исторические данные")

    try:
        # applications_dates_2023 = pd.read_csv(templates_folder + master_applications_file_2023, parse_dates=[0], date_format="%d.%m.%Y")
        # applications_dates_2024 = pd.read_csv(templates_folder + master_applications_file_2024, parse_dates=[0], date_format="%d.%m.%Y")
        # contracts_dates_2023 = pd.read_csv(templates_folder + master_contracts_file_2023, parse_dates=[0], date_format="%d.%m.%Y")
        # contracts_dates_2024 = pd.read_csv(templates_folder + master_contracts_file_2024, parse_dates=[0], date_format="%d.%m.%Y")
        leads_dates_2023 = pd.read_csv(templates_folder + master_leads_file_2023, parse_dates=[0], date_format="%d.%m.%Y")
        leads_dates_2024 = pd.read_csv(templates_folder + master_leads_file_2024, parse_dates=[0], date_format="%d.%m.%Y")
        leads_dates_2024_by_program = pd.read_excel(templates_folder + bitrix_file_2024, usecols="J:N") #, parse_dates=[0], date_format="%d.%m.%Y  %hh:%mm:%ss")
        leads_dates_2024_by_program['leads_dates'] = pd.to_datetime(leads_dates_2024_by_program['leads_dates'], errors='coerce', format="%d.%m.%Y  %hh:%mm:%ss")
        leads_dates_2024_by_program[col_programs_names].fillna(main_studyonline, inplace=True)

        bachelor_2024 = pd.read_excel(templates_folder + bachelor_file_2024) #, usecols='A:H,J:AB')
        bachelor_2024['applications_dates'] = pd.to_datetime(bachelor_2024['applications_dates'], errors='coerce', format='%d.%m.%Y')
        bachelor_2024['contracts_dates']    = pd.to_datetime(bachelor_2024['contracts_dates'],    errors='coerce', format='%d.%m.%Y')

        asav_2023 = pd.read_excel(templates_folder + asav_file_2023, parse_dates=[0, 1], skiprows=1, date_format="%d.%m.%Y")
        asav_2023['applications_dates'] = pd.to_datetime(asav_2023['applications_dates'], format='%Y-%m-%d 00:00:00')
        asav_2023['contracts_dates'] = pd.to_datetime(asav_2023['contracts_dates'], errors='coerce', format='%d.%m.%Y')

        asav_2024 = pd.read_excel(templates_folder + asav_file_2024, parse_dates=[0, 1], skiprows=1, date_format="%d.%m.%Y")
        asav_2024['applications_dates'] = pd.to_datetime(asav_2024['applications_dates'], format='%Y-%m-%d 00:00:00')
        asav_2024['contracts_dates'] = pd.to_datetime(asav_2024['contracts_dates'], errors='coerce', format='%d.%m.%Y')

    except:
        print("Files of previous years are not founded or have errors")
        # , master_applications_file_2023)
        # print(master_contracts_file_2023)
        # print(master_applications_file_2024)
        # print(master_contracts_file_2024)
        print(master_leads_file_2023)
        print(master_leads_file_2024)
        print(bitrix_file_2024)
        print(asav_file_2023)
        print(asav_file_2024)
        print(bachelor_2024)


    now = datetime.now()
    delta_now_2023 = timedelta(days=365+366)
    delta_now_2024 = timedelta(days=365)

    asav_2023_no_duplicates = asav_2023.drop_duplicates(subset=[col_id_asav])
    asav_2024_no_duplicates = asav_2024.drop_duplicates(subset=[col_id_asav])
    bachelor_2024_no_duplicates = bachelor_2024.drop_duplicates(subset=[col_id_bachelor])

    df_pivot = pd.DataFrame.from_dict({'leads' :
                                {2023: leads_dates_2023.where(leads_dates_2023 + delta_now_2023 <= now).count(),
                                 2024: leads_dates_2024.where(leads_dates_2024 + delta_now_2024 <= now).count()},
                                # 'applications_old' :
                                # {2023: applications_dates_2023.where(applications_dates_2023 + delta_now_2023 <= now).count(),
                                #  2024: applications_dates_2024.where(applications_dates_2024 + delta_now_2024 <= now).count()},
                                # 'contracts_old' :
                                # {2023: contracts_dates_2023.where(contracts_dates_2023 + delta_now_2023 <= now).count(),
                                #  2024: contracts_dates_2024.where(contracts_dates_2024 + delta_now_2024 <= now).count()},
                                'applications' :
                                {2023: asav_2023[asav_2023['applications_dates'] + delta_now_2023 <= now]['applications_dates'].count(),
                                 2024: asav_2024[asav_2024['applications_dates'] + delta_now_2024 <= now]['applications_dates'].count() + bachelor_2024[bachelor_2024['applications_dates'] + delta_now_2024 <= now]['applications_dates'].count()},
                                'contracts' :
                                {2023: asav_2023[asav_2023['contracts_dates'] + delta_now_2023 <= now]['contracts_dates'].count(),
                                 2024: asav_2024[asav_2024['contracts_dates'] + delta_now_2024 <= now]['contracts_dates'].count() + bachelor_2024[bachelor_2024['contracts_dates'] + delta_now_2024 <= now]['contracts_dates'].count()},
                                'applications_unique' :
                                {2023: asav_2023_no_duplicates[asav_2023_no_duplicates['applications_dates'] + delta_now_2023 <= now]['applications_dates'].count(),
                                 2024: asav_2024_no_duplicates[asav_2024_no_duplicates['applications_dates'] + delta_now_2024 <= now]['applications_dates'].count() + bachelor_2024_no_duplicates[bachelor_2024_no_duplicates['applications_dates'] + delta_now_2024 <= now]['applications_dates'].count()}
                                }) #columns=['year', 'applications', 'contracts']
    df_leads_prev        = leads_dates_2024_by_program[leads_dates_2024_by_program['leads_dates'] + delta_now_2024 <= now].groupby(col_programs_names)[col_programs_names].count()
    df_applications_prev = pd.concat([asav_2024[asav_2024['applications_dates'] + delta_now_2024 <= now].groupby(master_col_programs)[master_col_programs].count(),
                                     bachelor_2024[bachelor_2024['applications_dates'] + delta_now_2024 <= now].groupby(bachelor_col_programs)[bachelor_col_programs].count()])
    df_contracts_prev    = pd.concat([asav_2024[asav_2024['contracts_dates'] + delta_now_2024 <= now].groupby(master_col_programs)[master_col_programs].count(),
                                     bachelor_2024[bachelor_2024['contracts_dates'] + delta_now_2024 <= now].groupby(bachelor_col_programs)[bachelor_col_programs].count()])


    print("Исторические данные считаны")
    return df_pivot, df_leads_prev, df_applications_prev, df_contracts_prev

def process_foreign_programs(df, programs_names):
    df[master_foreign_col_programs_2] = df[master_foreign_col_programs_2].fillna("")
    is_online = df[master_foreign_col_programs_1].isin(programs_names)
    for i, row in df.iterrows():
        if not is_online.loc[i] or row[master_foreign_col_faculty_1] == "Факультет Санкт-Петербургская школа экономики и менеджмента" or row[master_foreign_col_faculty_1] == "Факультет экономики":
            df.loc[i, master_foreign_col_programs_1] = df.loc[i, master_foreign_col_programs_2]

    # df[master_foreign_col_programs_1] = df[master_foreign_col_programs_1] if df[master_foreign_col_programs_1].isin(programs_names) and df[master_foreign_col_faculty_1] != "Факультет Санкт-Петербургская школа экономики и менеджмента" else df[master_foreign_col_programs_2]
    # df[master_foreign_col_programs_1].fillna(df[master_foreign_col_programs_2])
    df = df[df[master_foreign_col_programs_1].isin(programs_names)]
    # df[master_foreign_col_faculty_1] = df[master_foreign_col_faculty_1].fillna("")
    # df[master_foreign_col_faculty_2] = df[master_foreign_col_faculty_2].fillna("")
    # # df[master_foreign_col_program_final] = df[master_foreign_col_program_final].fillna("")
    # # df[master_foreign_col_faculty_final] = df[master_foreign_col_faculty_final].fillna("")
    # df[col_program] = df[master_foreign_col_programs_1] + df[master_foreign_col_programs_2]
    return df

def process_by_week(df, col_program, col_date, col_values='count', format='%d.%m.%Y %H:%M:%S'):
    df_temp = df.copy().dropna(subset=[col_date])
    df_temp[col_date] = pd.to_datetime(df_temp[col_date], format=format)

    # Вычисляем номер недели (можно также использовать понедельник недели как якорь)
    df_temp['week_start'] = df_temp[col_date].dt.to_period('W-SUN').apply(lambda r: r.start_time) # немного магии - тут надо начинать с пн

    # Группируем по программе и неделе
    weekly_counts = df_temp.groupby([col_program, 'week_start']).size().reset_index(name=col_values)

    # Получим все уникальные программы и все недели
    all_programs = weekly_counts[col_program].unique()
    all_weeks = pd.date_range(start=pd.Timestamp(year=2025, month=4, day=7, hour=0, minute=0, second=0),
                            end=weekly_counts['week_start'].max(),
                            freq='W-MON')  # каждую неделю по вторникам TODO: check различия MON & TUE

    # Создаем полную сетку: программа × неделя
    full_index = pd.MultiIndex.from_product([all_programs, all_weeks], names=[col_program, 'week_start'])
    full_df = pd.DataFrame(index=full_index).reset_index()

    # Объединяем с посчитанными заявками
    merged = pd.merge(full_df, weekly_counts, how='left', on=[col_program, 'week_start'])
    merged[col_values] = merged[col_values].fillna(0).astype(int)

    # Группируем по программе и объединяем значения в строку через ";"
    return merged.groupby(col_program)[col_values].apply(lambda x: ';'.join(map(str, x))).reset_index()

def find_first_file(mask: str, default: str, folder: str = "") -> str:
    file_list = glob.glob(folder + mask)
    if len(file_list) > 0:
        if file_list[0][0] != '~':
            return file_list[0]
        else: return file_list[1]
    else:
        return folder + default


def process_current_files(debug=None):

    if debug is None:
        import warnings
        # Suppress the FutureWarning
        warnings.simplefilter(action='ignore', category=FutureWarning)
        #pd.set_option('future.no_silent_downcasting', True)

    NEEDED_APPLICATIONS_RATIO = 45 / 100 #percents

    # папки и файлы для загрузки
    relative_folder = "data/"
    templates_folder = "templates/"

    programs_file = "programs.xlsx"
    template_file = "template.xlsx"

    # dashboard_file = "dashboard.xlsx"

    bitrix_file = find_first_file('*стади*.xls*', "bitrix.xls", relative_folder)

    bitrix_file_before_april = "bitrix_2024-10-01_2025-03-31.xlsx"
    portal_file = find_first_file('*порт*.xls*', "portal.xls", relative_folder)

    master_file = find_first_file('*асав*.xls*', "asav.xlsx", relative_folder)
    master_file_foreign = find_first_file('*инос*.xls*', "asav_foreign.xlsx", relative_folder)
    # master_file_sheet_name = "только онлайн"

    bachelor_app_file = find_first_file('*заявл*.xls*', "bac_applications.xlsx", relative_folder)
    bachelor_con_file = find_first_file('*дог*.xls*', "bac_contracts.xlsx", relative_folder)
    bachelor_enr_file = find_first_file('*зач*.xls*', "bac_enrolled.xlsx", relative_folder)

    enr_file = relative_folder + "зачисленные.xlsx" #find_first_file('*зач*.xls*', "bac_enrolled.xlsx", relative_folder)

    # считывание файлов
    try:
        # cчитываем базу данных програм
        print("Начинаем считывать базу программ")
        df_online_programs = pd.read_excel(templates_folder + programs_file)
        df_online_programs = df_online_programs[df_online_programs['format'] != 'offline'].reset_index(drop=True)
        df_online_master_programs = df_online_programs[df_online_programs['level'] == 'master'].drop(columns=["format"]).sort_values(by=col_program).reset_index(drop=True)
        df_online_bachelor_programs = df_online_programs[df_online_programs['level'] == 'bachelor'].drop(columns=["format"]).sort_values(by=col_program).reset_index(drop=True)
        print("База программ обработана")
    except:
        print("Потерялся " + programs_file + " - нужна база программ")
        return "Error program database"


    try:# Шаблон дашборда / template
        print("Начинаем считывать шаблон дашборда")
        # создаем дашборд по магистратурам добавляя туда программы из базы
        df_master_dashboard = pd.read_excel(templates_folder + template_file)
        df_master_dashboard = pd.concat([df_online_master_programs, df_master_dashboard])
        df_master_dashboard['program_bitrix'] = df_master_dashboard['program_bitrix'].fillna("")
        df_master_dashboard = df_master_dashboard.fillna(0)
        df_master_dashboard[[col_plan_rus, col_plan_foreign]] = df_master_dashboard[[col_plan_rus, col_plan_foreign]].astype(int)

        df_bachelor_dashboard = pd.read_excel(templates_folder + template_file)
        df_bachelor_dashboard = pd.concat([df_online_bachelor_programs, df_bachelor_dashboard])
        df_bachelor_dashboard['program_bitrix'] = df_bachelor_dashboard['program_bitrix'].fillna("")
        df_bachelor_dashboard = df_bachelor_dashboard.fillna(0)
        df_bachelor_dashboard[[col_plan_rus, col_plan_foreign]] = df_bachelor_dashboard[[col_plan_rus, col_plan_foreign]].astype(int)
        print("Шаблон дашборда считан")
    except:
        print("Потерялся " + template_file + " - без него дашборд не собрать")
        return "Error dashboard template"

    try:# Число лидов со studyonline с 1 апреля по настоящее время. Почему-то это html таблица, хотя файл xls
        print("Начинаем считывать данные от Битрикса в html-формате")
        df_bitrix_after_april = pd.read_html(bitrix_file, header=0)[0]
        df_bitrix_after_april[col_programs_names].fillna(main_studyonline, inplace=True)
        print("Данные от Битрикса считаны")
        # pd.read_excel(bitrix_file)
    except Exception as e:
        print(e)
        try:# Число лидов со studyonline с 1 апреля по настоящее время. На случай, если html чтение не сработало
            print("Начинаем считывать данные от Битрикса в xls-формате")
            df_bitrix_after_april = pd.read_excel(bitrix_file, header=0)
            df_bitrix_after_april[col_programs_names].fillna(main_studyonline, inplace=True)
            print("Данные от Битрикса считаны")
            # pd.read_excel(bitrix_file)
        except Exception as e:
            print(e)
            try:# Число лидов со studyonline с 1 апреля по настоящее время. На случай, если html чтение не сработало
                print("Начинаем считывать данные от Битрикса в xlsx-формате")
                df_bitrix_after_april = pd.read_excel(bitrix_file + 'x', header=0)
                df_bitrix_after_april[col_programs_names].fillna(main_studyonline, inplace=True)
                print("Данные от Битрикса считаны")
                # pd.read_excel(bitrix_file)
            except Exception as e:
                print(e)
                print("Нет выгрузки из Битрикса или она называется не " + bitrix_file)
                df_bitrix_after_april = pd.DataFrame()



    try:# Число лидов c портала c 1 октября по настоящее время. Почему-то это html таблица, хотя файл xls
        print("Начинаем считывать данные от Портала")
        df_portal = pd.read_html(portal_file, header=0)[0]
        df_portal[col_programs_names].fillna(main_studyonline, inplace=True)
        print("Данные от Портала считаны")
        # pd.read_excel(bitrix_file)
    except:
        print("Нет выгрузки заявок с Портала или она называется не " + portal_file)
        df_portal = pd.DataFrame()

    try:# Число лидов из битрикс до 1 апреля (не включительно). Почему-то это html таблица, хотя файл xls
        print("Начинаем считывать данные от Битрикса до 31.03")
        df_bitrix_before_april = pd.read_excel(templates_folder + bitrix_file_before_april, usecols="I:N")
        df_bitrix_before_april[col_programs_names].fillna(main_studyonline, inplace=True)
        print("Данные от Битрикса до 31.03 считаны")
        # pd.read_excel(bitrix_file)
    except:
        print("Нет выгрузки заявок из Битрикса до 31.03 или она называется не " + bitrix_file_before_april)
        df_bitrix_before_april = pd.DataFrame()

    try:
        leads_after_april = df_bitrix_after_april.groupby(col_programs_names)[col_programs_names].count()
        leads_after_april = pd.DataFrame({'program_bitrix':leads_after_april.index, 'values':leads_after_april.values})
    except:
        leads_after_april = pd.DataFrame(columns=['program_bitrix', 'values'])

    try:
        leads_before_april = df_bitrix_before_april.groupby(col_programs_names)[col_programs_names].count()
        leads_before_april = pd.DataFrame({'program_bitrix':leads_before_april.index, 'values':leads_before_april.values})
    except:
        leads_before_april = pd.DataFrame(columns=['program_bitrix', 'values'])

    #     # достаем данные по ЛК, договорам, оплатам и зачислениям из АСАВ
    # master_applications = df_master.groupby(master_col_programs)[master_col_programs].count() #.rename("program")#.sort_values(ascending=False)
    # master_applications = pd.DataFrame({col_program:master_applications.index, 'values':master_applications.values})
    # df_master_dashboard[col_applications_since_april] = insert_values(df_master_dashboard, master_applications, col_program, col_applications_since_april)


    # try:
    #     df_master_2024 = pd.read_csv(templates_folder + master_file_before_april, usecols="J:T")
    #     master_applications_2024 = df_master_2024.groupby(master_col_programs)[master_col_programs].count()
    #     master_applications_2024 = pd.DataFrame({col_program:master_applications_2024.index, 'values':master_applications_2024.values})
    #     df_master_dashboard[col_applications] = insert_values(df_master_dashboard, master_applications_2024, col_program, col_applications)
    # except:
    #     print("Ошибка при чтении файла asav_2024-10-01_2025-03-31.csv")
    #     df_master_dashboard[col_applications] = 0


    try:
        leads_portal = df_portal.groupby(col_programs_names)[col_programs_names].count()
        leads_portal = pd.DataFrame({'program_bitrix':leads_portal.index, 'values':leads_portal.values})
    except:
        leads_portal = pd.DataFrame(columns=['program_bitrix', 'values'])


    df_master_dashboard  [col_leads] = insert_values(df_master_dashboard,   leads_before_april, 'program_bitrix', col_leads)
    df_bachelor_dashboard[col_leads] = insert_values(df_bachelor_dashboard, leads_before_april, 'program_bitrix', col_leads)

    df_master_dashboard  [col_leads_after_april] = insert_values(df_master_dashboard,   leads_after_april, 'program_bitrix', col_leads_after_april)
    df_bachelor_dashboard[col_leads_after_april] = insert_values(df_bachelor_dashboard, leads_after_april, 'program_bitrix', col_leads_after_april)

    df_master_dashboard  [col_leads] += df_master_dashboard  [col_leads_after_april] # TODO check
    df_bachelor_dashboard[col_leads] += df_bachelor_dashboard[col_leads_after_april]

    main_leads             = leads_before_april.loc[leads_before_april['program_bitrix'] == main_studyonline, 'values'].values[0]
    main_leads_after_april = leads_after_april.loc[leads_after_april['program_bitrix'] == main_studyonline, 'values'].values[0]
    main_leads            += main_leads_after_april # TODO check

    df_master_dashboard  [col_leads_partners] = insert_values(df_master_dashboard,   leads_portal, 'program_bitrix', col_leads_partners)
    df_bachelor_dashboard[col_leads_partners] = insert_values(df_bachelor_dashboard, leads_portal, 'program_bitrix', col_leads_partners)
    # main_leads_portal = leads_portal.loc[leads_portal['program_bitrix'] == main_studyonline, 'values'].values[0]


    # АСАВ иностранцы
    try:
        print("Начинаем считывать данные от АСАВ по иностранцам")
        df_master_foreign = pd.read_excel(master_file_foreign, skiprows=1, usecols="F:BJ") #sheet_name=master_file_sheet_name,
        print("Данные от АСАВ по иностранцам считаны")
    except:
        print("Ошибка в обработке АСАВ по иностранцам, возможно нет выгрузки из АСАВ или она называется не " + master_file_foreign)
        df_master_foreign = pd.DataFrame(columns=[master_col_programs, master_foreign_col_contracts, master_foreign_col_payments, master_foreign_col_enrollments])


    df_master_foreign = process_foreign_programs(df_master_foreign, df_online_programs[col_program])

    master_applications_foreign = df_master_foreign.groupby(master_foreign_col_programs_1)[master_foreign_col_programs_1].count()
    master_applications_foreign = pd.DataFrame({col_program:master_applications_foreign.index, 'values':master_applications_foreign.values})
    df_master_dashboard[col_applications_foreign] = insert_values(df_master_dashboard, master_applications_foreign, col_program, col_applications_foreign)

    master_contracts_foreign = df_master_foreign[df_master_foreign[master_foreign_col_contracts] == "Да"].groupby(master_foreign_col_programs_1)[master_foreign_col_programs_1].count()
    master_contracts_foreign = pd.DataFrame({col_program:master_contracts_foreign.index, 'values':master_contracts_foreign.values})
    df_master_dashboard[col_contracts_foreign] = insert_values(df_master_dashboard, master_contracts_foreign, col_program, col_contracts_foreign)

    master_payments_foreign = df_master_foreign[df_master_foreign[master_foreign_col_payments] == "Да"].groupby(master_foreign_col_programs_1)[master_foreign_col_programs_1].count()
    master_payments_foreign = pd.DataFrame({col_program:master_payments_foreign.index, 'values':master_payments_foreign.values})
    df_master_dashboard[col_payments_foreign] = insert_values(df_master_dashboard, master_payments_foreign, col_program, col_payments_foreign)


#     master_applications = df_master.groupby(master_col_programs)[master_col_programs].count() #.rename("program")#.sort_values(ascending=False)
#     master_applications = pd.DataFrame({col_program:master_applications.index, 'values':master_applications.values})
#     df_master_dashboard[col_applications] = insert_values(df_master_dashboard, master_applications, col_program, col_applications)

#     master_contracts = df_master[df_master[master_col_contracts].notna()].groupby(master_col_programs)[master_col_programs].count()
#     master_contracts = pd.DataFrame({col_program:master_contracts.index, 'values':master_contracts.values})
#     df_master_dashboard[col_contracts] = insert_values(df_master_dashboard, master_contracts, col_program, col_contracts)

#     master_payments = df_master[df_master[master_col_payments] == "Оплачено"].groupby(master_col_programs)[master_col_programs].count()
#     master_payments = pd.DataFrame({col_program:master_payments.index, 'values':master_payments.values})
#     df_master_dashboard[col_payments] = insert_values(df_master_dashboard, master_payments, col_program, col_payments)

#     master_enrollments = df_master[df_master[master_col_enrollments].notna()].groupby(master_col_programs)[master_col_programs].count()
#     master_enrollments = pd.DataFrame({col_program:master_enrollments.index, 'values':master_enrollments.values})
#     df_master_dashboard[col_enrollments] = insert_values(df_master_dashboard, master_enrollments, col_program, col_enrollments)

#     master_male = df_master[df_master[col_gender_asav] == "Муж."].groupby(master_col_programs)[master_col_programs].count()
#     master_male = pd.DataFrame({col_program:master_male.index, 'values':master_male.values})
#     df_master_dashboard[col_male] = insert_values(df_master_dashboard, master_male, col_program, col_male)

#     master_female = df_master[df_master[col_gender_asav] == "Жен."].groupby(master_col_programs)[master_col_programs].count()
#     master_female = pd.DataFrame({col_program:master_female.index, 'values':master_female.values})
#     df_master_dashboard[col_female] = insert_values(df_master_dashboard, master_female, col_program, col_female)


#     df_master[col_birthday] = pd.to_datetime(df_master[col_birthday]).apply(num_years)
#     master_years_bars = df_master.groupby(master_col_programs)[col_birthday].apply(categorize_ages)
#     master_years_bars = pd.DataFrame({col_program:master_years_bars.index, 'values':master_years_bars.values})
#     df_master_dashboard[col_ages] = insert_values(df_master_dashboard, master_years_bars, col_program, col_ages)

#     master_years_mean = df_master.groupby(master_col_programs)[col_birthday].mean()
#     master_years_mean = pd.DataFrame({col_program:master_years_mean.index, 'values':master_years_mean.values})
#     df_master_dashboard[col_ages_mean] = insert_values(df_master_dashboard, master_years_mean, col_program, col_ages_mean)

# №№№№№№№№№№№


    # АСАВ
    try:
        print("Начинаем считывать данные от АСАВ")
        df_master = pd.read_excel(master_file, skiprows=1, usecols="A:AB, CY:DW, DZ") #sheet_name=master_file_sheet_name,
        df_master = df_master.dropna(how='all', ignore_index=True)
        print("Данные от АСАВ считаны")
    except:
        print("Ошибка в обработке АСАВ, возможно нет выгрузки из АСАВ или она называется не " + master_file)
        df_master = pd.DataFrame(columns=[master_col_programs, master_col_contracts, master_col_payments, master_col_enrollments])

    # убираем офлайн-треки и финансы из СПб
    df_master = df_master[~((df_master['Кампус конкурса'].str.contains("НИУ ВШЭ - Санкт-Петербург")) & (df_master[master_col_programs] == "Финансы")) ]
    df_master = df_master[~((df_master['Кампус конкурса'].str.contains("НИУ ВШЭ - Нижний Новгород")) & (df_master[master_col_programs] == "Финансы")) ]
    df_master['Магистерская специализация'] = df_master['Магистерская специализация'].fillna('')
    df_master = df_master[~df_master['Магистерская специализация'].str.contains("офлайн")]
    df_master = df_master[df_master["Основание зачисления/выбытия"] != "Завершение приемной кампании"]
    df_master = df_master.dropna(subset=[col_birthday])

    df_master = df_master.rename(columns={df_master.columns[-1]: 'gosuslugi', df_master.columns[-2]: 'applications_dates'})

    # достаем данные по ЛК, договорам, оплатам и зачислениям из АСАВ
    master_applications = df_master.groupby(master_col_programs)[master_col_programs].count() #.rename("program")#.sort_values(ascending=False)
    master_applications = pd.DataFrame({col_program:master_applications.index, 'values':master_applications.values})
    df_master_dashboard[col_applications] = insert_values(df_master_dashboard, master_applications, col_program, col_applications)

    master_contracts = df_master[df_master[master_col_contracts].notna()].groupby(master_col_programs)[master_col_programs].count()
    master_contracts = pd.DataFrame({col_program:master_contracts.index, 'values':master_contracts.values})
    df_master_dashboard[col_contracts] = insert_values(df_master_dashboard, master_contracts, col_program, col_contracts)

    master_payments = df_master[df_master[master_col_payments] == "Оплачено"].groupby(master_col_programs)[master_col_programs].count()
    master_payments = pd.DataFrame({col_program:master_payments.index, 'values':master_payments.values})
    df_master_dashboard[col_payments] = insert_values(df_master_dashboard, master_payments, col_program, col_payments)

    master_enrollments = df_master[df_master[master_col_enrollments].notna()].groupby(master_col_programs)[master_col_programs].count()
    master_enrollments = pd.DataFrame({col_program:master_enrollments.index, 'values':master_enrollments.values})
    df_master_dashboard[col_enrollments] = insert_values(df_master_dashboard, master_enrollments, col_program, col_enrollments)

    master_male = df_master[df_master[col_gender_asav] == "Муж."].groupby(master_col_programs)[master_col_programs].count()
    master_male = pd.DataFrame({col_program:master_male.index, 'values':master_male.values})
    df_master_dashboard[col_male] = insert_values(df_master_dashboard, master_male, col_program, col_male)

    master_female = df_master[df_master[col_gender_asav] == "Жен."].groupby(master_col_programs)[master_col_programs].count()
    master_female = pd.DataFrame({col_program:master_female.index, 'values':master_female.values})
    df_master_dashboard[col_female] = insert_values(df_master_dashboard, master_female, col_program, col_female)


    df_master[col_birthday] = pd.to_datetime(df_master[col_birthday]).apply(num_years)
    master_years_bars = df_master.groupby(master_col_programs)[col_birthday].apply(categorize_ages)
    master_years_bars = pd.DataFrame({col_program:master_years_bars.index, 'values':master_years_bars.values})
    df_master_dashboard[col_ages] = insert_values(df_master_dashboard, master_years_bars, col_program, col_ages)

    master_years_mean = df_master.groupby(master_col_programs)[col_birthday].mean()
    master_years_mean = pd.DataFrame({col_program:master_years_mean.index, 'values':master_years_mean.values})
    df_master_dashboard[col_ages_mean] = insert_values(df_master_dashboard, master_years_mean, col_program, col_ages_mean)



    print("Считаем регистрации и договоры по неделям")
    # считаем регистрации и договоры по неделям
    df_master_applications_by_week = process_by_week(df_master, master_col_programs, 'applications_dates', 'count', '%d.%m.%Y')
    df_master_applications_by_week = pd.DataFrame({col_program:df_master_applications_by_week[master_col_programs], 'values':df_master_applications_by_week['count']})
    df_master_dashboard[col_applications_by_week] = insert_values(df_master_dashboard, df_master_applications_by_week, col_program, col_applications_by_week)
    df_master['contracts_dates'] = df_master[master_col_contracts].str[-10:]
    df_master_contracts_by_week = process_by_week(df_master, master_col_programs, 'contracts_dates', 'count', '%Y-%m-%d')
    df_master_contracts_by_week = pd.DataFrame({col_program:df_master_contracts_by_week[master_col_programs], 'values':df_master_contracts_by_week['count']})
    df_master_dashboard[col_contracts_by_week] = insert_values(df_master_dashboard, df_master_contracts_by_week, col_program, col_contracts_by_week)


        # # считаем договоры по неделям
        # df_bachelor_contracts_by_week = process_by_week(df_bachelor_con, bachelor_col_programs_contracts, bachelor_col_date_contract, 'count', '%d.%m.%Y')
        # df_bachelor_contracts_by_week = df_bachelor_contracts_by_week.groupby(bachelor_col_programs_contracts)['count'].sum()
        # df_bachelor_contracts_by_week = df_bachelor_contracts_by_week.rename(index=bachelor_dict)
        # df_bachelor_contracts_by_week = pd.DataFrame({col_program:df_bachelor_contracts_by_week.index, 'values':df_bachelor_contracts_by_week.values})
        # df_bachelor_dashboard[col_contracts_by_week] = insert_values(df_bachelor_dashboard, df_bachelor_contracts_by_week, col_program, col_contracts_by_week)


    # АИС ПК
    # достаем данные по ЛК, договорам, оплатам и зачислениям из АИС ПК
    print("Начинаем считывать данные от АИС ПК")
    try:
        df_bachelor_app = pd.read_excel(bachelor_app_file) #, usecols="A,B,I:Z") #, sheet_name=master_file_sheet_name, skiprows=1, usecols="L:DT")
        bachelor_applications = df_bachelor_app.groupby(bachelor_col_programs)[bachelor_col_programs].count() #.rename("program")#.sort_values(ascending=False)
        bachelor_applications = pd.DataFrame({col_program:bachelor_applications.index, 'values':bachelor_applications.values})
        df_bachelor_dashboard[col_applications] = insert_values(df_bachelor_dashboard, bachelor_applications, col_program, col_applications)

        # считаем регистрации по неделям
        df_bachelor_applications_by_week = process_by_week(df_bachelor_app, bachelor_col_programs, bachelor_col_date)
        df_bachelor_applications_by_week = pd.DataFrame({col_program:df_bachelor_applications_by_week[bachelor_col_programs], 'values':df_bachelor_applications_by_week['count']})
        df_bachelor_dashboard[col_applications_by_week] = insert_values(df_bachelor_dashboard, df_bachelor_applications_by_week, col_program, col_applications_by_week)

    except pd.errors.EmptyDataError:
        print(bachelor_app_file + ' is empty')
    except FileNotFoundError:
        print(bachelor_app_file + ' file not found')


    try:
        df_bachelor_con = pd.read_excel(bachelor_con_file) #, usecols="J:V") #, sheet_name=master_file_sheet_name, skiprows=1, usecols="L:DT")
        bachelor_contracts = df_bachelor_con.groupby(col_programs_names)[col_programs_names].count()
        bachelor_contracts = bachelor_contracts.rename(index=bachelor_dict)
        bachelor_contracts = pd.DataFrame({col_program:bachelor_contracts.index, 'values':bachelor_contracts.values})
        df_bachelor_dashboard[col_contracts] = insert_values(df_bachelor_dashboard, bachelor_contracts, col_program, col_contracts)

        bachelor_payments = df_bachelor_con[(df_bachelor_con[bachelor_col_payments] == "Оплачен")|(df_bachelor_con[bachelor_col_payments] == "Оплачен по квитанциям")].groupby(col_programs_names)[col_programs_names].count()
        bachelor_payments = bachelor_payments.rename(index=bachelor_dict)
        bachelor_payments = pd.DataFrame({col_program:bachelor_payments.index, 'values':bachelor_payments.values})
        df_bachelor_dashboard[col_payments] = insert_values(df_bachelor_dashboard, bachelor_payments, col_program, col_payments)

        # считаем договоры по неделям
        df_bachelor_contracts_by_week = process_by_week(df_bachelor_con, bachelor_col_programs_contracts, bachelor_col_date_contract, 'count', '%d.%m.%Y')
        df_bachelor_contracts_by_week = df_bachelor_contracts_by_week.groupby(bachelor_col_programs_contracts)['count'].sum()
        df_bachelor_contracts_by_week = df_bachelor_contracts_by_week.rename(index=bachelor_dict)
        df_bachelor_contracts_by_week = pd.DataFrame({col_program:df_bachelor_contracts_by_week.index, 'values':df_bachelor_contracts_by_week.values})
        df_bachelor_dashboard[col_contracts_by_week] = insert_values(df_bachelor_dashboard, df_bachelor_contracts_by_week, col_program, col_contracts_by_week)

    except pd.errors.EmptyDataError:
        print(bachelor_con_file + ' is empty')
    except FileNotFoundError:
        print(bachelor_con_file + ' file not found')


    try:
        df_bachelor_enr = pd.read_excel(bachelor_enr_file, usecols="E:H") #, sheet_name=master_file_sheet_name, skiprows=1)
        print("Данные от АИС ПК считаны")

        bachelor_enrollments = df_bachelor_enr.groupby(bachelor_col_enrollments)[bachelor_col_enrollments].count()
        bachelor_enrollments = pd.DataFrame({col_program:bachelor_enrollments.index, 'values':bachelor_enrollments.values})
        df_bachelor_dashboard[col_enrollments] = insert_values(df_bachelor_dashboard, bachelor_enrollments, col_program, col_enrollments)

    except:
        print("Ошибка в обработке АИС ПК, возможно нет выгрузки из АИС ПК или она называется не:\n")
        print(bachelor_enr_file)
        # df_master = pd.DataFrame(columns=[master_col_programs, master_col_contracts, master_col_payments, master_col_enrollments])

    # расчет данных по госуслугам
    df_master_gosuslugi = df_master[df_master['gosuslugi'] == 'Администратор БД (SYSDBA)'].groupby(master_col_programs)['gosuslugi'].count()
    df_master_gosuslugi = pd.DataFrame({col_program:df_master_gosuslugi.index, 'values':df_master_gosuslugi.values})
    df_master_dashboard[col_applications_gosuslugi] = insert_values(df_master_dashboard, df_master_gosuslugi, col_program, col_applications_gosuslugi)

    df_bachelor_gosuslugi = df_bachelor_app[df_bachelor_app[bachelor_col_gosuslugi] == 'ЕПГУ'].groupby(bachelor_col_programs)[bachelor_col_gosuslugi].count()
    df_bachelor_gosuslugi = pd.DataFrame({col_program:df_bachelor_gosuslugi.index, 'values':df_bachelor_gosuslugi.values})
    df_bachelor_dashboard[col_applications_gosuslugi] = insert_values(df_bachelor_dashboard, df_bachelor_gosuslugi, col_program, col_applications_gosuslugi)




    # расчет данных прошлых лет
    df_history, df_leads_prev, df_applications_prev, df_contracts_prev = process_history_files() # Only for master for now

    masters_list = df_online_master_programs[col_program].unique()
    master_2025_no_duplicates = df_master[df_master[master_col_programs].isin(masters_list)].drop_duplicates(subset=[master_col_reg_number])
    bachelor_2025_no_duplicates = df_bachelor_app.drop_duplicates(subset=[bachelor_col_reg_number])

    df_history.loc[2025, 'applications_unique'] = master_2025_no_duplicates[master_col_programs].count() + bachelor_2025_no_duplicates[bachelor_col_programs].count()

    df_leads_prev = pd.DataFrame({col_program_bitrix:df_leads_prev.index, 'values':df_leads_prev.values})
    main_leads_after_april_prev = df_leads_prev[df_leads_prev[col_program_bitrix] == main_studyonline]['values'].values[0]

    df_main_dashboard = pd.DataFrame(columns=df_master_dashboard.columns)
    df_main_dashboard.loc[len(df_main_dashboard)] = {col_program: main_studyonline, col_program_bitrix: main_studyonline, col_leads: main_leads, col_leads_after_april: main_leads_after_april, col_leads_after_april_prev: main_leads_after_april_prev}
    df = pd.concat([df_main_dashboard, df_master_dashboard, df_bachelor_dashboard], ignore_index=True, sort=False)


    df_applications_prev = pd.DataFrame({col_program:df_applications_prev.index, 'values':df_applications_prev.values})
    df[col_applications_prev] = insert_values(df, df_applications_prev, col_program, col_applications_prev)

    df_contracts_prev = pd.DataFrame({col_program:df_contracts_prev.index, 'values':df_contracts_prev.values})
    df[col_contracts_prev] = insert_values(df, df_contracts_prev, col_program, col_contracts_prev)

    # считываем тренды по неделям (заявки)
    df_leads_by_week = process_by_week(df_bitrix_after_april, col_programs_names, bitrix_col_date)
    df_leads_by_week = pd.DataFrame({col_program_bitrix:df_leads_by_week[col_programs_names], 'values':df_leads_by_week['count']})
    df[col_leads_after_april_by_week] = insert_values(df, df_leads_by_week, col_program_bitrix, col_leads_after_april_by_week)

    df[col_leads_after_april_prev] = insert_values(df, df_leads_prev, col_program_bitrix, col_leads_after_april_prev)
    df = df.drop(columns=['program_bitrix'])

    df.fillna(0, inplace=True)

    # считаем зачисленных, если не посчитаны ранее
    try:
        df_enr = pd.read_excel(enr_file)
        print("Данные по зачисленным из базы считаны")
        df[col_enrollments] = insert_values(df, df_enr[[col_program, col_enrollments]], col_program, col_enrollments)
        df[col_enrollments_foreign] = insert_values(df, df_enr[[col_program, col_enrollments_foreign]], col_program, col_enrollments_foreign)
    except:
        print("Нет базы по зачисленным или она называется не:\n")
        print(enr_file)

    # считаем второстепенные столбцы
    df[col_leads_total]                            = df[col_leads_partners] + df[col_leads]
    df[col_conversion_leads_to_contracts]          = df[col_contracts] / df[col_leads_total]
    df[col_needed_applications]              = round(df[col_plan_rus]/ NEEDED_APPLICATIONS_RATIO)
    df[col_conversion_applications_to_contracts]   = df[col_contracts] / df[col_applications]
    df[col_conversion_contracts_to_payments]       = df[col_payments]  / df[col_contracts]
    df[col_conversion_contracts_to_enrollments]    = df[col_enrollments]  / df[col_contracts]
    df[col_payments_div_plan_rus]                  = df[col_payments]  / df[col_plan_rus]
    df[col_payments_div_plan_foreign]              = df[col_payments_foreign]  / df[col_plan_foreign]
    df[col_income_1year]                           = df['price'] * df[col_payments] / 1000 # from thousands to millions
    df.loc[df['level'] == 'master', col_income_all]   = df[col_income_1year] * 2
    df.loc[df['level'] == 'bachelor', col_income_all] = df[col_income_1year] * 4
    # df[col_income_all       ] = df[col_income_1year]  * (2 if df['level'] == 'master' else 4) # TODO check later
    df[col_income_1year_hse ] = df[col_income_1year] * df['income_percent'] / 100
    df[col_income_all_hse   ] = df[col_income_all]   * df['income_percent'] / 100

    df.replace(np.inf, 0, inplace=True)
    df.fillna(0, inplace=True)

    return df, df_history