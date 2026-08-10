import glob
import pandas as pd
import numpy as np

from col_names import *
from time_const import *

from general_pipeline import insert_values, process_by_week, num_years, categorize_ages


def load_dashboard_template(templates_folder: str) -> pd.DataFrame:
    programs_file = 'programs.xlsx'
    template_file = 'template.xlsx'
    df_online_programs = pd.read_excel(templates_folder + programs_file)
    df_online_programs = df_online_programs[df_online_programs['format'] != 'offline'].reset_index(drop=True)
    df_dashboard_template = pd.read_excel(templates_folder + template_file)
    return pd.concat([df_online_programs, df_dashboard_template], ignore_index=True, sort=False).fillna(0)

def _count_history_data(df: pd.DataFrame, now: datetime, delta: timedelta, col_count: str, col_id: str = "") -> int:
    '''Count the number of rows in the dataframe that meet the condition: date (in column col_count) + delta <= now. Also drop duplicates if non empty col_id is provided.'''
    df_temp = df.copy()
    df_temp = df_temp[df_temp[col_count] + delta <= now]

    if col_id in df_temp.columns:
        df_temp = df_temp.drop_duplicates(subset=[col_id])
    elif col_id != "":
        print(f'Column {col_id} not found in dataframe in _count_history_data')
        raise ValueError(f'Column {col_id} not found in dataframe')

    return df_temp[col_count].count()

def process_history_files(templates_folder: str = 'templates/'):

    now = datetime.now()

    master_leads_file_2023        = 'bitrix_2023-04-01_2023-09-15.csv'
    master_leads_file_2024        = 'bitrix_2024-04-01_2024-09-15.csv'
    master_leads_file_2025        = 'bitrix_2025-04-01_2025-09-15.csv'

    bitrix_file_2024              = 'bitrix_2024-04-01_2024-09-15.xlsx'
    bitrix_file_2025              = 'bitrix_2025-04-01_2025-09-15.xlsx'
    bitrix_file_2025_before_april = 'bitrix_2024-10-01_2025-03-31.xlsx'

    asav_file_2023                = 'asav_2023.xlsx'
    asav_file_2024                = 'asav_2024.xlsx'
    asav_file_2025                = 'asav_2025.xlsx'
    bachelor_file_2024            = 'bachelor_2024.xls'
    bachelor_file_2025            = 'bachelor_2025.xls'


    # TODO вписать сводные расчеты

    print('Начинаем считывать исторические данные')

    try:
        # applications_dates_2023 = pd.read_csv(templates_folder + master_applications_file_2023, parse_dates=[0], date_format='%d.%m.%Y')
        # applications_dates_2024 = pd.read_csv(templates_folder + master_applications_file_2024, parse_dates=[0], date_format='%d.%m.%Y')
        # contracts_dates_2023 = pd.read_csv(templates_folder + master_contracts_file_2023, parse_dates=[0], date_format='%d.%m.%Y')
        # contracts_dates_2024 = pd.read_csv(templates_folder + master_contracts_file_2024, parse_dates=[0], date_format='%d.%m.%Y')
        # TODO добавить разделение по датам до и после 1 апреляы
        leads_dates_2023 = pd.read_csv(templates_folder + master_leads_file_2023, parse_dates=[0], date_format='%d.%m.%Y')
        leads_dates_2024 = pd.read_csv(templates_folder + master_leads_file_2024, parse_dates=[0], date_format='%d.%m.%Y')
        leads_dates_2025 = pd.read_csv(templates_folder + master_leads_file_2025, parse_dates=[0], date_format='%d.%m.%Y')

        print('Даты по лидам считаны')

        leads_dates_2024_by_program = pd.read_excel(templates_folder + bitrix_file_2024, usecols='J:N') #, parse_dates=[0], date_format='%d.%m.%Y  %hh:%mm:%ss')
        leads_dates_2024_by_program[leads_dates] = pd.to_datetime(leads_dates_2024_by_program[leads_dates], errors='coerce', format='%d.%m.%Y  %hh:%mm:%ss')
        leads_dates_2024_by_program[col_programs_names] = leads_dates_2024_by_program[col_programs_names].fillna(main_studyonline)

        print('Лиды в привязке к программам 2024 считаны')

        leads_dates_2025_by_program = pd.read_excel(templates_folder + bitrix_file_2025) #, parse_dates=[0], date_format='%d.%m.%Y  %hh:%mm:%ss')
        leads_dates_2025_by_program[leads_dates] = pd.to_datetime(leads_dates_2025_by_program[leads_dates], errors='coerce', format='%d.%m.%Y  %hh:%mm:%ss')
        leads_dates_2025_by_program[col_programs_names] = leads_dates_2025_by_program[col_programs_names].fillna(main_studyonline)

        leads_dates_2025_before_april_by_program = pd.read_excel(templates_folder + bitrix_file_2025_before_april) #, parse_dates=[0], date_format='%d.%m.%Y  %hh:%mm:%ss')
        leads_dates_2025_before_april_by_program[leads_dates] = pd.to_datetime(leads_dates_2025_before_april_by_program[leads_dates], errors='coerce', format='%d.%m.%Y  %hh:%mm:%ss')
        leads_dates_2025_before_april_by_program[col_programs_names] = leads_dates_2025_before_april_by_program[col_programs_names].fillna(main_studyonline)


        print('Лиды в привязке к программам 2025 считаны')

        bachelor_2024 = pd.read_excel(templates_folder + bachelor_file_2024) #, usecols='A:H,J:AB')
        bachelor_2024[applications_dates] = pd.to_datetime(bachelor_2024[applications_dates], errors='coerce', format='%d.%m.%Y')
        bachelor_2024[contracts_dates]    = pd.to_datetime(bachelor_2024[contracts_dates],    errors='coerce', format='%d.%m.%Y')

        print('Данные АИС ПК 2024 года считаны')

        bachelor_2025 = pd.read_excel(templates_folder + bachelor_file_2025) #, usecols='A:H,J:AB')
        bachelor_2025[applications_dates] = pd.to_datetime(bachelor_2025[applications_dates], errors='coerce', format='%d.%m.%Y')
        bachelor_2025[contracts_dates]    = pd.to_datetime(bachelor_2025[contracts_dates],    errors='coerce', format='%d.%m.%Y')

        print('Данные АИС ПК 2025 года считаны')

        asav_2023 = pd.read_excel(templates_folder + asav_file_2023, parse_dates=[0, 1], skiprows=1, date_format='%d.%m.%Y')
        asav_2023[applications_dates] = pd.to_datetime(asav_2023[applications_dates], format='%Y-%m-%d 00:00:00')
        asav_2023[contracts_dates] = pd.to_datetime(asav_2023[contracts_dates], errors='coerce', format='%d.%m.%Y')

        print('Данные АСАВ 2023 года считаны')

        asav_2024 = pd.read_excel(templates_folder + asav_file_2024, parse_dates=[0, 1], skiprows=1, date_format='%d.%m.%Y')
        asav_2024[applications_dates] = pd.to_datetime(asav_2024[applications_dates], format='%Y-%m-%d 00:00:00')
        asav_2024[contracts_dates] = pd.to_datetime(asav_2024[contracts_dates], errors='coerce', format='%d.%m.%Y')

        print('Данные АСАВ 2024 года считаны')

        asav_2025 = pd.read_excel(templates_folder + asav_file_2025, parse_dates=[0, 1], skiprows=1, date_format='%d.%m.%Y')
        asav_2025[applications_dates] = pd.to_datetime(asav_2025[applications_dates], format='%Y-%m-%d 00:00:00') # CHECK
        asav_2025[contracts_dates] = pd.to_datetime(asav_2025[contracts_dates], errors='coerce', format='%d.%m.%Y')

        asav_2025 = asav_2025[~((asav_2025[master_col_campus].str.contains('НИУ ВШЭ - Санкт-Петербург')) & (asav_2025[master_col_programs] == 'Финансы')) ]
        asav_2025 = asav_2025[~((asav_2025[master_col_campus].str.contains('НИУ ВШЭ - Нижний Новгород')) & (asav_2025[master_col_programs] == 'Финансы')) ]
        # asav_2025[master_col_program_specialization] = asav_2025[master_col_program_specialization].fillna('')
        # asav_2025 = asav_2025[~asav_2025[master_col_program_specialization].str.contains('офлайн')]

        print('Данные АСАВ 2025 года считаны')

    except:
        print('Files of previous years are not founded or have errors')
        print(master_leads_file_2023)
        print(master_leads_file_2024)
        print(master_leads_file_2025)
        print(bitrix_file_2024)
        print(bitrix_file_2025)
        print(bitrix_file_2025_before_april)
        print(asav_file_2023)
        print(asav_file_2024)
        print(asav_file_2025)
        print(bachelor_2024)
        print(bachelor_2025)

    delta_now_2023 = timedelta(days=365+366+365)
    delta_now_2024 = timedelta(days=365+365)
    delta_now_2025 = timedelta(days=365)

    asav_2025_no_rossokhins = asav_2025[~asav_2025[master_col_programs].str.startswith('Психоанализ и')]

    applications_2025_masters_unique   = _count_history_data(asav_2025,     now, delta_now_2025, applications_dates, col_id_asav)
    applications_2025_bachelors_unique = _count_history_data(bachelor_2025, now, delta_now_2025, applications_dates, col_id_bachelor)
    applications_2025_masters_no_rossokhins_unique = _count_history_data(asav_2025_no_rossokhins, now, delta_now_2025, applications_dates, col_id_asav)

    df_pivot = pd.DataFrame.from_dict({'leads' :
                                {2023: _count_history_data(leads_dates_2023, now, delta_now_2023, 'leads_dates'),
                                 2024: _count_history_data(leads_dates_2024, now, delta_now_2024, 'leads_dates'),
                                 2025: _count_history_data(leads_dates_2025, now, delta_now_2025, 'leads_dates')},
                                'applications' :
                                {2023: _count_history_data(asav_2023, now, delta_now_2023, applications_dates),
                                 2024: _count_history_data(asav_2024, now, delta_now_2024, applications_dates) + _count_history_data(bachelor_2024, now, delta_now_2024, applications_dates),
                                 2025: _count_history_data(asav_2025, now, delta_now_2025, applications_dates) + _count_history_data(bachelor_2025, now, delta_now_2025, applications_dates)},

                                'contracts' :
                                {2023: _count_history_data(asav_2023, now, delta_now_2023, contracts_dates),
                                 2024: _count_history_data(asav_2024, now, delta_now_2024, contracts_dates) + _count_history_data(bachelor_2024, now, delta_now_2024, contracts_dates),
                                 2025: _count_history_data(asav_2025, now, delta_now_2025, contracts_dates) + _count_history_data(bachelor_2025, now, delta_now_2025, contracts_dates)},

                                'applications_unique' :
                                {2023: _count_history_data(asav_2023, now, delta_now_2023, applications_dates, col_id_asav),
                                 2024: _count_history_data(asav_2024, now, delta_now_2024, applications_dates, col_id_asav) + _count_history_data(bachelor_2024, now, delta_now_2024, applications_dates, col_id_bachelor),
                                 2025: applications_2025_masters_unique + applications_2025_bachelors_unique},

                                'applications_masters_unique' :
                                {2025: applications_2025_masters_unique},

                                'applications_bachelors_unique' :
                                {2025: applications_2025_bachelors_unique},

                                'applications_no_rossokhins_unique' :
                                {2025: applications_2025_masters_no_rossokhins_unique + applications_2025_bachelors_unique},

                                'applications_masters_no_rossokhins_unique' :
                                {2025: applications_2025_masters_no_rossokhins_unique},
                                })

    df_leads_after_april_prev = leads_dates_2025_by_program[leads_dates_2025_by_program[leads_dates] + delta_now_2025 <= now].groupby(col_programs_names)[col_programs_names].count()
    df_leads_all_prev         = df_leads_after_april_prev.add(leads_dates_2025_before_april_by_program[leads_dates_2025_before_april_by_program[leads_dates] + delta_now_2025 <= now].groupby(col_programs_names)[col_programs_names].count(), fill_value=0)

    df_applications_prev = pd.concat([asav_2025[asav_2025[applications_dates] + delta_now_2025 <= now].groupby(master_col_programs)[master_col_programs].count(),
                                     bachelor_2025[bachelor_2025[applications_dates] + delta_now_2025 <= now].groupby(bachelor_col_programs)[bachelor_col_programs].count()])
    df_contracts_prev    = pd.concat([asav_2025[asav_2025[contracts_dates] + delta_now_2025 <= now].groupby(master_col_programs)[master_col_programs].count(),
                                     bachelor_2025[bachelor_2025[contracts_dates] + delta_now_2025 <= now].groupby(bachelor_col_programs)[bachelor_col_programs].count()])


    print('Исторические данные считаны')
    return df_pivot, df_leads_all_prev, df_leads_after_april_prev, df_applications_prev, df_contracts_prev

def _process_foreign_programs(df, programs_names):
    try:
        df[master_foreign_col_programs_2] = df[master_foreign_col_programs_2].fillna('')
        is_online = df[master_foreign_col_programs_1].isin(programs_names)
        for i, row in df.iterrows():
            if not is_online.loc[i] or row[master_foreign_col_faculty_1] == 'Факультет Санкт-Петербургская школа экономики и менеджмента' or row[master_foreign_col_faculty_1] == 'Факультет экономики':
                df.loc[i, master_foreign_col_programs_1] = df.loc[i, master_foreign_col_programs_2]

        # df[master_foreign_col_programs_1] = df[master_foreign_col_programs_1] if df[master_foreign_col_programs_1].isin(programs_names) and df[master_foreign_col_faculty_1] != 'Факультет Санкт-Петербургская школа экономики и менеджмента' else df[master_foreign_col_programs_2]
        # df[master_foreign_col_programs_1].fillna(df[master_foreign_col_programs_2])
        df = df[df[master_foreign_col_programs_1].isin(programs_names)]
        # df[master_foreign_col_faculty_1] = df[master_foreign_col_faculty_1].fillna('')
        # df[master_foreign_col_faculty_2] = df[master_foreign_col_faculty_2].fillna('')
        # # df[master_foreign_col_program_final] = df[master_foreign_col_program_final].fillna('')
        # # df[master_foreign_col_faculty_final] = df[master_foreign_col_faculty_final].fillna('')
        # df[col_program] = df[master_foreign_col_programs_1] + df[master_foreign_col_programs_2]
    except:
        print('Problem with foreign programs file')
    return df

def _find_first_file(mask: str, default: str, folder: str = '') -> str:
    file_list = glob.glob(folder + mask)
    if len(file_list) > 0:
        if file_list[0].find('~') == -1:
            return file_list[0]
        else:
            return file_list[1]
    else:
        return folder + default

def _preprocess_bitrix_file(df: pd.DataFrame) -> pd.DataFrame:
    df[col_programs_names].fillna(main_studyonline, inplace=True)

    namings_to_drop = ['тест тест', 'richkos test', '-дубль', 'тест тест-дубль', 'тест тесст-дубль' ] # приведены к нижнему регистру
    for name in namings_to_drop:
        # df = df.query('@bitrix_col_contact.str.contains')
        df = df[df[bitrix_col_contact].str.lower() != name]
        df = df[df[bitrix_col_deal_name].str.lower() != name]

    # костыль от переименования коллегами названий в битрексе по ходу ПК, можно придумать как исправить в TODO
    df.loc[df[col_programs_names] == 'ИНТДИЗ. Интерактивный дизайн / Москва / 540401 Дизайн / факультет креативных индустрий / Магистратура', col_programs_names] = 'ИНТДИЗ. Интерактивный дизайн'

    return df

def process_from_current_files(debug=None):
    '''Legacy function for processing data from current files - Bitrix, ASAV, AISPK exports to xsl(x)'''

    if debug is None:
        import warnings
        # Suppress the FutureWarning
        warnings.simplefilter(action='ignore', category=FutureWarning)
        #pd.set_option('future.no_silent_downcasting', True)

    NEEDED_APPLICATIONS_RATIO = 45 / 100 #percents

    # папки и файлы для загрузки
    relative_folder = 'data/'
    templates_folder = 'templates/'

    programs_file = 'programs.xlsx'
    template_file = 'template.xlsx'

    # dashboard_file = 'dashboard.xlsx'

    bitrix_file = _find_first_file('*DEAL*.xls*', 'bitrix.xls', relative_folder)

    bitrix_file_before_april = 'bitrix_2025-10-01_2026-03-31.xlsx' # TODO объединить за счет получения данных с помощью API
    portal_file = _find_first_file('*порт*.xls*', 'portal.xls', relative_folder)

    master_file = _find_first_file('*асав*.xls*', 'asav.xlsx', relative_folder)
    master_file_foreign = _find_first_file('*инос*.xls*', 'asav_foreign.xlsx', relative_folder)

    master_file_early_invitation = _find_first_file('*РП*.xls*', 'asav_early_invitation.xlsx', relative_folder)
    # master_file_sheet_name = 'только онлайн'

    bachelor_app_file = _find_first_file('*заявл*.xls*', 'bac_applications.xlsx', relative_folder)
    bachelor_con_file = _find_first_file('*дог*.xls*', 'bac_contracts.xlsx', relative_folder)
    bachelor_enr_file = _find_first_file('*зач*.xls*', 'bac_enrolled.xlsx', relative_folder)

    enr_file = relative_folder + 'зачисленные.xlsx' #find_first_file('*зач*.xls*', 'bac_enrolled.xlsx', relative_folder)

    # считывание базовых файлов
    try:
        # cчитываем базу данных програм
        print('Начинаем считывать базу программ')
        df_online_programs = pd.read_excel(templates_folder + programs_file)
        df_online_programs = df_online_programs[df_online_programs['format'] != 'offline'].reset_index(drop=True)
        df_online_master_programs = df_online_programs[df_online_programs['level'] == 'master'].drop(columns=['format']).sort_values(by=col_program).reset_index(drop=True)
        df_online_bachelor_programs = df_online_programs[df_online_programs['level'] == 'bachelor'].drop(columns=['format']).sort_values(by=col_program).reset_index(drop=True)
        print('База программ обработана')
    except:
        print('Потерялся ' + programs_file + ' - нужна база программ')
        return 'Error program database'


    try:# Шаблон дашборда / template
        print('Начинаем считывать шаблон дашборда')
        # создаем дашборд по магистратурам добавляя туда программы из базы
        df_master_dashboard = pd.read_excel(templates_folder + template_file)
        df_master_dashboard = pd.concat([df_online_master_programs, df_master_dashboard])
        df_master_dashboard[col_program_bitrix] = df_master_dashboard[col_program_bitrix].fillna('')
        df_master_dashboard = df_master_dashboard.fillna(0)
        df_master_dashboard[[col_plan_rus, col_plan_foreign]] = df_master_dashboard[[col_plan_rus, col_plan_foreign]].astype(int)

        df_bachelor_dashboard = pd.read_excel(templates_folder + template_file)
        df_bachelor_dashboard = pd.concat([df_online_bachelor_programs, df_bachelor_dashboard])
        df_bachelor_dashboard[col_program_bitrix] = df_bachelor_dashboard[col_program_bitrix].fillna('')
        df_bachelor_dashboard = df_bachelor_dashboard.fillna(0)
        df_bachelor_dashboard[[col_plan_rus, col_plan_foreign]] = df_bachelor_dashboard[[col_plan_rus, col_plan_foreign]].astype(int)
        print('Шаблон дашборда считан')
    except:
        print('Потерялся ' + template_file + ' - без него дашборд не собрать')
        return 'Error dashboard template'

    # TODO заменить на API
    now = datetime.now()


    if (now >= DATE_01_04_2026):
        try: # Число лидов со studyonline с 1 апреля по настоящее время. Почему-то это html таблица, хотя файл xls
            print('Начинаем считывать данные от Битрикса в html-формате')
            df_bitrix_after_april = pd.read_html(bitrix_file, header=0)[0]
            df_bitrix_after_april = _preprocess_bitrix_file(df_bitrix_after_april)
            df_bitrix_after_april[bitrix_col_date] = pd.to_datetime(df_bitrix_after_april[bitrix_col_date], dayfirst=True, errors='raise') # , format='%d.%m.%Y  %H:%M'

            df_bitrix_after_april = df_bitrix_after_april[df_bitrix_after_april[bitrix_col_date] >= DATE_01_04_2026] # надо отфильтровать с началом от 1.04, иначе может быть дублирование лидов с апреля и далее

            print('Данные от Битрикса считаны')
            # pd.read_excel(bitrix_file)
        except Exception as e:
            print(e)
            # try:# Число лидов со studyonline с 1 апреля по настоящее время. На случай, если html чтение не сработало
            #     print('Начинаем считывать данные от Битрикса в xls-формате')
            #     df_bitrix_after_april = pd.read_excel(bitrix_file, header=0)
            #     df_bitrix_after_april[col_programs_names].fillna(main_studyonline, inplace=True)
            #     print('Данные от Битрикса считаны')
            #     # pd.read_excel(bitrix_file)
            # except Exception as e:
            #     print(e)
            #     try:# Число лидов со studyonline с 1 апреля по настоящее время. На случай, если html чтение не сработало
            #         print('Начинаем считывать данные от Битрикса в xlsx-формате')
            #         df_bitrix_after_april = pd.read_excel(bitrix_file + 'x', header=0)
            #         df_bitrix_after_april[col_programs_names].fillna(main_studyonline, inplace=True)
            #         print('Данные от Битрикса считаны')
            #         # pd.read_excel(bitrix_file)
            #     except Exception as e:
            #         print(e)
            #         print('Нет выгрузки из Битрикса или она называется не ' + bitrix_file)
    else:
        df_bitrix_after_april = pd.DataFrame(columns=[bitrix_col_date, col_programs_names])

    # TODO заменить с помощью API

    if (now >= DATE_01_04_2026):
        try:# Число лидов из битрикс до 1 апреля (не включительно). Почему-то это html таблица, хотя файл xls
            print('Начинаем считывать данные от Битрикса до 31.03')
            df_bitrix_before_april = pd.read_excel(templates_folder + bitrix_file_before_april) # , usecols=columns_from_bitrix_file_2026= H:Q
            df_bitrix_before_april = _preprocess_bitrix_file(df_bitrix_before_april)
            print('Данные от Битрикса до 31.03 считаны')
            # pd.read_excel(bitrix_file)
        except:
            print('Нет выгрузки заявок из Битрикса до 31.03 или она называется не ' + bitrix_file_before_april)
    else:
        try: # Число лидов со studyonline с 1 октября по настоящее время. Почему-то это html таблица, хотя файл xls
            print('Начинаем считывать данные от Битрикса в html-формате')
            df_bitrix_before_april = pd.read_html(bitrix_file, header=0)[0]
            df_bitrix_before_april = _preprocess_bitrix_file(df_bitrix_before_april)
            print('Данные от Битрикса считаны')
            # pd.read_excel(bitrix_file)
        except Exception as e:
            print(e)
            try:# Число лидов со studyonline с 1 октября по настоящее время. На случай, если html чтение не сработало
                print('Начинаем считывать данные от Битрикса в xls-формате')
                df_bitrix_before_april = pd.read_excel(bitrix_file, header=0)
                df_bitrix_before_april = _preprocess_bitrix_file(df_bitrix_before_april)
                print('Данные от Битрикса считаны')
                # pd.read_excel(bitrix_file)
            except Exception as e:
                print(e)
                try:# Число лидов со studyonline с 1 октября по настоящее время. На случай, если html чтение не сработало
                    print('Начинаем считывать данные от Битрикса в xlsx-формате')
                    df_bitrix_before_april = pd.read_excel(bitrix_file + 'x', header=0)
                    df_bitrix_before_april = _preprocess_bitrix_file(df_bitrix_before_april)
                    print('Данные от Битрикса считаны')
                    # pd.read_excel(bitrix_file)
                except Exception as e:
                    print(e)
                    print('Нет выгрузки из Битрикса или она называется не ' + bitrix_file)


    try:# Число лидов c портала c 1 октября по настоящее время. Почему-то это html таблица, хотя файл xls
        print('Начинаем считывать данные от Портала')
        df_portal = pd.read_html(portal_file, header=0)[0]
        df_portal[col_programs_names].fillna(main_studyonline, inplace=True)
        print('Данные от Портала считаны')
        # pd.read_excel(bitrix_file)
    except:
        print('Нет выгрузки заявок с Портала или она называется не ' + portal_file)
        df_portal = pd.DataFrame()


    try:
        leads_after_april = df_bitrix_after_april.groupby(col_programs_names)[col_programs_names].count()
        leads_after_april = pd.DataFrame({col_program_bitrix:leads_after_april.index, 'values':leads_after_april.values})
    except:
        leads_after_april = pd.DataFrame(columns=[col_program_bitrix, 'values'])

    try:
        leads_before_april = df_bitrix_before_april.groupby(col_programs_names)[col_programs_names].count()
        leads_before_april = pd.DataFrame({col_program_bitrix:leads_before_april.index, 'values':leads_before_april.values})
    except:
        leads_before_april = pd.DataFrame(columns=[col_program_bitrix, 'values'])

    try:
        leads_portal = df_portal.groupby(col_programs_names)[col_programs_names].count()
        leads_portal = pd.DataFrame({col_program_bitrix:leads_portal.index, 'values':leads_portal.values})
    except:
        leads_portal = pd.DataFrame(columns=[col_program_bitrix, 'values'])


    df_master_dashboard  [col_leads] = insert_values(df_master_dashboard,   leads_before_april, col_program_bitrix, col_leads)
    df_bachelor_dashboard[col_leads] = insert_values(df_bachelor_dashboard, leads_before_april, col_program_bitrix, col_leads)

    df_master_dashboard  [col_leads_after_april] = insert_values(df_master_dashboard,   leads_after_april, col_program_bitrix, col_leads_after_april)
    df_bachelor_dashboard[col_leads_after_april] = insert_values(df_bachelor_dashboard, leads_after_april, col_program_bitrix, col_leads_after_april)

    df_master_dashboard  [col_leads] += df_master_dashboard  [col_leads_after_april] # adding leads after april to leads before april to count sum
    df_bachelor_dashboard[col_leads] += df_bachelor_dashboard[col_leads_after_april]

    main_leads = leads_before_april.loc[leads_before_april[col_program_bitrix] == main_studyonline, 'values'].values[0]

    if now >= DATE_01_04_2026:
        try:
            main_leads_after_april = leads_after_april.loc[leads_after_april[col_program_bitrix] == main_studyonline, 'values'].values[0]
            main_leads            += main_leads_after_april # TODO check
        except:
            print('Problem with main_leads_after_april')
    else:
        main_leads_after_april = 0

    df_master_dashboard  [col_leads_partners] = insert_values(df_master_dashboard,   leads_portal, col_program_bitrix, col_leads_partners)
    df_bachelor_dashboard[col_leads_partners] = insert_values(df_bachelor_dashboard, leads_portal, col_program_bitrix, col_leads_partners)
    # main_leads_portal = leads_portal.loc[leads_portal[col_program_bitrix] == main_studyonline, 'values'].values[0]

    # АСАВ раннее приглашение
    try:
        print('Начинаем считывать данные от АСАВ по РП')
        df_master_early = pd.read_excel(master_file_early_invitation, skiprows=1)
        df_master_early = df_master_early.rename(columns={df_master_early.columns[1]: col_id_asav})
        print('Данные от АСАВ по РП считаны')
    except:
        print('Ошибка в обработке АСАВ по РП, возможно нет выгрузки из АСАВ или она называется не ' + master_file_early_invitation)
        df_master_early = pd.DataFrame(columns=[col_id_asav, col_programs_names, col_gender_asav])



    # АСАВ иностранцы
    try:
        print('Начинаем считывать данные от АСАВ по иностранцам')
        df_master_foreign = pd.read_excel(master_file_foreign, skiprows=1, usecols='F:BJ') #sheet_name=master_file_sheet_name,
        print('Данные от АСАВ по иностранцам считаны')
    except:
        print('Ошибка в обработке АСАВ по иностранцам, возможно нет выгрузки из АСАВ или она называется не ' + master_file_foreign)
        df_master_foreign = pd.DataFrame(columns=[master_col_programs, master_foreign_col_contracts, master_foreign_col_payments, master_foreign_col_enrollments])

    df_master_foreign = _process_foreign_programs(df_master_foreign, df_online_programs[col_program])

    try:
        master_applications_foreign = df_master_foreign.groupby(master_foreign_col_programs_1)[master_foreign_col_programs_1].count()
        master_applications_foreign = pd.DataFrame({col_program:master_applications_foreign.index, 'values':master_applications_foreign.values})
        df_master_dashboard[col_applications_foreign] = insert_values(df_master_dashboard, master_applications_foreign, col_program, col_applications_foreign)
    except:
        print('Problem with foreign applications')
        df_master_dashboard[col_applications_foreign] = 0

    try:
        master_contracts_foreign = df_master_foreign[df_master_foreign[master_foreign_col_contracts] == 'Да'].groupby(master_foreign_col_programs_1)[master_foreign_col_programs_1].count()
        master_contracts_foreign = pd.DataFrame({col_program:master_contracts_foreign.index, 'values':master_contracts_foreign.values})
        df_master_dashboard[col_contracts_foreign] = insert_values(df_master_dashboard, master_contracts_foreign, col_program, col_contracts_foreign)
    except:
        print('Problem with foreign contracts')
        df_master_dashboard[col_contracts_foreign] = 0

    try:
        master_payments_foreign = df_master_foreign[df_master_foreign[master_foreign_col_payments] == 'Да'].groupby(master_foreign_col_programs_1)[master_foreign_col_programs_1].count()
        master_payments_foreign = pd.DataFrame({col_program:master_payments_foreign.index, 'values':master_payments_foreign.values})
        df_master_dashboard[col_payments_foreign] = insert_values(df_master_dashboard, master_payments_foreign, col_program, col_payments_foreign)
    except:
        print('Problem with foreign payments')
        df_master_dashboard[col_payments_foreign] = 0

    # АСАВ
    try:
        print('Начинаем считывать данные от АСАВ')
        df_master = pd.read_excel(master_file, skiprows=1) #sheet_name=master_file_sheet_name, #, usecols='A:AB, CY:DW, DZ'
        df_master = df_master.dropna(how='all', ignore_index=True)
        df_master = df_master.rename(columns={df_master.columns[-2]: applications_dates})
        print('Данные от АСАВ считаны')
    except:
        print('Ошибка в обработке АСАВ, возможно нет выгрузки из АСАВ или она называется не ' + master_file)
        df_master = pd.DataFrame(columns=[master_col_programs, master_col_contracts, master_col_payments, master_col_enrollments, master_col_campus, master_col_program_specialization, col_birthday, applications_dates])

    # убираем офлайн-треки и финансы из СПб
    df_master = df_master[~((df_master[master_col_campus].str.contains('НИУ ВШЭ - Санкт-Петербург')) & (df_master[master_col_programs] == 'Финансы')) ]
    df_master = df_master[~((df_master[master_col_campus].str.contains('НИУ ВШЭ - Нижний Новгород')) & (df_master[master_col_programs] == 'Финансы')) ]
    df_master[master_col_program_specialization] = df_master[master_col_program_specialization].fillna('')
    df_master = df_master[~df_master[master_col_program_specialization].str.contains('офлайн')]

    df_master = df_master.dropna(subset=[col_birthday]) # удаляем пустые строки

    # TODO uncomment later
    # df_master_applications_by_week = process_by_week(df_master, master_col_programs, applications_dates, 'count', '%d.%m.%Y') # TODO check - здесь отфильтровывался костыль из-за изменений в АСАВ 31.08
    #df_master = df_master[df_master['Основание зачисления/выбытия'] != 'Завершение приемной кампании'] # TODO check - здесь отфильтровывался костыль из-за изменений в АСАВ 31.08


    # считаем подачи РП
    master_early = df_master_early.groupby(col_programs_names)[col_programs_names].count() #.rename('program')#.sort_values(ascending=False)
    master_early = pd.DataFrame({col_program:master_early.index, 'values':master_early.values})
    df_master_dashboard[col_early_invitation] = insert_values(df_master_dashboard, master_early, col_program, col_early_invitation)


    # достаем данные по ЛК, договорам, оплатам и зачислениям из АСАВ
    master_applications = df_master.groupby(master_col_programs)[master_col_programs].count() #.rename('program')#.sort_values(ascending=False)
    master_applications = pd.DataFrame({col_program:master_applications.index, 'values':master_applications.values})
    df_master_dashboard[col_applications] = insert_values(df_master_dashboard, master_applications, col_program, col_applications)

    master_contracts = df_master[df_master[master_col_contracts].notna()].groupby(master_col_programs)[master_col_programs].count()
    master_contracts = pd.DataFrame({col_program:master_contracts.index, 'values':master_contracts.values})
    df_master_dashboard[col_contracts] = insert_values(df_master_dashboard, master_contracts, col_program, col_contracts)

    master_payments = df_master[df_master[master_col_payments] == 'Оплачено'].groupby(master_col_programs)[master_col_programs].count()
    master_payments = pd.DataFrame({col_program:master_payments.index, 'values':master_payments.values})
    df_master_dashboard[col_payments] = insert_values(df_master_dashboard, master_payments, col_program, col_payments)

    master_enrollments = df_master[df_master[master_col_enrollments].notna()].groupby(master_col_programs)[master_col_programs].count()
    master_enrollments = pd.DataFrame({col_program:master_enrollments.index, 'values':master_enrollments.values})
    df_master_dashboard[col_enrollments] = insert_values(df_master_dashboard, master_enrollments, col_program, col_enrollments)

    # TODO проверить 20.06
    if (now >= datetime(year=2026, month=6, day=20)):
        master_male = df_master[df_master[col_gender_asav] == 'Муж.'].groupby(master_col_programs)[master_col_programs].count()
        master_male = pd.DataFrame({col_program:master_male.index, 'values':master_male.values})
        df_master_dashboard[col_male] = insert_values(df_master_dashboard, master_male, col_program, col_male)

        master_female = df_master[df_master[col_gender_asav] == 'Жен.'].groupby(master_col_programs)[master_col_programs].count()
        master_female = pd.DataFrame({col_program:master_female.index, 'values':master_female.values})
        df_master_dashboard[col_female] = insert_values(df_master_dashboard, master_female, col_program, col_female)

        df_master[col_birthday] = pd.to_datetime(df_master[col_birthday]).apply(num_years)
        master_years_bars = df_master.groupby(master_col_programs)[col_birthday].apply(categorize_ages)
        master_years_bars = pd.DataFrame({col_program:master_years_bars.index, 'values':master_years_bars.values})
        df_master_dashboard[col_ages] = insert_values(df_master_dashboard, master_years_bars, col_program, col_ages)

        master_years_mean = df_master.groupby(master_col_programs)[col_birthday].mean()
        master_years_mean = pd.DataFrame({col_program:master_years_mean.index, 'values':master_years_mean.values})
        df_master_dashboard[col_ages_mean] = insert_values(df_master_dashboard, master_years_mean, col_program, col_ages_mean)

    else: # временная версия с данными из РП:
        master_male = df_master_early[df_master_early[col_gender_asav] == 'Муж.'].groupby(col_programs_names)[col_programs_names].count()
        master_male = pd.DataFrame({col_program:master_male.index, 'values':master_male.values})
        df_master_dashboard[col_male] = insert_values(df_master_dashboard, master_male, col_program, col_male)

        master_female = df_master_early[df_master_early[col_gender_asav] == 'Жен.'].groupby(col_programs_names)[col_programs_names].count()
        master_female = pd.DataFrame({col_program:master_female.index, 'values':master_female.values})
        df_master_dashboard[col_female] = insert_values(df_master_dashboard, master_female, col_program, col_female)

        df_master_early[col_birthday] = pd.to_datetime(df_master_early[col_birthday], dayfirst=True, errors='coerce').apply(num_years)
        master_years_bars = df_master_early.groupby(col_programs_names)[col_birthday].apply(categorize_ages)
        master_years_bars = pd.DataFrame({col_program:master_years_bars.index, 'values':master_years_bars.values})
        df_master_dashboard[col_ages] = insert_values(df_master_dashboard, master_years_bars, col_program, col_ages)

        master_years_mean = df_master_early.groupby(col_programs_names)[col_birthday].mean()
        master_years_mean = pd.DataFrame({col_program:master_years_mean.index, 'values':master_years_mean.values})
        df_master_dashboard[col_ages_mean] = insert_values(df_master_dashboard, master_years_mean, col_program, col_ages_mean)

    print('Считаем регистрации и договоры по неделям')
    # считаем регистрации и договоры по неделям TODO uncomment later
    # df_master_applications_by_week = process_by_week(df_master, master_col_programs, applications_dates, 'count', '%d.%m.%Y') # сделано раньше как костыль из-за изменений в АСАВ 31.08
    # try:
    #     df_master_applications_by_week = pd.DataFrame({col_program:df_master_applications_by_week[master_col_programs], 'values':df_master_applications_by_week['count']})
    # except:
    #     print('Problem with master applications by week')
    #     df_master_applications_by_week = pd.DataFrame(columns=[col_program, col_applications_by_week])
    # df_master_dashboard[col_applications_by_week] = insert_values(df_master_dashboard, df_master_applications_by_week, col_program, col_applications_by_week)
    df_master[contracts_dates] = df_master[master_col_contracts].str[-10:]
    # TODO uncomment later
    # df_master_contracts_by_week = process_by_week(df_master, master_col_programs, contracts_dates, 'count', '%Y-%m-%d')
    # try:
    #     df_master_contracts_by_week = pd.DataFrame({col_program:df_master_contracts_by_week[master_col_programs], 'values':df_master_contracts_by_week['count']})
    # except:
    #     print('Problem with master contracts by week')
    #     df_master_contracts_by_week = pd.DataFrame(columns=[col_program, col_contracts_by_week])
    # df_master_dashboard[col_contracts_by_week] = insert_values(df_master_dashboard, df_master_contracts_by_week, col_program, col_contracts_by_week)


    # АИС ПК
    # достаем данные по ЛК, договорам, оплатам и зачислениям из АИС ПК
    print('Начинаем считывать данные от АИС ПК')
    try:
        df_bachelor_app = pd.read_excel(bachelor_app_file) #, usecols='A,B,I:Z') #, sheet_name=master_file_sheet_name, skiprows=1, usecols='L:DT')
        bachelor_applications = df_bachelor_app.groupby(bachelor_col_programs)[bachelor_col_programs].count() #.rename('program')#.sort_values(ascending=False)
        bachelor_applications = pd.DataFrame({col_program:bachelor_applications.index, 'values':bachelor_applications.values})
        df_bachelor_dashboard[col_applications] = insert_values(df_bachelor_dashboard, bachelor_applications, col_program, col_applications)

        # считаем регистрации по неделям TODO uncomment later
        # df_bachelor_applications_by_week = process_by_week(df_bachelor_app, bachelor_col_programs, bachelor_col_date, 'count', '%d/%m/%Y %H:%M:%S')
        # df_bachelor_applications_by_week = pd.DataFrame({col_program:df_bachelor_applications_by_week[bachelor_col_programs], 'values':df_bachelor_applications_by_week['count']})
        # df_bachelor_dashboard[col_applications_by_week] = insert_values(df_bachelor_dashboard, df_bachelor_applications_by_week, col_program, col_applications_by_week)

    except pd.errors.EmptyDataError:
        print(bachelor_app_file + ' is empty')
    except FileNotFoundError:
        print(bachelor_app_file + ' file not found')


    try:
        df_bachelor_con = pd.read_excel(bachelor_con_file) #, usecols='J:V') #, sheet_name=master_file_sheet_name, skiprows=1, usecols='L:DT')
        bachelor_contracts = df_bachelor_con.groupby(col_programs_names)[col_programs_names].count()
        bachelor_contracts = bachelor_contracts.rename(index=bachelor_dict)
        bachelor_contracts = pd.DataFrame({col_program:bachelor_contracts.index, 'values':bachelor_contracts.values})
        df_bachelor_dashboard[col_contracts] = insert_values(df_bachelor_dashboard, bachelor_contracts, col_program, col_contracts)

        bachelor_payments = df_bachelor_con[(df_bachelor_con[bachelor_col_payments] == 'Оплачен')|(df_bachelor_con[bachelor_col_payments] == 'Оплачен по квитанциям')].groupby(col_programs_names)[col_programs_names].count()
        bachelor_payments = bachelor_payments.rename(index=bachelor_dict)
        bachelor_payments = pd.DataFrame({col_program:bachelor_payments.index, 'values':bachelor_payments.values})
        df_bachelor_dashboard[col_payments] = insert_values(df_bachelor_dashboard, bachelor_payments, col_program, col_payments)

        # считаем договоры по неделям TODO uncomment later
        # df_bachelor_contracts_by_week = process_by_week(df_bachelor_con, bachelor_col_programs_contracts, bachelor_col_date_contract, 'count', '%d.%m.%Y')
        # df_bachelor_contracts_by_week = df_bachelor_contracts_by_week.groupby(bachelor_col_programs_contracts)['count'].sum()
        # df_bachelor_contracts_by_week = df_bachelor_contracts_by_week.rename(index=bachelor_dict)
        # df_bachelor_contracts_by_week = pd.DataFrame({col_program:df_bachelor_contracts_by_week.index, 'values':df_bachelor_contracts_by_week.values})
        # df_bachelor_dashboard[col_contracts_by_week] = insert_values(df_bachelor_dashboard, df_bachelor_contracts_by_week, col_program, col_contracts_by_week)

    except pd.errors.EmptyDataError:
        print(bachelor_con_file + ' is empty')
    except FileNotFoundError:
        print(bachelor_con_file + ' file not found')


    try:
        df_bachelor_enr = pd.read_excel(bachelor_enr_file, usecols='E:H') #, sheet_name=master_file_sheet_name, skiprows=1)
        print('Данные от АИС ПК считаны')

        bachelor_enrollments = df_bachelor_enr.groupby(bachelor_col_enrollments)[bachelor_col_enrollments].count()
        bachelor_enrollments = pd.DataFrame({col_program:bachelor_enrollments.index, 'values':bachelor_enrollments.values})
        df_bachelor_dashboard[col_enrollments] = insert_values(df_bachelor_dashboard, bachelor_enrollments, col_program, col_enrollments)

    except:
        print('Ошибка в обработке АИС ПК, возможно нет выгрузки из АИС ПК или она называется не:\n')
        print(bachelor_enr_file)
        # df_master = pd.DataFrame(columns=[master_col_programs, master_col_contracts, master_col_payments, master_col_enrollments])

    # расчет данных прошлых лет
    df_history, df_leads_prev, df_leads_after_april_prev, df_applications_prev, df_contracts_prev = process_history_files()

    masters_list = df_online_master_programs[col_program].unique()
    master_2026_no_duplicates = df_master[df_master[master_col_programs].isin(masters_list)].drop_duplicates(subset=[master_col_reg_number])
    try:
        bachelor_2026_no_duplicates = df_bachelor_app.drop_duplicates(subset=[bachelor_col_reg_number])
    except:
        bachelor_2026_no_duplicates = pd.DataFrame(columns=[bachelor_col_programs])

    df_history.loc[2026, 'applications_unique'] = master_2026_no_duplicates[master_col_programs].count() + bachelor_2026_no_duplicates[bachelor_col_programs].count()
    df_history.loc[2026, 'early_invitations_unique'] = df_master_early[df_master_early[col_programs_names].isin(masters_list)].drop_duplicates(subset=[col_id_asav])[col_programs_names].count()


    df_leads_prev = pd.DataFrame({col_program_bitrix:df_leads_prev.index, 'values':df_leads_prev.values})
    df_leads_after_april_prev = pd.DataFrame({col_program_bitrix:df_leads_after_april_prev.index, 'values':df_leads_after_april_prev.values})

    try: #TODO change to date comparison from try
        main_leads_after_april_prev = df_leads_after_april_prev[df_leads_after_april_prev[col_program_bitrix] == main_studyonline]['values'].values[0]
    except:
        main_leads_after_april_prev = 0

    try: #TODO change to date comparison from try
        main_leads_prev = df_leads_prev[df_leads_prev[col_program_bitrix] == main_studyonline]['values'].values[0]
    except:
        main_leads_prev = 0

    df_main_dashboard = pd.DataFrame(columns=df_master_dashboard.columns)
    df_main_dashboard.loc[len(df_main_dashboard)] = {col_program: main_studyonline,
                                                     col_program_bitrix: main_studyonline,
                                                     col_leads: main_leads,
                                                     col_leads_prev : main_leads_prev,
                                                     col_leads_after_april: main_leads_after_april,
                                                     col_leads_after_april_prev: main_leads_after_april_prev}
    df = pd.concat([df_main_dashboard, df_master_dashboard, df_bachelor_dashboard], ignore_index=True, sort=False)


    df_applications_prev = pd.DataFrame({col_program:df_applications_prev.index, 'values':df_applications_prev.values})
    df[col_applications_prev] = insert_values(df, df_applications_prev, col_program, col_applications_prev)

    df_contracts_prev = pd.DataFrame({col_program:df_contracts_prev.index, 'values':df_contracts_prev.values})
    df[col_contracts_prev] = insert_values(df, df_contracts_prev, col_program, col_contracts_prev)

    # считываем тренды по неделям (заявки)
    if now > DATE_01_04_2026:
        df_bitrix_before_april.rename(columns={leads_dates: bitrix_col_date}, inplace=True)

    df_bitrix = pd.concat([df_bitrix_before_april, df_bitrix_after_april])

    # TODO uncomment later
    # df_leads_by_week = process_by_week(df_bitrix, col_programs_names, bitrix_col_date)
    # df_leads_by_week = pd.DataFrame({col_program_bitrix:df_leads_by_week[col_programs_names], 'values':df_leads_by_week['count']})
    # df[col_leads_by_week] = insert_values(df, df_leads_by_week, col_program_bitrix, col_leads_by_week)

    df[col_leads_after_april_prev] = insert_values(df, df_leads_after_april_prev, col_program_bitrix, col_leads_after_april_prev)
    df[col_leads_prev] = insert_values(df, df_leads_prev, col_program_bitrix, col_leads_prev)
    df = df.drop(columns=[col_program_bitrix, 'tg_chat_id', 'campus', 'start_year'])

    df.fillna(0, inplace=True)

    # считаем зачисленных, если не посчитаны ранее
    try:
        df_enr = pd.read_excel(enr_file)
        print('Данные по зачисленным из базы считаны')
        df[col_enrollments] = insert_values(df, df_enr[[col_program, col_enrollments]], col_program, col_enrollments)
        df[col_enrollments_foreign] = insert_values(df, df_enr[[col_program, col_enrollments_foreign]], col_program, col_enrollments_foreign)
    except:
        df[col_enrollments] = 0
        df[col_enrollments_foreign] = 0
        print('Нет базы по зачисленным или она называется не:\n')
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

def _get_reg_number(reg_numbers: dict[int, int], epgu_id: int) -> int:
    asav_id = reg_numbers.get(int(epgu_id), -1)
    if asav_id == -1:
        print(f'Не найден регистрационный номер для ЕПГУ ID: {epgu_id}')
    return asav_id

def process_exams_dates_file(exams_dates_path: str = 'data/ВИ.xlsx', reg_numbers_path: str = 'data/АСАВ.xlsx', dict_path: str = 'templates/entrance_exams_dict.csv') -> dict[str, pd.DataFrame]:
    '''Обрабатывает файл выгрузки дат ВИ в словарь DataFrames по предметам.

    Каждая вкладка (ключ словаря) соответствует одному уникальному "Предмет".
    Внутри вкладки колонки — уникальные "Начало сдачи" по возрастанию,
    под каждой колонкой — список "Код в ЕПГУ".

    Args:
        exams_dates_path: путь к файлу выгрузки дат ВИ (выгрузка данных из АСАВ по датам ВИ из ССПВО)
        reg_numbers_path: путь к файлу с регистрационными номерами (выгрузка абитуриентов из АСАВ)
        dict_path: путь к файлу с словарем онлайн-программ (конкурс-программа)

    Returns:
        Словарь {название_предмета: DataFrame}, где колонки DataFrame — даты начала сдачи,
        а строки — коды ЕПГУ.
    '''
    print('Начинаем считывать данные по онлайн-программам')
    exams_dict = pd.read_csv(dict_path, sep=';').set_index(col_exams_subject).to_dict(orient='dict')['Программа']

    print('Начинаем считывать данные заявлений из АСАВ')

    reg_numbers = pd.read_excel(reg_numbers_path, usecols='A:B', skiprows=1) #.drop_duplicates()
    col_asav_id = reg_numbers.columns[0]
    col_epgu_id = reg_numbers.columns[1]
    reg_numbers = reg_numbers.set_index(col_epgu_id).to_dict(orient='dict')[col_asav_id]

    print('Начинаем считывать данные дат ВИ')


    df = pd.read_excel(exams_dates_path, usecols='A:S', skiprows=1)
    col_despatch_id = df.columns[0]                                             # Столбец A — "Идентификатор DespatchID"
    col_exam_status = df.columns[6]                                             # Столбец G — "Запись на ВИ или Отказ"
    df = df.dropna(subset=[col_exams_subject, col_exams_start])                 # удаляем строки, где нет предмета или даты начала сдачи
    df = df[df[col_exams_subject].isin(exams_dict.keys())]                      # выбираем только те записи на ВИ, которые относятся к онлайн-программам
    df = df[df[col_exam_status] == 'Да'].sort_values(by=col_despatch_id, ascending=False, ignore_index=True) # сортируем по идентификатору DespatchID, чтобы выше были последние записи
    df = df.drop_duplicates(subset=[col_exams_subject, col_exams_epgu], ignore_index=True)                    # удаляем дубликаты по предмету и коду ЕПГУ, оставляя только последнюю записи


    print(f'Данные дат ВИ считаны: {len(df)} записей')

    result = {}
    num_dates = {}
    for subject, group in df.groupby(col_exams_subject):
        grouped = group.groupby(col_exams_start)[col_exams_epgu].apply(
            lambda x: [_get_reg_number(reg_numbers, v) for v in x.dropna().tolist()]
        )
        grouped = grouped.sort_index()

        max_len = max(len(v) for v in grouped.values) if len(grouped) > 0 else 0
        data = {}
        for date, codes in grouped.items():
            num_dates[date] = num_dates.get(date, 0) + len(codes)
            padded = codes + [None] * (max_len - len(codes))
            data[date.strftime('%d.%m.%Y %H:%M')] = padded

        result[subject] = pd.DataFrame(data)

    print(f'Обработано предметов: {len(result)}')
    print(f'Распределение записей по датам: {num_dates}')

    return result, exams_dict


