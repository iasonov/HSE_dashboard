import pandas as pd

def insert_values(df_dashboard, df_values, col_join, col_values): # df_values should have "values" column
    df_dashboard.loc[df_dashboard[col_join].isin(df_values[col_join]), col_values] = df_values.loc[df_values[col_join].isin(df_dashboard[col_join]), 'values'].values
    df_dashboard[col_values] = df_dashboard[col_values].fillna(0).astype(int)
    return df_dashboard[col_values]

def process_excel_files():
    import numpy as np

    # папки и файлы для загрузки
    relative_folder = "data/"

    programs_file = "programs.xlsx"

    dashboard_file = "dashboard.xlsx"

    bitrix_file = "bitrix.xls"

    master_file = "asav.xlsx"
    master_file_sheet_name = "только онлайн"
    master_col_programs = "Конкурс в магистратуру"
    master_col_contracts = "Договор на обучение" # не пусто
    master_col_payments = "Оплата первого периода" # "Оплачено"
    master_col_enrollments = "Приказ о зачислении" # не пусто


    bachelor_app_file = "bac_applications.xls"

    bachelor_con_file = "bac_contracts.xls"

    bachelor_enr_file = "bac_enrolled.xlsx"

    bachelor_col_programs = "Конкурсная группа"
    bachelor_col_programs_names = "Образовательная программа"
    # bachelor_col_contracts = "Образовательная программа"
    bachelor_col_payments = "Статус оплаты" # "Оплачен" или "Оплачен по квитанциям"
    bachelor_col_enrollments = "Конкурс"

    main_studyonline = "Общий лендинг"

    template_file = "template.xlsx"


    # основные параметры
    col_program = 'program'
    col_plan = 'plan_rus'
    col_leads = "Общее кол-во заявок (studyonline) ВСЕГО"
    col_applications = """Кол-во регистраций в ЛК абитуриента (РФ 1 и 2 приоритет)"""
    col_contracts = """Договоры (ПК) РФ"""
    col_payments = """Оплаты (ПК)"""
    col_enrollments = "Зачисленные (ПК)"


    # дополнительные и расчетные параметры
    col_leads_total = """Общее кол-во заявок ВСЕГО (портал+РК)"""
    col_leads_partners = """Общее кол-во заявок  c  hse.ru по кнопкам/партнерских стр"""
    col_conversion_leads_to_contracts = """Конверсия заявка -> договор (без учета заявок с общего ленда)"""
    col_needed_applications = """Необходимо регистраций в ЛК для обеспечения набора (30% поступили от регистрации в АСАВ)"""
    NEEDED_APPLICATIONS_RATIO = 30 / 100 #percents
    col_conversion_applications_to_contracts = """Конверсия ЛК -> договор"""
    col_conversion_contracts_to_payments = """Конверсия договор -> оплата"""
    col_conversion_contracts_to_enrollments = """Конверсия договор -> зачисление"""
    col_payments_div_plan = """Выполнение плана %"""
    col_income_1year     = "Выручка за 1 год, млн.руб."
    col_income_all       = "Выручка за весь период обучения"
    col_income_1year_hse = "Выручка за 1 год после отчислений партнерам"
    col_income_all_hse   = "Выручка за весь период обучения после отчислений партнерам"


    # считывание файлов

    try:
        # cчитываем базу данных програм
        print("Начинаем считывать базу программ")
        df_online_programs = pd.read_excel(relative_folder + programs_file)
        df_online_programs = df_online_programs[df_online_programs['format'] != 'offline'].reset_index(drop=True)
        #df_online_programs[['plan_rus', 'plan_foreign']] = df_online_programs[['plan_rus', 'plan_foreign']].astype(int)
        df_online_master_programs = df_online_programs[df_online_programs['level'] == 'master'].drop(columns=["format"]).sort_values(by=col_program).reset_index(drop=True)
        df_online_bachelor_programs = df_online_programs[df_online_programs['level'] == 'bachelor'].drop(columns=["format"]).sort_values(by=col_program).reset_index(drop=True)
        print("База программ обработана")
    except:
        print("Потерялся " + programs_file + " - нужна база программ")
        return "Error program database"


    try:# Шаблон дашборда / template
        print("Начинаем считывать шаблон дашборда")
        # создаем дашборд по магистратурам добавляя туда программы из базы
        df_master_dashboard = pd.read_excel(relative_folder + template_file)
        df_master_dashboard = pd.concat([df_online_master_programs, df_master_dashboard])
        df_master_dashboard['program_bitrix'] = df_master_dashboard['program_bitrix'].fillna("")
        df_master_dashboard = df_master_dashboard.fillna(0)
        df_master_dashboard[['plan_rus', 'plan_foreign']] = df_master_dashboard[['plan_rus', 'plan_foreign']].astype(int)

        df_bachelor_dashboard = pd.read_excel(relative_folder + template_file)
        df_bachelor_dashboard = pd.concat([df_online_bachelor_programs, df_bachelor_dashboard])
        df_bachelor_dashboard['program_bitrix'] = df_bachelor_dashboard['program_bitrix'].fillna("")
        df_bachelor_dashboard = df_bachelor_dashboard.fillna(0)
        df_bachelor_dashboard[['plan_rus', 'plan_foreign']] = df_bachelor_dashboard[['plan_rus', 'plan_foreign']].astype(int)
        print("Шаблон дашборда считан")
    except:
        print("Потерялся " + template_file + " - без него дашборд не собрать")
        return "Error dashboard template"

    try:# Число лидов. Почему-то это html таблица, хотя файл xls
        print("Начинаем считывать данные от Битрикса")
        df_bitrix = pd.read_html(relative_folder + bitrix_file, header=0)[0]
        df_bitrix['Образовательная программа'].fillna(main_studyonline, inplace=True)
        print("Данные от Битрикса считаны")
        # pd.read_excel(relative_folder + bitrix_file)
    except:
        print("Нет выгрузки из Битрикса или она называется не " + bitrix_file)
        df_bitrix = pd.DataFrame()

    try:
        leads = df_bitrix.groupby('Образовательная программа')['Образовательная программа'].count()
        # leads = leads * 20 #only for test! REMOVE LATER
        leads = pd.DataFrame({'program_bitrix':leads.index, 'values':leads.values})
    except:
        leads = pd.DataFrame(columns=['program_bitrix', 'values'])
    df_master_dashboard[col_leads]   = insert_values(df_master_dashboard,   leads, 'program_bitrix', col_leads)
    df_bachelor_dashboard[col_leads] = insert_values(df_bachelor_dashboard, leads, 'program_bitrix', col_leads)
    main_leads = leads.loc[leads['program_bitrix'] == main_studyonline, 'values'].values[0]

    # АСАВ
    try:
        print("Начинаем считывать данные от АСАВ")
        df_master = pd.read_excel(relative_folder + master_file, sheet_name=master_file_sheet_name, skiprows=1, usecols="L:DT")
        print("Данные от АСАВ считаны")
    except:
        print("Ошибка в обработке АСАВ, возможно нет выгрузки из АСАВ или она называется не " + master_file)
        df_master = pd.DataFrame(columns=[master_col_programs, master_col_contracts, master_col_payments, master_col_enrollments])

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




    # АИС ПК
    try:
        print("Начинаем считывать данные от АИС ПК")
        df_bachelor_app = pd.read_excel(relative_folder + bachelor_app_file, usecols="I:Z") #, sheet_name=master_file_sheet_name, skiprows=1, usecols="L:DT")
        df_bachelor_con = pd.read_excel(relative_folder + bachelor_con_file, usecols="H:T") #, sheet_name=master_file_sheet_name, skiprows=1, usecols="L:DT")
        df_bachelor_enr = pd.read_excel(relative_folder + bachelor_enr_file, usecols="E:H") #, sheet_name=master_file_sheet_name, skiprows=1)
        print("Данные от АИС ПК считаны")

        bachelor_dict = {
    'Глобальные цифровые коммуникации (онлайн)'                        :'Глобальные цифровые коммуникации - онлайн (О К)'        ,
    'Глобальные цифровые коммуникации (онлайн)'                        :'Глобальные цифровые коммуникации (Медиа) - онлайн (О К)',
    'Компьютерные науки и анализ данных (онлайн)'                      :'Компьютерные науки и анализ данных - онлайн (О К)'      ,
    'Экономический анализ (онлайн)'                                    :'Экономический анализ - онлайн (О К)'                    ,
    'Дизайн  (онлайн)'                                                 :'Дизайн - онлайн (О К)'                                  ,
    'Программные системы и автоматизация процессов разработки (онлайн)':'Программные системы и автоматизация процессов разработки - онлайн (О К)'
        }
    except:
        print("Ошибка в обработке АИС ПК, возможно нет выгрузки из АИС ПК или она называется не:\n" + bachelor_app_file)
        print(bachelor_con_file)
        print(bachelor_enr_file)
        # df_master = pd.DataFrame(columns=[master_col_programs, master_col_contracts, master_col_payments, master_col_enrollments])

    # достаем данные по ЛК, договорам, оплатам и зачислениям из АСАВ
    bachelor_applications = df_bachelor_app.groupby(bachelor_col_programs)[bachelor_col_programs].count() #.rename("program")#.sort_values(ascending=False)
    bachelor_applications = pd.DataFrame({col_program:bachelor_applications.index, 'values':bachelor_applications.values})
    df_bachelor_dashboard[col_applications] = insert_values(df_bachelor_dashboard, bachelor_applications, col_program, col_applications)

    bachelor_contracts = df_bachelor_con.groupby(bachelor_col_programs_names)[bachelor_col_programs_names].count()
    bachelor_contracts = bachelor_contracts.rename(index=bachelor_dict)
    bachelor_contracts = pd.DataFrame({col_program:bachelor_contracts.index, 'values':bachelor_contracts.values})
    df_bachelor_dashboard[col_contracts] = insert_values(df_bachelor_dashboard, bachelor_contracts, col_program, col_contracts)

    bachelor_payments = df_bachelor_con[(df_bachelor_con[bachelor_col_payments] == "Оплачен")|(df_bachelor_con[bachelor_col_payments] == "Оплачен по квитанциям")].groupby(bachelor_col_programs_names)[bachelor_col_programs_names].count()
    bachelor_payments = bachelor_payments.rename(index=bachelor_dict)
    bachelor_payments = pd.DataFrame({col_program:bachelor_payments.index, 'values':bachelor_payments.values})
    df_bachelor_dashboard[col_payments] = insert_values(df_bachelor_dashboard, bachelor_payments, col_program, col_payments)

    bachelor_enrollments = df_bachelor_enr.groupby(bachelor_col_enrollments)[bachelor_col_enrollments].count()
    bachelor_enrollments = pd.DataFrame({col_program:bachelor_enrollments.index, 'values':bachelor_enrollments.values})
    df_bachelor_dashboard[col_enrollments] = insert_values(df_bachelor_dashboard, bachelor_enrollments, col_program, col_enrollments)


    df_main_dashboard = pd.DataFrame(columns=df_master_dashboard.columns)
    df_main_dashboard.loc[len(df_main_dashboard)] = {col_program: main_studyonline, col_leads: main_leads}
    df = pd.concat([df_main_dashboard, df_master_dashboard, df_bachelor_dashboard], ignore_index=True, sort=False).drop(columns=['program_bitrix'])


    # считаем второстепенные столбцы
    df[col_leads_total]                            = df[col_leads_partners] + df[col_leads]
    df[col_conversion_leads_to_contracts]          = df[col_contracts] / df[col_leads_total]
    df[col_needed_applications]              = round(df[col_plan]/ NEEDED_APPLICATIONS_RATIO)
    df[col_conversion_applications_to_contracts]   = df[col_contracts] / df[col_applications]
    df[col_conversion_contracts_to_payments]       = df[col_payments]  / df[col_contracts]
    df[col_conversion_contracts_to_enrollments]    = df[col_payments]  / df[col_contracts]
    df[col_payments_div_plan]                      = df[col_payments]  / df[col_plan]
    df[col_income_1year     ] = df['price'] * df[col_payments] / 1000
    # df[col_income_all       ] = df['price'] * df[col_payments] * (2 if df['level'] == 'master' else 4)
    # df[col_income_1year_hse ] = df[col_income_1year] * df['income_percent'] / 100
    # df[col_income_all_hse   ] = df[col_income_all] * df['income_percent'] / 100

    df.replace(np.inf, 0, inplace=True)
    df.fillna(0, inplace=True)

    return df