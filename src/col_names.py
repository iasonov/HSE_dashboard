master_col_programs = "Конкурс в магистратуру" #названия вида "Прикладная социальная психология"
master_col_contracts = "Договор на обучение" # не пусто
master_col_payments = "Оплата первого периода" # "Оплачено"
master_col_enrollments = "Приказ о зачислении" # не пусто
master_col_reg_number = "Код в ИС-ПРО" # для удаления дублей
master_col_campus = 'Кампус конкурса' # для нужд фильтрации программы по Финансам
master_col_program_specialization = 'Магистерская специализация' # для нужд фильтрации программ c офлайн-треками



master_foreign_col_programs_1 = "Программа 1 приоритета"
master_foreign_col_programs_2 = "Программа 2 приоритета"
master_foreign_col_faculty_1 = "Факультет первой программы"
master_foreign_col_faculty_2 = "Факультет второй программы"
master_foreign_col_program_final = "Рекомендован на конкурс"
master_foreign_col_faculty_final = "Факультет конкурса, на который рекомендован"
master_foreign_col_contracts = "Договор подписан"
master_foreign_col_payments = "Договор оплачен"
master_foreign_col_enrollments = None #"Приказ о зачислении" # пока не ясно, откуда брать этот столбец

bitrix_col_date = "Дата создания"

bachelor_col_date = "Дата"
bachelor_col_date_contract = "Дата заключения"
bachelor_col_programs = "Конкурсная группа" # названия вида "Глобальные цифровые коммуникации (Медиа) - онлайн (О К)"
bachelor_col_programs_contracts = "Образовательная программа"
bachelor_col_payments = "Статус оплаты" # "Оплачен" или "Оплачен по квитанциям"
bachelor_col_enrollments = bachelor_col_programs_contracts # "Конкурс"
# bachelor_col_gosuslugi = "Источник данных" # ЕПГУ - госуслуги
bachelor_col_reg_number = "Регистрационный номер" # для подсчета уников

bachelor_dict = {
    'Глобальные цифровые коммуникации (онлайн)'                        :'Глобальные цифровые коммуникации - онлайн (О К)'        ,
    'Глобальные цифровые коммуникации (онлайн) '                       :'Глобальные цифровые коммуникации (Медиа) - онлайн (О К)',
    'Компьютерные науки и анализ данных (онлайн)'                      :'Компьютерные науки и анализ данных - онлайн (О К)'      ,
    'Экономический анализ'                                             :'Экономический анализ - онлайн (О К)'                    ,
    'Дизайн  (онлайн)'                                                 :'Дизайн - онлайн (О К)'                                  ,
    'Управление бизнесом'                                              :'Управление бизнесом - онлайн (О К)',
    'Программные системы и автоматизация процессов разработки (онлайн)':'Программные системы и автоматизация процессов разработки - онлайн (О К)'
        }

# не название столбца, а название для лидов с разводящего ленда
main_studyonline = "Общий лендинг"

# основные параметры
col_program = 'program'
col_program_bitrix = "program_bitrix" # названия вида "ПСИХТЕР. Психоанализ и психоаналитическая психотерапия / Москва / 370401 Психология / факультет социальных наук / Магистратура"
col_programs_names = "Образовательная программа"
col_plan_rus = 'plan_rus'
col_plan_foreign = 'plan_foreign'
col_leads = "Кол-во заявок (online.hse.ru) с 1.10"
col_leads_after_april = """Кол-во заявок (online.hse.ru) с 1.04"""
col_leads_by_week = """Кол-во заявок (online.hse.ru) с 1.10 по неделям"""
col_leads_after_april_prev = """Прошлогоднее Кол-во заявок (studyonline) с 1.10"""
col_applications = """Регистрации в ЛК (РФ все приоритеты)"""
col_applications_prev = """Прошлогодние Регистрации в ЛК (РФ все приоритеты)"""
col_applications_by_week = """Регистрации в ЛК (РФ все приоритеты) по неделям"""
# col_applications_gosuslugi = "Регистрации в ЛК из Госуслуг"
col_contracts = """Договоры (ПК) РФ"""
col_contracts_prev =  """Прошлогодние Договоры (ПК) РФ"""
col_contracts_by_week = """Договоры (ПК) РФ по неделям"""
col_payments = """Оплаты (ПК)"""
col_applications_foreign = """Регистрации в ЛК иностранцы"""
col_contracts_foreign = """Договоры иностранцы"""
col_payments_foreign = """Оплаты иностранцы"""
col_enrollments_foreign = """Зачисленные иностранцы"""
col_enrollments = "Зачисленные (ПК)"
col_gender_asav = "Пол"
col_birthday = "Дата рождения"
col_male = "male"
col_female	= "female"
col_ages = "ages"
col_ages_mean = "ages_mean"
col_early_invitation = "Раннее приглашение" # название столбца в дашборде - число заявок по раннему приглашению в АСАВ


# дополнительные и расчетные параметры
col_leads_total = """Общее кол-во заявок ВСЕГО (портал+РК)"""
col_leads_partners = """Общее кол-во заявок  c  hse.ru по кнопкам/партнерских стр"""
col_conversion_leads_to_contracts = """Конверсия заявка -> договор (без учета заявок с общего ленда)"""
col_needed_applications = """Необходимо регистраций в ЛК для обеспечения набора (45% поступили от регистрации в АСАВ)"""
col_conversion_applications_to_contracts = """Конверсия ЛК -> договор"""
col_conversion_contracts_to_payments = """Конверсия договор -> оплата"""
col_conversion_contracts_to_enrollments = """Конверсия договор -> зачисление"""
col_payments_div_plan_rus = """Выполнение плана % РФ"""
col_payments_div_plan_foreign = """Выполнение плана % иностранцы"""

col_income_1year     = "Выручка за 1 год, млн.руб."
col_income_all       = "Выручка за весь период обучения"
col_income_1year_hse = "Выручка за 1 год после отчислений партнерам"
col_income_all_hse   = "Выручка за весь период обучения после отчислений партнерам"
col_leads_delta = """Общее кол-во заявок (studyonline) за 1\\2 недели"""
col_applications_delta = """Кол-во регистраций в ЛК абитуриента за 1\\2 недели"""

#asav
col_id_asav = "Рег. номер"
#ais pk have no id:
col_id_bachelor = "Абитуриент"

#dashboard unused
col_gender = "gender"
col_ages_bars = "ages_bars"

