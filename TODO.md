# 2026
## Текущая разработка
[ ] Бюджет и коммерция разделить в заявлениях
[ ] Сделать бины для осреднения по числу заявлений и договоров
[ ] Документировать число оплат и контрактов по дням
[ ] Количество регистраций за 1/2 недели
Поля нужные считываются, пока лимит в 10 запросов, чтобы не грузить систему.
Что надо сделать:
1) Уники отдельно по магам, отдельно по бакам, отдельно сумма, отдельно сумма без психологий
2) Сверить выгрузку с выгрузкой из АСАВ и Битрикс (старый расчет)
3) Сделать поле понедельной выгрузки начиная с разных дат (1 апреля, 20 июня)
6) Замедлить запросы, чтобы не нагружать сервер + спросить у Александра заметно ли замедление
8) Пройтись по всем TODO
9) Изучить сколько пропусков в ufDealEducationProgram и можно ли их заполнить

## Features
[ ] deal stage histogram
[ ] success of managers (conversions? leads?)
[ ] get data from Bitrix by API
[ ] append new programs and convert to csv
[x] read early invitation from Katya
[ ] group by manager
[x] add comparison to 2025 year - asav_file_2025, bachelor_file_2025, bitrix_file_2025?
[x] gender for each program (based on applications?)
[x] age for each program (based on applications?)
[x] archive data (leads, applications, contracts) for each program
--- delta for 1 weeks
[ ] add yaml for texts https://habr.com/ru/articles/1035714/
[x] add unique client's (applications column) for 2023-2025 years
[x] dynamics graphics
[x] comparison with prev. year
[x] sparkline for all data
[x] exclude offline finance master
[x] prev year for bachelors
[x] bar for each program of gosuslugi

## Technical debt
[ ] fill n/a in ufDealEducationProgram
[ ] refactor old years using AI and compare it with tests with different dates
[ ] delete unused files
[x] change bachelor of design fee to 420 and ПРВИС to 500
[x] choose "ЭКАНБАК" vs. "БАКЭКАН" and put it into database
[x] test after leads 2023 upload added
[ ] notation of csv & xlsx files
[ ] aliases for program names (bitrix, aispk)
[ ] secrets and history filest - to secret folder?
[ ] old files to bd? pyspark?

# Backlog
## Features
[ ] leads by channel
[ ] labels for bac_contracts check
[ ] check unique applications for bachelors
[ ] check unique applications for masters (filter online programs)
[x] trends from monday

## Technical debt
[ ] filter asav to online programs only
[ ] check col_conversion_contracts_to_enrollments calculations
[ ] deltas has mistakes
[ ] move templates and data to private account
[ ] check all TODO marks
[ ] solve all trunk warnings
[ ] solve all python warnings