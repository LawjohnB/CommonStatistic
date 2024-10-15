import glob
import xlrd
from datetime import datetime
from create_tables_in_db import create_all_tables
import insert_data_to_db

backslash_char = "\\"

# номер текущей недели
current_week = datetime.isocalendar(datetime.now())[1]

# составление списка всех файлов 
files = glob.glob('./Журналы регионов/*.xls*')

# обход каталогов для построения всех файлов Excel (файлы регионов)
xlsx_files = glob.glob('./Журналы регионов/*.xls*')

# распределение файлов Excel по типам
xlsx_exps = sorted([file for file in xlsx_files if 'Журнал экспертиз' in file], key=lambda x: 'Ростов' not in x)
xlsx_issls = sorted([file for file in xlsx_files if 'Журнал исследований' in file], key=lambda x: 'Ростов' not in x)
xlsx_sipd = sorted([file for file in xlsx_files if 'Журнал следственных действий' in file], key=lambda x: 'Ростов' not in x)
xlsx_consults = sorted([file for file in xlsx_files if 'Журнал консультаций' in file], key=lambda x: 'Ростов' not in x)

# Функция проверки наличия правильного количества файлов нужного вида
def check_files_count(files):
    if len(files) != 5:
        print(f'Количество файлов {files} некорректно!')

# контрольная проверка наличия всех файлов (по 5 каждого вида)
check_files_count(xlsx_exps)
check_files_count(xlsx_issls)
check_files_count(xlsx_sipd)
check_files_count(xlsx_consults)

# Возможные листы
allowed_sheets = ('НАЛОГ', 'ИАЭ', 'ЛИНГВ', 'КТЭ', 'ФАЭ', 'ФОНО', 'БУХГ', 'ОЦЕН', 'ОИТИ', 'СМЭ')
allowed_sheets += ('СЭИ', 'ОКТИ', 'ОФиЛИ', 'ОСМИ')
allowed_sheets += ('Бухгалтерская', 'Инф.-аналитич.', 'Компьютерная',\
                   'Лингвистическая', 'Судебно-медицинская', 'Фоноскопическая')

# Исправление косяков в названиях Краснодара 
renamed_sheets = {'Бухгалтерская': 'БУХГ', 'Инф.-аналитич.': 'ИАЭ', \
                  'Компьютерная': 'КТЭ', 'Лингвистическая': 'ЛИНГВ', 'Судебно-медицинская': 'СМЭ', \
                  'Фоноскопическая': 'ФОНО'}

## Создание пустых таблиц текущей нелели для внесения данных 
create_all_tables(current_week)
print(f'Таблицы {current_week} недели успешно созданы')


## Обход файлов Excel
# Проход по каждому файлам экспертиз
for excel_file in xlsx_exps:
    exps = xlrd.open_workbook(excel_file, on_demand = True)
    for sheet_name in exps.sheet_names():
        if sheet_name not in allowed_sheets:   
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пропущен лист "{sheet_name}"')
            continue
            # пропуск СМЭ не из Ростова
        if sheet_name == 'СМЭ' and not 'Ростов' in excel_file:
            print(f'Пропущен лист СМЭ в файле {excel_file}')
            continue
        current_sheet = exps.sheet_by_name(sheet_name)
        rows = [current_sheet.row_values(x) for x in range(1, current_sheet.nrows)]
        if not rows:
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пустой лист "{sheet_name}"')
            continue
        if current_sheet in renamed_sheets:
            sheet_name = renamed_sheets[sheet_name]
        table_name = f'Week_{current_week}_{sheet_name}_Exps'
        res = insert_data_to_db.table_query_exps[sheet_name](table_name, rows)
        if res:
            print(f'Заполнение таблицы {table_name} из файла {excel_file} завершено с ошибками')
    exps.release_resources()
    del exps

# Проход по каждому файлам исследований
for excel_file in xlsx_issls:
    issls = xlrd.open_workbook(excel_file, on_demand = True)
    for sheet_name in issls.sheet_names():
        if sheet_name not in allowed_sheets:
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пропущен лист "{sheet_name}"')
            continue
        current_sheet = issls.sheet_by_name(sheet_name)
        rows = [current_sheet.row_values(x) for x in range(1, current_sheet.nrows)]
        if not rows:
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пустой лист "{sheet_name}"')
            continue
        if sheet_name in renamed_sheets:
            sheet_name = renamed_sheets[sheet_name]
        table_name = f'Week_{current_week}_{sheet_name}_Issls'
        insert_data_to_db.table_query_issls[sheet_name](table_name, rows)
    issls.release_resources()
    del issls

# Проход по каждому файлам СиПД
for excel_file in xlsx_sipd:
    sipd = xlrd.open_workbook(excel_file, on_demand = True)
    for sheet_name in sipd.sheet_names():
        if sheet_name not in allowed_sheets:
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пропущен лист "{sheet_name}"')
            continue
        current_sheet = sipd.sheet_by_name(sheet_name)
        rows = [current_sheet.row_values(x) for x in range(1, current_sheet.nrows)]
        if not rows:
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пустой лист "{sheet_name}"')
            continue
        if sheet_name in renamed_sheets:
            sheet_name = renamed_sheets[sheet_name]
        table_name = f'Week_{current_week}_{sheet_name}_SiPD'
        insert_data_to_db.table_query_sipd[sheet_name](table_name, rows)
    sipd.release_resources()
    del sipd


# Проход по каждому файлу консультаций и командировок
for excel_file in xlsx_consults:
    # консультации
    xlsx_consults = xlrd.open_workbook(excel_file, on_demand = True)
    current_sheet = xlsx_consults.sheet_by_name('Иная_деятельность')
    rows = [current_sheet.row_values(x) for x in range(1, current_sheet.nrows)]
    if not rows:
        print(f'В файле "{excel_file.split(backslash_char)[-1]}" пустой лист консультаций')
        continue
    table_name = f'Week_{current_week}_Consults'
    insert_data_to_db.consults(table_name, rows)
    # командировки
    xlsx_trips = xlrd.open_workbook(excel_file, on_demand = True)
    current_sheet = xlsx_trips.sheet_by_name('Командировки')
    rows = [current_sheet.row_values(x) for x in range(1, current_sheet.nrows)]
    if not rows:
        print(f'В файле "{excel_file.split(backslash_char)[-1]}" пустой лист командировок')
        continue
    table_name = f'Week_{current_week}_Trips'
    insert_data_to_db.trips(table_name, rows)
    xlsx_consults.release_resources()
    del xlsx_consults
