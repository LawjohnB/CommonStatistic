import os
import glob
import xlrd
import sqlite3
from datetime import datetime

## Запросы к БД
# Создание таблицы КТЭ экспертиз
kte_db_create_exp_exp = '''
"id" integer not null primary key unique,
"exp_number" text,
"difficult" text,
"exp_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,
"UK_state" text,
"fabula" text,
"exps_fio" text,
"objects_info" text,
"objs_first_count" integer,
"objs_first_mobile" integer,
"objs_first_digital" integer,
"exp_type" text,
"exp_status" text,
"exp_end_date" text,
"exp_result" text,
"exp_days_count" integer,
"facts_est" integer,
"exp_vyvod" text,
"objs_count_finish" integer,
"server" integer,
"computer_stat" integer,
"computer_mobile" integer,
"HDD" integer,
"flash" integer,
"CompactDisk" integer,
"AudioTech" integer,
"OtherComp" integer,
"PaperDocs" integer,
"MobilePhone" integer,
"SIMcard" integer,
"VideoRecorder" integer,
"PhotoVideoTech" integer,
"Videofiles" integer,
"DigitalPhotos" integer,
"MailserverDatabase" integer,
"EmailLetter" integer,
"TabletPC" integer,
"UDvolume" integer,
"AirbagControlUnit" integer,
"GPStrack" integer,
"FitnessBracelet" integer,
"Router" integer,
"EmailBox" integer,
"CloudServer" integer,
"Database" integer,
"Systemboard" integer
'''


# номер текущей недели
current_week = datetime.isocalendar(datetime.now())[1]

# обработка даты из формата Excel 
def xls_date_to_ordinal(xls_date):
    xls_date = xlrd.xldate_as_tuple(xls_date, 0)
    xls_date = datetime(*xls_date).date()
    xls_date = datetime.toordinal(xls_date)
    return xls_date

# составление списка всех файлов 
files = glob.glob('./Журналы регионов/*/*.xls*')

###############################################################
## ФУНКЦИИ

# Функция проверки наличия правильного количества файлов нужного вида
def check_files_count(files):
    if len(files) != 5:
        print(f'Количество файлов {files} некорректно!')

## Функции работы с базой данных

# Функция создания таблицы в соответствии с номером недели и названием листа
def create_table_in_db(table_name):
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    cursor.execute(f'''
        CREATE TABLE IF NOT EXISTS {table_name}
        ({kte_db_create_exp_exp})
        ''')
    sqlite_connection.commit()
    if sqlite_connection:
        sqlite_connection.close()

def insert_kte_exp(table_name, rows):
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        difficult = row[0].split(',')[-1].strip()
        exp_date_in = xls_date_to_ordinal(row[1])
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        fabula = row[7]
        exp_fio = row[8]
        objects_info = row[9]
        objs_first_count = row[10]
        objs_first_mobile = row[11]
        objs_first_digital = row[12]
        exp_type = row[13]
        exp_status = row[14]
        exp_end_date = 0
        if exp_status != 'В производстве':
            exp_end_date = xls_date_to_ordinal(row[15])
        exp_result = row[16]
        exp_days_count = str(row[17]).strip('() ').split('.')[0]
        exp_days_count = int(exp_days_count)
        # 18 - лиц установлено всегда 0, пропуск
        facts_est = int(row[19]) if row[19] else 0
        exp_vyvod = row[20]
        objs_finish_count = int(row[21]) if isinstance(row[21], (int, float)) else 0
        if not isinstance(row[21], (int, float)):
            print(row[21])

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
              (exp_number, difficult, exp_in_date, initiator_organ,
              initiator_territory, initiator_fio,
              mat_number, UK_state, fabula, exps_fio,
              objects_info, objs_first_count, objs_first_mobile,
              objs_first_digital, exp_type, exp_status, exp_end_date,
              exp_result, exp_days_count, facts_est, exp_vyvod, 
              objs_count_finish, server, computer_stat, computer_mobile, HDD, flash, CompactDisk,
              AudioTEch, OtherComp, PaperDocs, MobilePhone, SIMcard, VideoRecorder,
              PhotoVideoTech, Videofiles, DigitalPhotos, MailserverDatabase, EmailLetter, TabletPC,
              UDvolume, AirbagControlUnit, GPStrack, FitnessBracelet, Router, EmailBox,
              CloudServer, Database, Systemboard)
              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                      ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,  
                      ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            # кортеж основных параметров
            data_tuple_main = (exp_num, difficult, exp_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, fabula, exp_fio, objects_info, objs_first_count, 
            objs_first_mobile, objs_first_digital, exp_type, exp_status, exp_end_date,
            exp_result, exp_days_count, facts_est, exp_vyvod, objs_finish_count)

            # кортеж количества объектов
            data_obj_tuple = tuple(row[22:49])
            data_tuple = data_tuple_main + data_obj_tuple
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            print(error)
            # log_data(str(error), 'sql_errors_log.txt')
        sqlite_connection.commit()
    if sqlite_connection:
        sqlite_connection.close()



## Конец функций
################################################################

# обход каталогов для построения всех файлов Excel (файлы регионов)
xlsx_files = glob.glob('./Журналы регионов/*/*.xls*')

# распределение файлов Excel по типам
xlsx_exps = [file for file in xlsx_files if 'Журнал экспертиз' in file]
xlsx_issls = [file for file in xlsx_files if 'Журнал исследований' in file]
xlsx_sipd =[file for file in xlsx_files if 'Журнал следственных действий' in file]
xlsx_consults = [file for file in xlsx_files if 'Журнал консультаций' in file]

# контрольная проверка наличия всех файлов (по 5 от каждого региона)
check_files_count(xlsx_exps)
check_files_count(xlsx_issls)
check_files_count(xlsx_sipd)
check_files_count(xlsx_consults)

# возможные листы
allowed_sheets = ('НАЛОГ', 'ИАЭ', 'ЛИНГВ', 'КТЭ', 'ФВТЭ', 'ФАЭ', 'ФОНО', 'БУХГ', 'ОЦЕН', 'ОИТИ', 'СМЭ')
allowed_sheets += ('СЭИ', 'ОКТИ', 'ОФиЛИ', 'ОСМИ')
allowed_sheets += ('Бухгалтерская', 'Видеотехническая', 'Инф.-аналитич.', 'Компьютерная', 'Лингвистическая', 'Судебно-медицинская', 'Фоноскопическая')
renamed_sheets = {'Бухгалтерская': 'БУХГ', 'Видеотехническая': 'ФВТЭ', 'Инф.-аналитич.': 'ИАЭ', 'Компьютерная': 'КТЭ', 'Лингвистическая': 'ЛИНГВ', 'Судебно-медицинская': 'СМЭ', 'Фоноскопическая': 'ФОНО'}

## Обход файлов Excel
# проход по каждому файлу экспертиз
for excel_file in xlsx_exps:
    exps = xlrd.open_workbook(excel_file, on_demand = True)
    for sheet_name in exps.sheet_names():
        if sheet_name not in allowed_sheets:
            backslash_char = "\\"
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пропущен лист "{sheet_name}"')
            continue
        current_sheet = exps.sheet_by_name(sheet_name)
        rows = [current_sheet.row_values(x) for x in range(1, current_sheet.nrows)]
        if not rows:
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пустой лист "{sheet_name}"')
            continue
        else:
            if sheet_name in renamed_sheets:
                sheet_name = renamed_sheets[sheet_name]
            table_name = f'Week_{current_week}_{sheet_name}_Exps'
            create_table_in_db(table_name)
            if sheet_name == 'КТЭ':
                insert_kte_exp(table_name, rows)
            # insert_data_in_db(sheet_name, table_name, rows)
    exps.release_resources()
    del exps

# проход по каждому файлу исследований
for excel_file in xlsx_issls:
    issls = xlrd.open_workbook(excel_file, on_demand = True)
    for sheet_name in issls.sheet_names():
        if sheet_name not in allowed_sheets:
            backslash_char = "\\"
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пропущен лист "{sheet_name}"')
            continue
        current_sheet = issls.sheet_by_name(sheet_name)
        rows = [current_sheet.row_values(x) for x in range(1, current_sheet.nrows)]
        if not rows:
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пустой лист "{sheet_name}"')
            continue
        else:
            if sheet_name in renamed_sheets:
                sheet_name = renamed_sheets[sheet_name]
            table_name = f'Week_{current_week}_{sheet_name}_Issls'
            create_table_in_db(table_name)
            # insert_data_in_db(sheet_name, table_name, rows)
    issls.release_resources()
    del issls


# проход по каждому файлу консультаций
for excel_file in xlsx_consults:
    consults = xlrd.open_workbook(excel_file, on_demand = True)
    for sheet_name in consults.sheet_names():
        if sheet_name not in allowed_sheets:
            backslash_char = "\\"
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пропущен лист "{sheet_name}"')
            continue
        current_sheet = consults.sheet_by_name(sheet_name)
        rows = [current_sheet.row_values(x) for x in range(1, current_sheet.nrows)]
        if not rows:
            # print(f'В файле "{excel_file.split(backslash_char)[-1]}" пустой лист "{sheet_name}"')
            continue
        else:
            table_name = f'Week_{current_week}_{sheet_name}_consults'
            create_table_in_db(table_name)
            # insert_data_in_db(sheet_name, table_name, rows)
    consults.release_resources()
    del consults