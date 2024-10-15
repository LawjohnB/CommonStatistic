import sqlite3
from datetime import datetime

current_week = datetime.isocalendar(datetime.now())[1]

kte_db_create = '''
"id" integer not null primary key unique,
"exp_number" text,
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

def exps_update(rows, sheet_name):
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()

    if sheet_name == 'КТЭ':

    # создание таблицы
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS Week_{current_week}_{sheet_name}
            ({kte_db_create})
            ''')
        sqlite_connection.commit()


    if sqlite_connection:
        print('ok')
        sqlite_connection.close()




exps_update(1, 'КТЭ')

exit()


# # работа с файлом excel
# exps = xlrd.open_workbook(xlsx_file, on_demand = True)
# kte = exps.sheet_by_name('КТЭ')
# rows = (kte.row_values(x) for x in range(1, kte.nrows))
# # если нет номера эксп - прерываем (значит строка пустая)
# for row in rows:
#     if not row[0]:
#         break
#     exp_num = row[0]
#     exp_date_in = xls_date_to_ordinal(row[1])
#     init_organ = row[2]
#     init_ter = row[3]
#     init_fio = row[4]
#     mat_number = row[5]
#     uk_state = row[6]
#     fabula = row[7]
#     exp_fio = row[8]
#     objects_info = row[9]
#     objs_first_count = row[10]
#     objs_first_mobile = row[11]
#     objs_first_digital = row[12]
#     exp_type = row[13]
#     exp_status = row[14]
#     exp_end_date = 0
#     if exp_status != 'В производстве':
#         exp_end_date = xls_date_to_ordinal(row[15])
#     exp_result = row[16]
#     exp_days_count = str(row[17]).strip('() ').split('.')[0]
#     exp_days_count = int(exp_days_count)
#     # 18 - лиц установлено всегда 0, пропуск
#     facts_est = int(row[19]) if row[19] else 0
#     exp_vyvod = row[20]
#     objs_finish_count = int(row[21]) if row[21] else 0

#     # занесение данных в БД
#     try:    
#         sqlite_insert_with_param = f'''INSERT INTO {db_table_name}
#               (exp_number, exp_in_date, initiator_organ,
#               initiator_territory, initiator_fio,
#               mat_number, UK_state, fabula, exps_fio,
#               objects_info, objs_first_count, objs_first_mobile,
#               objs_first_digital, exp_type, exp_status, exp_end_date,
#               exp_result, exp_days_count, facts_est, exp_vyvod, 
#               objs_count_finish, server, computer_stat, computer_mobile, HDD, flash, CompactDisk,
#               AudioTEch, OtherComp, PaperDocs, MobilePhone, SIMcard, VideoRecorder,
#               PhotoVideoTech, Videofiles, DigitalPhotos, MailserverDatabase, EmailLetter, TabletPC,
#               UDvolume, AirbagControlUnit, GPStrack, FitnessBracelet, Router, EmailBox,
#               CloudServer, Database, Systemboard)
#               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 
#                       ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,  
#                       ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

# # кортеж основных параметров
#         data_tuple_main = (exp_num, exp_date_in, init_organ, init_ter, init_fio,
#         mat_number, uk_state, fabula, exp_fio, objects_info, objs_first_count, 
#         objs_first_mobile, objs_first_digital, exp_type, exp_status, exp_end_date,
#         exp_result, exp_days_count, facts_est, exp_vyvod, objs_finish_count)

# # кортеж количества объектов
#         data_obj_tuple = tuple(row[22:49])
#         data_tuple = data_tuple_main + data_obj_tuple
#         cursor.execute(sqlite_insert_with_param, data_tuple)
#     except sqlite3.Error as error:
#         log_data(str(error), 'sql_errors_log.txt')
#         if sqlite_connection:
#             sqlite_connection.close()
#         return 1
# sqlite_connection.commit()
# if sqlite_connection:
#     sqlite_connection.close()