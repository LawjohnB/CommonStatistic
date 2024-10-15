import sqlite3
from datetime import datetime

# Функция создания таблицы в соответствии с номером недели и названием листа
def create_table_in_db(table_name, query):
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    cursor.execute(f'''
        DROP TABLE IF EXISTS {table_name}''')
    sqlite_connection.commit()
    cursor.execute(f'''
        CREATE TABLE IF NOT EXISTS {table_name}
        ({query})
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

