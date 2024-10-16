import sqlite3
import xlrd
from datetime import datetime

sqlite_connection = sqlite3.connect('Все журналы.db')
cursor = sqlite_connection.cursor()

# перевод даты Excel в Unix
def xls_date_to_ordinal(xls_date):
    xls_date = xlrd.xldate_as_tuple(xls_date, 0)
    xls_date = datetime(*xls_date).date()
    xls_date = datetime.toordinal(xls_date)
    return xls_date

# перевод даты из строки в Unix
def str_date_to_ordinal(str_date):
    try:
        str_date = str_date.replace(',', '.')
        str_date = datetime.strptime(str_date, '%d.%m.%Y')
        str_date = datetime.toordinal(str_date)
        return str_date
    except:
        return 0

# Фукнция обработки ошибок в файл
def log_data(data, file_name):
     with open(file_name, 'a', encoding='UTF-8') as file:
          file.write(str(datetime.now()))
          file.write('\n')
          file.write(data)
          file.write('\n\n')

# Добавление данных в таблицы экспертиз
def nalog_exps(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простая'
            print(table_name, row[0], 'не указана сложность!')
        try:
            exp_date_in = xls_date_to_ordinal(row[1])
        except:
            exp_date_in = str_date_to_ordinal(row[1])
        if not exp_date_in:
            print(f'Некорректная дата поступления налоговой экспертизы {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        exp_type = row[9]
        exp_status = row[10].capitalize()
        exp_end_date = 0
        if exp_status.capitalize() != 'В производстве':
            try:
                exp_end_date = xls_date_to_ordinal(row[11])
            except:
                exp_end_date = str_date_to_ordinal(row[11])
            if not exp_end_date:
                print(f'Возможно некорректная дата окончания налоговой экспертизы {row[0]} ({row[11]})')
        exp_result = row[12]
        exp_days_count = str(row[13]).strip('() ').split('.')[0]
        try:
            exp_days_count = int(exp_days_count)
        except:
            exp_days_count = datetime.toordinal(datetime.now()) - exp_date_in
        persons_est = int(row[14]) if row[14] else 0
        facts_est = int(row[15]) if row[15] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (exp_number, difficult, exp_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_count, exp_type, exp_status,
                  exp_end_date, exp_result, exp_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (exp_num, difficult, exp_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_count, exp_type, exp_status, exp_end_date,
            exp_result, exp_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def ia_exps(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простая'
            print(table_name, row[0], 'не указана сложность!')
        try:
            exp_date_in = xls_date_to_ordinal(row[1])
        except:
            exp_date_in = str_date_to_ordinal(row[1])
        if not exp_date_in:
            print(f'Некорректная дата поступления информационно-аналитической экспертизы {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        exp_type = row[9]
        exp_status = row[10].capitalize()
        exp_end_date = 0
        if exp_status.capitalize() != 'В производстве':
            try:
                exp_end_date = xls_date_to_ordinal(row[11])
            except:
                exp_end_date = str_date_to_ordinal(row[11])
            if not exp_end_date:
                print(f'Возможно некорректная дата окончания информационно-аналитической экспертизы {row[0]} ({row[11]})')
        exp_result = row[12]
        exp_days_count = str(row[13]).strip('() ').split('.')[0]
        try:
            exp_days_count = int(exp_days_count)
        except:
            exp_days_count = datetime.toordinal(datetime.now()) - exp_date_in
        persons_est = int(row[14]) if row[14] else 0
        facts_est = int(row[15]) if row[15] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (exp_number, difficult, exp_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_count, exp_type, exp_status,
                  exp_end_date, exp_result, exp_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (exp_num, difficult, exp_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_count, exp_type, exp_status, exp_end_date,
            exp_result, exp_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors


def lingv_exps(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простая'
            print(table_name, row[0], 'не указана сложность!')
        try:
            exp_date_in = xls_date_to_ordinal(row[1])
        except:
            exp_date_in = str_date_to_ordinal(row[1])
        if not exp_date_in:
            print(f'Некорректная дата поступления лингвистической экспертизы {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_info = row[8]
        objs_count = row[9]
        exp_type = row[10]
        exp_status = row[11].capitalize()
        exp_end_date = 0
        if exp_status.capitalize() != 'В производстве':
            try:
                exp_end_date = xls_date_to_ordinal(row[12])
            except:
                exp_end_date = str_date_to_ordinal(row[12])
            if not exp_end_date:
                print(f'Возможно некорректная дата окончания лингвистической экспертизы {row[0]} ({row[12]})')
        exp_result = row[13]
        exp_days_count = str(row[14]).strip('() ').split('.')[0]
        try:
            exp_days_count = int(exp_days_count)
        except:
            exp_days_count = datetime.toordinal(datetime.now()) - exp_date_in
        crime_persons_est = int(row[15]) if row[15] else 0
        facts_est = int(row[16]) if row[16] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (exp_number, difficult, exp_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_info, objs_count, exp_type, exp_status,
                  exp_end_date, exp_result, exp_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (exp_num, difficult, exp_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_info, objs_count, exp_type, exp_status, exp_end_date,
            exp_result, exp_days_count, crime_persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def kt_exps(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простая'
            print(table_name, row[0], 'не указана сложность!')
        try:
            exp_date_in = xls_date_to_ordinal(row[1])
        except:
            exp_date_in = str_date_to_ordinal(row[1])
        if not exp_date_in:
            print(f'Некорректная дата поступления компьютерно-технической экспертизы {row[0]} ({row[1]})')
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
        exp_status = row[14].capitalize()
        exp_end_date = 0
        if exp_status.capitalize() != 'В производстве':
            try:
                exp_end_date = xls_date_to_ordinal(row[15])
            except:
                exp_end_date = str_date_to_ordinal(row[15])
            if not exp_end_date:
                print(f'Возможно некорректная дата окончания информационно-аналитической экспертизы {row[0]} ({row[15]})')
        exp_result = row[16]
        exp_days_count = str(row[17]).strip('() ').split('.')[0]
        exp_days_count = int(exp_days_count)
        # 18 - лиц установлено всегда 0, пропуск
        facts_est = int(row[19]) if row[19] else 0
        exp_vyvod = row[20]
        try: 
            objs_finish_count = int(row[21]) if row[21] else 0
        except:
            objs_finish_count = 0

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
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def fa_exps(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простая'
            print(table_name, row[0], 'не указана сложность!')
        try:
            exp_date_in = xls_date_to_ordinal(row[1])
        except:
            exp_date_in = str_date_to_ordinal(row[1])
        if not exp_date_in:
            print(f'Некорректная дата поступления финансово-аналитической экспертизы {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        exp_type = row[9]
        exp_status = row[10].capitalize()
        exp_end_date = 0
        if exp_status.capitalize() != 'В производстве':
            try:
                exp_end_date = xls_date_to_ordinal(row[11])
            except:
                exp_end_date = str_date_to_ordinal(row[11])
            if not exp_end_date:
                print(f'Возможно некорректная дата окончания финансово-аналитической экспертизы {row[0]} ({row[11]})')
        exp_result = row[12]
        exp_days_count = str(row[13]).strip('() ').split('.')[0]
        try:
            exp_days_count = int(exp_days_count)
        except:
            exp_days_count = datetime.toordinal(datetime.now()) - exp_date_in
        persons_est = int(row[14]) if row[14] else 0
        facts_est = int(row[15]) if row[15] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (exp_number, difficult, exp_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_count, exp_type, exp_status,
                  exp_end_date, exp_result, exp_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (exp_num, difficult, exp_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_count, exp_type, exp_status, exp_end_date,
            exp_result, exp_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def fono_exps(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простая'
            print(table_name, row[0], 'не указана сложность!')
        try:
            exp_date_in = xls_date_to_ordinal(row[1])
        except:
            exp_date_in = str_date_to_ordinal(row[1])
        if not exp_date_in:
            print(f'Некорректная дата поступления фоноскопической экспертизы {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        acoustic_exps_fio = row[7]
        lingv_exps_fio = row[8]
        objs_info = row[9]
        objs_count = row[10]
        exp_type = row[11]
        exp_status = row[12].capitalize()
        exp_end_date = 0
        if exp_status.capitalize() != 'В производстве':
            try:
                exp_end_date = xls_date_to_ordinal(row[13])
            except:
                exp_end_date = str_date_to_ordinal(row[13])
            if not exp_end_date:
                print(f'Возможно некорректная дата окончания фоноскопической экспертизы {row[0]} ({row[13]})')
        exp_result = row[14]
        verbatim_duration = row[15]
        idents_count = row[16]
        persons_est = row[17]
        exp_days_count = str(row[18]).strip('() ').split('.')[0]
        try:
            exp_days_count = int(exp_days_count)
        except:
            exp_days_count = datetime.toordinal(datetime.now()) - exp_date_in
        crime_persons_est = int(row[19]) if row[19] else 0
        facts_est = int(row[20]) if row[20] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (exp_number, difficult, exp_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, acoustic_exps_fio, lingv_exps_fio, objs_info,
                  objs_count, exp_type, exp_status, exp_end_date, exp_result, verbatim_duration,
                  idents_count, persons_est, exp_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (exp_num, difficult, exp_date_in, init_organ, init_ter, init_fio, mat_number, uk_state,
                        acoustic_exps_fio, lingv_exps_fio, objs_info, objs_count, exp_type,
                        exp_status, exp_end_date, exp_result, verbatim_duration, idents_count,
                        persons_est, exp_days_count, crime_persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors


def buh_exps(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простая'
            print(table_name, row[0], 'не указана сложность!')
        try:
            exp_date_in = xls_date_to_ordinal(row[1])
        except:
            exp_date_in = str_date_to_ordinal(row[1])
        if not exp_date_in:
            print(f'Некорректная дата поступления бухгалтерской экспертизы {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        exp_type = row[9]
        exp_status = row[10].capitalize()
        exp_end_date = 0
        if exp_status.capitalize() != 'В производстве':
            try:
                exp_end_date = xls_date_to_ordinal(row[11])
            except:
                exp_end_date = str_date_to_ordinal(row[11])
            if not exp_end_date:
                print(f'Возможно некорректная дата окончания бухгалтерской экспертизы {row[0]} ({row[11]})')
        exp_result = row[12]
        exp_days_count = str(row[13]).strip('() ').split('.')[0]
        try:
            exp_days_count = int(exp_days_count)
        except:
            exp_days_count = datetime.toordinal(datetime.now()) - exp_date_in
        persons_est = int(row[14]) if row[14] else 0
        facts_est = int(row[15]) if row[15] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (exp_number, difficult, exp_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_count, exp_type, exp_status,
                  exp_end_date, exp_result, exp_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (exp_num, difficult, exp_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_count, exp_type, exp_status, exp_end_date,
            exp_result, exp_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def ocen_exps(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простая'
            print(table_name, row[0], 'не указана сложность!')
        try:
            exp_date_in = xls_date_to_ordinal(row[1])
        except:
            exp_date_in = str_date_to_ordinal(row[1])
        if not exp_date_in:
            print(f'Некорректная дата поступления оценочной экспертизы {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        exp_type = row[9]
        exp_status = row[10].capitalize()
        exp_end_date = 0
        if exp_status.capitalize() != 'В производстве':
            try:
                exp_end_date = xls_date_to_ordinal(row[11])
            except:
                exp_end_date = str_date_to_ordinal(row[11])
            if not exp_end_date:
                print(f'Возможно некорректная дата окончания оценочной экспертизы {row[0]} ({row[11]})')
        exp_result = row[12]
        exp_days_count = str(row[13]).strip('() ').split('.')[0]
        try:
            exp_days_count = int(exp_days_count)
        except:
            exp_days_count = datetime.toordinal(datetime.now()) - exp_date_in
        persons_est = int(row[14]) if row[14] else 0
        facts_est = int(row[15]) if row[15] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (exp_number, difficult, exp_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_count, exp_type, exp_status,
                  exp_end_date, exp_result, exp_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (exp_num, difficult, exp_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_count, exp_type, exp_status, exp_end_date,
            exp_result, exp_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def oiti_exps(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простая'
            print(table_name, row[0], 'не указана сложность!')
        try:
            exp_date_in = xls_date_to_ordinal(row[1])
        except:
            exp_date_in = str_date_to_ordinal(row[1])
        if not exp_date_in:
            print(f'Некорректная дата поступления ОИТИ экспертизы {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        exp_type = row[9]
        exp_status = row[10].capitalize()
        exp_end_date = 0
        if exp_status.capitalize() != 'В производстве':
            try:
                exp_end_date = xls_date_to_ordinal(row[11])
            except:
                exp_end_date = str_date_to_ordinal(row[11])
            if not exp_end_date:
                print(f'Возможно некорректная дата окончания ОИТИ экспертизы {row[0]} ({row[11]})')
        exp_result = row[12]
        exp_days_count = str(row[13]).strip('() ').split('.')[0]
        try:
            exp_days_count = int(exp_days_count)
        except:
            exp_days_count = datetime.toordinal(datetime.now()) - exp_date_in
        persons_est = int(row[14]) if row[14] else 0
        facts_est = int(row[15]) if row[15] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (exp_number, difficult, exp_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_count, exp_type, exp_status,
                  exp_end_date, exp_result, exp_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (exp_num, difficult, exp_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_count, exp_type, exp_status, exp_end_date,
            exp_result, exp_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors


def sm_exps(table_name, rows):
    res_with_errors = 0
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        exp_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Сложная'
            print(table_name, row[0], 'не указана сложность!')
        try:
            exp_date_in = xls_date_to_ordinal(row[1])
        except:
            exp_date_in = str_date_to_ordinal(row[1])
        if not exp_date_in:
            print(f'Некорректная дата поступления СМ экспертизы {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        exp_type = row[9]
        exp_status = row[10].capitalize()
        exp_end_date = 0
        if exp_status.capitalize() != 'В производстве':
            try:
                exp_end_date = xls_date_to_ordinal(row[11])
            except:
                exp_end_date = str_date_to_ordinal(row[11])
            if not exp_end_date:
                print(f'Возможно некорректная дата окончания СМ экспертизы {row[0]} ({row[11]})')
        exp_result = row[12]
        objs_count = row[13]
        patient_fio = row[14]
        patient_status = row[15]
        med_docs = row[16]
        exp_days_count = str(row[17]).strip('() ').split('.')[0]
        try:
            exp_days_count = int(exp_days_count)
        except:
            exp_days_count = datetime.toordinal(datetime.now()) - exp_date_in
        crime_persons_est = int(row[18]) if row[18] else 0
        facts_est = int(row[19]) if row[19] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (exp_number, difficult, exp_in_date, initiator_organ, initiator_territory, initiator_fio,
                  mat_number, UK_state, exps_fio, objs_count, exp_type, exp_status, exp_end_date, exp_result,
                  objs_count, patient_fio, patient_status, med_docs, exp_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (exp_num, difficult, exp_date_in, init_organ, init_ter, init_fio, mat_number, uk_state,
                          exp_fio, objs_count, exp_type, exp_status, exp_end_date, exp_result, objs_count,
                          patient_fio, patient_status, med_docs, exp_days_count, crime_persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

# Добавление данных в таблицы исследований
def nalog_issls(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        issl_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простое'
            print(table_name, row[0], 'не указана сложность!')
        try:
            issl_date_in = xls_date_to_ordinal(row[1])
        except:
            issl_date_in = str_date_to_ordinal(row[1])
        if not issl_date_in:
            print(f'Некорректная дата поступления налогового исследования {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        issl_status = row[9].capitalize()
        issl_end_date = 0
        if issl_status.capitalize() != 'В производстве':
            try:
                issl_end_date = xls_date_to_ordinal(row[10])
            except:
                issl_end_date = str_date_to_ordinal(row[10])
            if not issl_end_date:
                print(f'Возможно некорректная дата окончания налогового исследования {row[0]} ({row[10]})')
        issl_result = row[11]
        issl_days_count = str(row[12]).strip('() ').split('.')[0]
        try:
            issl_days_count = int(issl_days_count)
        except:
            issl_days_count = datetime.toordinal(datetime.now()) - issl_date_in
        persons_est = int(row[13]) if row[13] else 0
        facts_est = int(row[14]) if row[14] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (issl_number, difficult, issl_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_count, issl_status,
                  issl_end_date, issl_result, issl_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (issl_num, difficult, issl_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_count, issl_status, issl_end_date,
            issl_result, issl_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def ia_issls(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        issl_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простое'
            print(table_name, row[0], 'не указана сложность!')
        try:
            issl_date_in = xls_date_to_ordinal(row[1])
        except:
            issl_date_in = str_date_to_ordinal(row[1])
        if not issl_date_in:
            print(f'Некорректная дата поступления информационно-аналитического исследования {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        issl_status = row[9].capitalize()
        issl_end_date = 0
        if issl_status.capitalize() != 'В производстве':
            try:
                issl_end_date = xls_date_to_ordinal(row[10])
            except:
                issl_end_date = str_date_to_ordinal(row[10])
            if not issl_end_date:
                print(f'Возможно некорректная дата окончания информационно-аналитического исследования {row[0]} ({row[10]})')
        issl_result = row[11]
        issl_days_count = str(row[12]).strip('() ').split('.')[0]
        try:
            issl_days_count = int(issl_days_count)
        except:
            issl_days_count = datetime.toordinal(datetime.now()) - issl_date_in
        persons_est = int(row[13]) if row[13] else 0
        facts_est = int(row[14]) if row[14] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (issl_number, difficult, issl_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_count, issl_status,
                  issl_end_date, issl_result, issl_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (issl_num, difficult, issl_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_count, issl_status, issl_end_date,
            issl_result, issl_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def lingv_issls(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        issl_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простое'
            print(table_name, row[0], 'не указана сложность!')
        try:
            issl_date_in = xls_date_to_ordinal(row[1])
        except:
            issl_date_in = str_date_to_ordinal(row[1])
        if not issl_date_in:
            print(f'Некорректная дата поступления лингвистического исследования {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_info = row[8]
        objs_count = row[9]
        issl_status = row[10].capitalize()
        issl_end_date = 0
        if issl_status.capitalize() != 'В производстве':
            try:
                issl_end_date = xls_date_to_ordinal(row[11])
            except:
                issl_end_date = str_date_to_ordinal(row[11])
            if not issl_end_date:
                print(f'Возможно некорректная дата окончания лингвистического исследования {row[0]} ({row[11]})')
        issl_result = row[12]
        issl_days_count = str(row[13]).strip('() ').split('.')[0]
        try:
            issl_days_count = int(issl_days_count)
        except:
            issl_days_count = datetime.toordinal(datetime.now()) - issl_date_in
        persons_est = int(row[14]) if row[14] else 0
        facts_est = int(row[15]) if row[15] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (issl_number, difficult, issl_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_info, objs_count, issl_status,
                  issl_end_date, issl_result, issl_days_count, persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (issl_num, difficult, issl_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_info, objs_count, issl_status, issl_end_date,
            issl_result, issl_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def kt_issls(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        issl_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простая'
            print(table_name, row[0], 'не указана сложность!')
        try:
            issl_date_in = xls_date_to_ordinal(row[1])
        except:
            issl_date_in = str_date_to_ordinal(row[1])
        if not issl_date_in:
            print(f'Некорректная дата поступления компьютерно-технической экспертизы {row[0]} ({row[1]})')
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
        issl_status = row[13].capitalize()
        issl_end_date = 0
        if issl_status.capitalize() != 'В производстве':
            try:
                issl_end_date = xls_date_to_ordinal(row[14])
            except:
                issl_end_date = str_date_to_ordinal(row[14])
            if not issl_end_date:
                print(f'Возможно некорректная дата окончания информационно-аналитической экспертизы {row[0]} ({row[14]})')
        issl_result = row[15]
        issl_days_count = str(row[16]).strip('() ').split('.')[0]
        issl_days_count = int(issl_days_count)
        # 17 - лиц установлено всегда 0, пропуск
        facts_est = int(row[18]) if row[18] else 0
        issl_vyvod = row[19]
        try: 
            objs_finish_count = int(row[20]) if row[20] else 0
        except:
            objs_finish_count = 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (issl_number, difficult, issl_in_date, initiator_organ,
                  initiator_territory, initiator_fio,
                  mat_number, UK_state, fabula, exp_fio,
                  objects_info, objs_first_count, objs_first_mobile,
                  objs_first_digital, issl_status, issl_end_date,
                  issl_result, issl_days_count, facts_est, issl_vyvod, 
                  objs_count_finish, server, computer_stat, computer_mobile, HDD, flash, CompactDisk,
                  AudioTEch, OtherComp, PaperDocs, MobilePhone, SIMcard, VideoRecorder,
                  PhotoVideoTech, Videofiles, DigitalPhotos, MailserverDatabase, EmailLetter, TabletPC,
                  UDvolume, AirbagControlUnit, GPStrack, FitnessBracelet, Router, EmailBox,
                  CloudServer, Database, Systemboard)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                          ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,  
                          ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

    # кортеж основных параметров
            data_tuple_main = (issl_num, difficult, issl_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, fabula, exp_fio, objects_info, objs_first_count, 
            objs_first_mobile, objs_first_digital, issl_status, issl_end_date,
            issl_result, issl_days_count, facts_est, issl_vyvod, objs_finish_count)

    # кортеж количества объектов
            data_obj_tuple = tuple(row[22:49])
            data_tuple = data_tuple_main + data_obj_tuple
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def fa_issls(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        issl_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простое'
            print(table_name, row[0], 'не указана сложность!')
        try:
            issl_date_in = xls_date_to_ordinal(row[1])
        except:
            issl_date_in = str_date_to_ordinal(row[1])
        if not issl_date_in:
            print(f'Некорректная дата поступления ФА исследования {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        issl_status = row[9].capitalize()
        issl_end_date = 0
        if issl_status.capitalize() != 'В производстве':
            try:
                issl_end_date = xls_date_to_ordinal(row[10])
            except:
                issl_end_date = str_date_to_ordinal(row[10])
            if not issl_end_date:
                print(f'Возможно некорректная дата окончания ФА исследования {row[0]} ({row[10]})')
        issl_result = row[11]
        issl_days_count = str(row[12]).strip('() ').split('.')[0]
        try:
            issl_days_count = int(issl_days_count)
        except:
            issl_days_count = datetime.toordinal(datetime.now()) - issl_date_in
        persons_est = int(row[13]) if row[13] else 0
        facts_est = int(row[14]) if row[14] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (issl_number, difficult, issl_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_count, issl_status,
                  issl_end_date, issl_result, issl_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (issl_num, difficult, issl_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_count, issl_status, issl_end_date,
            issl_result, issl_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def fono_issls(table_name, rows):
    # print('ФОНО исследования')
    pass

def buh_issls(table_name, rows):
    res_with_errors = 0
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        if not row[0]:
            break
        issl_num = row[0].split(',')[0].strip()
        try:       
            difficult = row[0].split(',')[1].strip().capitalize()
        except:
            difficult = 'Простое'
            print(table_name, row[0], 'не указана сложность!')
        try:
            issl_date_in = xls_date_to_ordinal(row[1])
        except:
            issl_date_in = str_date_to_ordinal(row[1])
        if not issl_date_in:
            print(f'Некорректная дата поступления бухгалтерского исследования {row[0]} ({row[1]})')
        init_organ = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exp_fio = row[7]
        objs_count = row[8]
        issl_status = row[9].capitalize()
        issl_end_date = 0
        if issl_status.capitalize() != 'В производстве':
            try:
                issl_end_date = xls_date_to_ordinal(row[10])
            except:
                issl_end_date = str_date_to_ordinal(row[10])
            if not issl_end_date:
                print(f'Возможно некорректная дата окончания бухгалтерского исследования {row[0]} ({row[10]})')
        issl_result = row[11]
        issl_days_count = str(row[12]).strip('() ').split('.')[0]
        try:
            issl_days_count = int(issl_days_count)
        except:
            issl_days_count = datetime.toordinal(datetime.now()) - issl_date_in
        persons_est = int(row[13]) if row[13] else 0
        facts_est = int(row[14]) if row[14] else 0

        # занесение данных в БД
        try:    
            sqlite_insert_with_param = f'''INSERT INTO {table_name}
                  (issl_number, difficult, issl_in_date, initiator_organ, initiator_territory,
                  initiator_fio, mat_number, UK_state, exps_fio, objs_count, issl_status,
                  issl_end_date, issl_result, issl_days_count, crime_persons_est, facts_est)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''

            data_tuple = (issl_num, difficult, issl_date_in, init_organ, init_ter, init_fio,
            mat_number, uk_state, exp_fio, objs_count, issl_status, issl_end_date,
            issl_result, issl_days_count, persons_est, facts_est)
            cursor.execute(sqlite_insert_with_param, data_tuple)
        except sqlite3.Error as error:
            log_data(f'{row[0]}: {str(error)}', 'sql_errors.log')
            res_with_errors = 1
    if sqlite_connection:
        sqlite_connection.commit()
        sqlite_connection.close()
    return res_with_errors

def sm_issls(table_name, rows):
    # print('СМ исследования')
    pass

def sei_issls(table_name, rows):
    # print('СЭИ исследования')
    pass

# Добавление таблиц СиПД
def okti_sipd(table_name, rows):
    # print('ОКТИ СиПД')
    pass

def ofili_sipd(table_name, rows):
    # print('ОФиЛИ СиПД')
    pass

def osmi_sipd(table_name, rows):
    # print('ОСМИ СиПД')
    pass

def sei_sipd(table_name, rows):
    # print('СЭИ СиПД')
    pass

def oiti_sipd(table_name, rows):
    # print('ОиТИ СиПД')
    pass

# Консультации
def consults(table_name, rows):
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        # если нет даты проведения - стоп
        if not row[1]:
            break
        try:
            date = xls_date_to_ordinal(row[1])
        except:
            date = str_date_to_ordinal(row[1])
        exp_direct = row[2]
        init_ter = row[3]
        init_fio = row[4]
        mat_number = row[5]
        uk_state = row[6]
        exps_fio = row[7]
        cons_type = row[8]
        result = row[9]
        # работа с БД
        try:    
            sqlite_insert_with_params = f'''INSERT INTO {table_name}
              (date, direct, init_terr, init_podr, 
              mat_number, st_uk, exp_fio, type, result)
              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);'''
    # кортеж параметров
            data_tuple = (date, exp_direct, init_ter, init_fio, mat_number, uk_state, exps_fio, cons_type, result)
            cursor.execute(sqlite_insert_with_params, data_tuple)
        except sqlite3.Error as error:
            log_data(str(error), 'errors_log.txt')
            print('При обработке журналов консультаций произошла ошибка')
    sqlite_connection.commit()
    if sqlite_connection:
        sqlite_connection.close()

# Командировки
def trips(table_name, rows):
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    for row in rows:
        # если нет даты стоп
        if not row[1]:
            break
        try:
            dep_date = xls_date_to_ordinal(row[1])
        except:
            dep_date = str_date_to_ordinal(row[1])
        try:
            return_date = xls_date_to_ordinal(row[2])
        except:
            return_date = str_date_to_ordinal(row[2])
        trip_place = row[3]
        trip_target = row[4]
        init_podr = row[5]
        init_fio = row[6]
        mat_number = row[7]
        st_uk = row[8]
        exp_fio = row[9]
        trip_type = row[10]
        trip_result = row[11]
        # работа с БД
        try:    
            sqlite_insert_with_params = f'''INSERT INTO {table_name}
              (departure_date, return_date, trip_place, trip_target, 
              init_podr, init_fio, mat_number, st_uk, exp_fio, trip_type, trip_result)
              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);'''
    # кортеж параметров
            data_tuple = (dep_date, return_date, trip_place, trip_target, 
              init_podr, init_fio, mat_number, st_uk, exp_fio, trip_type, trip_result)
            cursor.execute(sqlite_insert_with_params, data_tuple)
        except sqlite3.Error as error:
            log_data(str(error), 'errors_log.txt')
            print('При обработке журналов консультаций произошла ошибка')
    sqlite_connection.commit()
    if sqlite_connection:
        sqlite_connection.close()


# не используется
def delete_data_from_table(table_name):
    delete_query = f'DELETE FROM {table_name}'
    cursor.execute(delete_query)
    sqlite_connection.commit()
    # Обнуляем Id autoincrement
    upd_ai_query = f'UPDATE SQLITE_SEQUENCE SET SEQ=0 WHERE NAME="table_name"';
    sqlite_connection.commit()

# Словари сопоставления названий листов и запросов
# Для экспертиз 
table_query_exps = {'НАЛОГ': nalog_exps, 'ИАЭ': ia_exps, 'ЛИНГВ': lingv_exps,\
               'КТЭ': kt_exps, 'ФАЭ': fa_exps, 'ФОНО': fono_exps, 'БУХГ': buh_exps,\
               'ОЦЕН': ocen_exps, 'ОИТИ': oiti_exps, 'СМЭ': sm_exps}

# Для исследований
table_query_issls = {'НАЛОГ': nalog_issls, 'ИАЭ': ia_issls, 'ЛИНГВ': lingv_issls,\
               'КТЭ': kt_issls, 'ФАЭ': fa_issls, 'ФОНО': fono_issls, 'БУХГ': buh_issls, 'СМЭ': sm_issls}

# Для СиПД
table_query_sipd = {'ОКТИ': okti_sipd, 'ОФиЛИ': ofili_sipd, 'ОСМИ': osmi_sipd, 'СЭИ': sei_sipd, 'ОИТИ': oiti_sipd}