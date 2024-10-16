import sqlite3

## Запросы на создание таблиц
# Структура таблицы КТЭ экспертиз
kt_exps = '''
"id" integer not null primary key unique,
"exp_number" text not null,
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

# Структура таблицы ИАЭ экспертиз
ia_exps = '''
"id" integer not null primary key unique,
"exp_number" text not null,
"difficult" text,
"exp_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,
"UK_state" text,
"exps_fio" text,
"objs_count" integer,
"exp_type" text,
"exp_status" text,
"exp_end_date" text,
"exp_result" text,
"exp_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы лингвистики эксп.
lingv_exps = '''
"id" integer not null primary key unique,
"exp_number" text not null,
"difficult" text,
"exp_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"objs_info" text,
"objs_count" integer,
"exp_type" text,
"exp_status" text,
"exp_end_date" text,
"exp_result" text,
"exp_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы фоно эксп.
fono_exps = '''
"id" integer not null primary key unique,
"exp_number" text not null,
"difficult" text,
"exp_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"acoustic_exps_fio" text,
"lingv_exps_fio" text,
"objs_info" text,
"objs_count" integer,
"exp_type" text,
"exp_status" text,
"exp_end_date" text,
"exp_result" text,
"verbatim_duration" integer,
"idents_count" integer,
"persons_est" integer,
"exp_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы бух. эксп.
buh_exps = '''
"id" integer not null primary key unique,
"exp_number" text not null,
"difficult" text,
"exp_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"objs_count" integer,
"exp_type" text,
"exp_status" text,
"exp_end_date" text,
"exp_result" text,
"exp_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы ФАЭ
fa_exps = '''
"id" integer not null primary key unique,
"exp_number" text not null,
"difficult" text,
"exp_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"objs_count" integer,
"exp_type" text,
"exp_status" text,
"exp_end_date" text,
"exp_result" text,
"exp_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы налог. эксп.
nalog_exps = '''
"id" integer not null primary key unique,
"exp_number" text not null,
"difficult" text,
"exp_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"objs_count" integer,
"exp_type" text,
"exp_status" text,
"exp_end_date" text,
"exp_result" text,
"exp_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы оцен. эксп.
ocen_exps = '''
"id" integer not null primary key unique,
"exp_number" text not null,
"difficult" text,
"exp_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"objs_count" integer,
"exp_type" text,
"exp_status" text,
"exp_end_date" text,
"exp_result" text,
"exp_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы СМЭ 
sm_exps = '''
"id" integer not null primary key unique,
"exp_number" text not null,
"difficult" text,
"exp_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"exps_fio_complex" text,
"exp_type" text,
"exp_status" text,
"exp_end_date" text,
"exp_result" text,
"objs_count" integer,
"patient_fio" text,
"patient_status" text,
"med_docs" text,
"exp_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы ОИТИ экспертиз
oiti_exps = '''
"id" integer not null primary key unique,
"exp_number" text not null,
"difficult" text,
"exp_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"objs_count" integer,
"exp_type" text,
"exp_status" text,
"exp_end_date" text,
"exp_result" text,
"exp_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

################################
## ИССЛЕДОВАНИЯ 
# КТ исследования
kt_issls = '''
"id" integer not null primary key unique,
"issl_number"  text,
"difficult" text,
"issl_in_date"  text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,
"uk_state" text,
"fabula" text,
"exps_fio" text,
"objs_info" text,
"objs_count" integer,
"objs_mobile" integer,
"objs_digital" integer,
"issl_status" text,
"issl_end_date" integer,
"issl_result" text,
"issl_days_count" integer,
"facts_est" integer,
"issl_vyvod" text,
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

# Структура таблицы ИА исследований
ia_issls = '''
"id" integer not null primary key unique,
"issl_number" text,
"difficult" text,
"issl_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,
"UK_state" text,
"exps_fio" text,
"objs_count" integer,
"issl_status" text,
"issl_end_date" text,
"issl_result" text,
"issl_days_count" integer,
"facts_est" integer,
"issl_vyvod" text
'''

# Структура таблицы бух. иссл. 
buh_issls = '''
"id" integer not null primary key unique,
"issl_number" text,
"difficult" text,
"issl_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"objs_count" integer,
"issl_status" text,
"issl_end_date" text,
"issl_result" text,
"issl_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы ФАИ
fa_issls = '''
"id" integer not null primary key unique,
"issl_number" text,
"difficult" text,
"issl_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"objs_count" integer,
"issl_status" text,
"issl_end_date" text,
"issl_result" text,
"issl_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы налог. иссл.
nalog_issls = '''
"id" integer not null primary key unique,
"issl_number" text,
"difficult" text,
"issl_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"objs_count" integer,
"issl_status" text,
"issl_end_date" text,
"issl_result" text,
"issl_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы лингвистики иссл.
lingv_issls = '''
"id" integer not null primary key unique,
"issl_number" text,
"difficult" text,
"issl_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"objs_info" text,
"objs_count" integer,
"issl_status" text,
"issl_end_date" text,
"issl_result" text,
"issl_days_count" integer,
"persons_est" integer,
"facts_est" integer
'''

# Структура таблицы фоно иссл.
fono_issls = '''
"id" integer not null primary key unique,
"issl_number" text,
"difficult" text,
"issl_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"acoustic_exps_fio" text,
"lingv_exps_fio" text,
"objs_info" text,
"objs_count" integer,
"issl_status" text,
"issl_end_date" text,
"issl_result" text,
"verbatim_duration" integer,
"idents_count" integer,
"persons_est" integer,
"issl_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

# Структура таблицы СМИ 
sm_issls = '''
"id" integer not null primary key unique,
"issl_number" text,
"difficult" text,
"issl_in_date" text,
"initiator_organ" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,  
"UK_state" text,
"exps_fio" text,
"exps_fio_complex" text,
"issl_status" text,
"issl_end_date" text,
"issl_result" text,
"objs_count" integer,
"patient_fio" text,
"patient_status" text,
"med_docs" text,
"issl_type" text,
"issl_days_count" integer,
"crime_persons_est" integer,
"facts_est" integer
'''

##############################################
## Таблицы СиПД

# СЭИ
sei_sipd = '''
"id" integer not null primary key unique,
"mat_in_date" text,
"action_date" text,
"action_type" text,
"initiator_podr" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,
"st_uk" text,
"exp_fio" text,
"objs_info" text,
"objs_count" integer,
"action_result" text,
"comments" text
'''

# ОКТИ
okti_sipd = '''
"id" integer not null primary key unique,
"date_mat_in" text,
"action_date" text,
"action_type" text,
"initiator_podr" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,
"st_uk" text,
"fabula" text,
"exp_fio" text,
"objs_info" text,
"objs_all_count" integer,
"objs_first_mobile" integer,
"objs_first_digital" integer,
"action_result" text,
"objs_finish_count" text,
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

# ОФиЛИ
ofili_sipd = '''
"id" integer not null primary key unique,
"mat_in_date" text,
"action_date" text,
"action_type" text,
"initiator_podr" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,
"st_uk" text,
"exp_fio" text,
"acoustic_exps_fio" text,
"objs_info" text,
"objs_count" integer,
"action_result" text
'''

# ОСМИ 
osmi_sipd = '''
"id" integer not null primary key unique,
"mat_in_date" text,
"action_date" text,
"action_type" text,
"initiator_podr" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,
"st_uk" text,
"exp_fio" text,
"objs_count" integer,
"action_result" text,
"patient_fio" text,
"patient status" text,
"med_docs" text
'''

# ОИТИ 
oiti_sipd = '''
"id" integer not null primary key unique,
"mat_in_date" text,
"action_date" text,
"action_type" text,
"initiator_podr" text,
"initiator_territory" text,
"initiator_fio" text,
"mat_number" text,
"st_uk" text,
"exp_fio" text,
"objs_info" text,
"objs_count" integer,
"action_result" text
'''

##############################################
# Таблица консультаций
consults = '''
"id" integer not null primary key unique,
"date" text,
"direct" text,
"initiator_territory" text,
"initiator_podr" text,
"mat_number" text,
"st_uk" text,
"exp_fio" text,
"type" text,
"result" text 
'''

# Таблица командировок
trips = '''
"id" integer not null primary key unique,
"departure_date" text,
"return_date" text,
"trip_place" text,
"trip_target" text,
"initiator_podr" text,
"initiator_fio" text,
"mat_number" text,
"st_uk" text,
"exp_fio" text,
"trip_type" text,
"trip_result" text
'''

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

def create_all_tables(current_week):
    # экспертизы 
    create_table_in_db(table_name = f'Week_{current_week}_КТЭ_Exps', query=kt_exps)
    create_table_in_db(table_name = f'Week_{current_week}_ИАЭ_Exps', query=ia_exps)
    create_table_in_db(table_name = f'Week_{current_week}_ЛИНГВ_Exps', query=lingv_exps)
    create_table_in_db(table_name = f'Week_{current_week}_ФОНО_Exps', query=fono_exps)
    create_table_in_db(table_name = f'Week_{current_week}_БУХГ_Exps', query=buh_exps)
    create_table_in_db(table_name = f'Week_{current_week}_ФАЭ_Exps', query=fa_exps)
    create_table_in_db(table_name = f'Week_{current_week}_НАЛОГ_Exps', query=nalog_exps)
    create_table_in_db(table_name = f'Week_{current_week}_ОЦЕН_Exps', query=ocen_exps)
    create_table_in_db(table_name = f'Week_{current_week}_СМЭ_Exps', query=sm_exps)
    create_table_in_db(table_name = f'Week_{current_week}_ОИТИ_Exps', query=oiti_exps)
    # исследования 
    create_table_in_db(table_name = f'Week_{current_week}_КТЭ_Issls', query=kt_issls)
    create_table_in_db(table_name = f'Week_{current_week}_ИАЭ_Issls', query=ia_issls)
    create_table_in_db(table_name = f'Week_{current_week}_ЛИНГВ_Issls', query=lingv_issls)
    create_table_in_db(table_name = f'Week_{current_week}_ФОНО_Issls', query=fono_issls)
    create_table_in_db(table_name = f'Week_{current_week}_БУХГ_Issls', query=buh_issls)
    create_table_in_db(table_name = f'Week_{current_week}_ФАЭ_Issls', query=fa_issls)
    create_table_in_db(table_name = f'Week_{current_week}_НАЛОГ_Issls', query=nalog_issls)
    create_table_in_db(table_name = f'Week_{current_week}_СМЭ_Issls', query=sm_issls)
    # СиПД 
    create_table_in_db(table_name = f'Week_{current_week}_СЭИ_SiPD', query=sei_sipd)
    create_table_in_db(table_name = f'Week_{current_week}_ОКТИ_SiPD', query=okti_sipd)
    create_table_in_db(table_name = f'Week_{current_week}_ОИТИ_SiPD', query=oiti_sipd)
    create_table_in_db(table_name = f'Week_{current_week}_ОСМИ_SiPD', query=osmi_sipd)
    create_table_in_db(table_name = f'Week_{current_week}_ОФиЛИ_SiPD', query=ofili_sipd)
    # Консультации и командировки
    create_table_in_db(table_name = f'Week_{current_week}_Consults', query=consults)
    create_table_in_db(table_name = f'Week_{current_week}_Trips', query=trips)