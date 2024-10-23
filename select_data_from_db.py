import sqlite3
from datetime import datetime

# Фукнция обработки ошибок в файл
def log_data(data, file_name):
     with open(file_name, 'a', encoding='UTF-8') as file:
          file.write(str(datetime.now()))
          file.write('\n')
          file.write(data)
          file.write('\n\n')

# функция отбора данных из базы
# с указанием отбора одной записи или всех
def query_data(query, fetch='one'):
    sqlite_connection = sqlite3.connect('CommonDB.db')
    cursor = sqlite_connection.cursor()
    sqlite_select_query = query
    cursor.execute(sqlite_select_query)
    if fetch == 'one':
        # почему-то приходит кортеж с одним значением 
        query_result = cursor.fetchone()[0]
        # иногда в сложных запросах возвращается None вместо 0
        if query_result is None:
            query_result = 0
    else:
        query_result = cursor.fetchall()
    cursor.close()
    sqlite_connection.close()
    return query_result

 # поступившие
# отбор поступивших экспертиз (всего)
def get_recieved_exp_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# поступило материалов
def get_recieved_mats_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT mat_number)
    FROM {table_name}
    WHERE exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# поступило первичных
def get_primary_recieved_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Первичная'
    AND exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# поступило дополнительных
def get_additional_recieved_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Дополнительная'
    AND exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# поступило повторных
def get_repeated_recieved_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Повторная'
    AND exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# поступило комиссионных 
def get_comission_recieved_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Комиссионная'
    AND exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# поступило дополнительных комиссионных
def get_additional_comission_recieved_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Дополнительная комиссионная'
    AND exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# поступило повторных комиссионных
def get_repeated_comission_recieved_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Повторная комиссионная'
    AND exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# поступило комплексных
def get_complex_recieved_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Комплексная'
    AND exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# поступило дополнительных комплексных
def get_additional_complex_recieved_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Дополнительная комплексная'
    AND exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# поступило повторных комплексных
def get_repeated_complex_recieved_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Повторная комплексная'
    AND exp_in_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

    # выполненные
# материалов по выполненным за весь год
def get_mats_complited(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT mat_number)
    FROM {table_name}
    WHERE exp_status = 'Сдана'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# всего выполнено
def get_complited_exps(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_status = 'Сдана'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

# объектов по выполненным
def get_complited_objs(start_date, end_date, table_name):
    # решается через табличное выражение
    # для отбора уникальных номеров экспертиз/исследований
    # в которых выбирается максимальный показатель объектов
    # и в основном запросе берётся сумма
    query = f'''
    WITH get_uniq_objs_by_exps AS 
        (SELECT DISTINCT exp_number, MAX(objs_count) AS objs_count
        FROM {table_name}
        WHERE exp_status = 'Сдана'
        AND initiator_fio NOT LIKE '%СВО'
        AND exp_end_date
        BETWEEN '{start_date}' AND '{end_date}'
        GROUP BY exp_number
        ORDER BY objs_count DESC)
    SELECT SUM(objs_count) FROM get_uniq_objs_by_exps
    '''
    # иногда прилетает None, что рушит всё
    res = query_data(query)
    return res

# возвращённых без исполнения
def get_without_exec(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_status = 'Без исполнения'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

    # Для выполнения данного запроса требуется Python > 3.8
# исполненные сроки (все)
# запрос отберёт сроки исполнения в формате строки (+0000-00-22 00:00:00.000),
# из которой отберёт подстроку с месяцем
# в итоге запрос вернёт таблицу
def get_complited_terms(start_date, end_date, table_name, status, terms_condition):
    # запрос через табличное выражение работает некорректно
    # query = f'''
    # WITH get_month AS 
    #     (SELECT
    #     SUBSTRING(TIMEDIFF(exp_end_date, exp_in_date), 7, 2) AS month
    #     FROM {table_name}
    #      WHERE exp_status = 'Сдана'
    #     AND exp_end_date
    #     BETWEEN '{start_date}' AND '{end_date}'
    #     GROUP BY exp_number)
    # SELECT COUNT(*) AS month_count
    # FROM get_month
    # HAVING month = '01'
    # GROUP BY month  
    # '''
    query = f'''SELECT COUNT(month)
        FROM 
          (SELECT
            SUBSTRING(TIMEDIFF(exp_end_date, exp_in_date), 7, 2) AS month
            FROM  {table_name}
            WHERE {status}
            AND initiator_fio NOT LIKE '%СВО'
            AND exp_end_date
            BETWEEN '{start_date}' AND '{end_date}'
            GROUP BY exp_number)
        WHERE {terms_condition}
        GROUP BY month
    '''
    # ??? запрос не работает с fetchone(), если указать условие WHERE 
    # fetchall нужен для отбора всех записей
    res = query_data(query, fetch='all')
    return sum((x[0]) for x in res)


# выполненные по видам
# выполнено первичных
def get_primary_completed_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Первичная'
    AND exp_status = 'Сдана'
    AND initiator_fio NOT LIKE '%СВО'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    '''
    return query_data(query)

# выполнено дополнительных
def get_additional_completed_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Дополнительная'
    AND exp_status = 'Сдана'
    AND initiator_fio NOT LIKE '%СВО'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    '''
    return query_data(query)

# выполнено повторных
def get_repeated_completed_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Повторная'
    AND exp_status = 'Сдана'
    AND initiator_fio NOT LIKE '%СВО'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    '''
    return query_data(query)

# выполнено комиссионных 
def get_comission_completed_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Комиссионная'
    AND exp_status = 'Сдана'
    AND initiator_fio NOT LIKE '%СВО'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    '''
    return query_data(query)

# выполнено дополнительных комиссионных
def get_additional_comission_completed_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Дополнительная комиссионная'
    AND exp_status = 'Сдана'
    AND initiator_fio NOT LIKE '%СВО'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    '''
    return query_data(query)

# выполнено повторных комиссионных
def get_repeated_comission_completed_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Повторная комиссионная'
    AND exp_status = 'Сдана'
    AND initiator_fio NOT LIKE '%СВО'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    '''
    return query_data(query)

# выполнено комплексных
def get_complex_completed_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Комплексная'
    AND exp_status = 'Сдана'
    AND initiator_fio NOT LIKE '%СВО'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    '''
    return query_data(query)

# выполнено дополнительных комплексных
def get_additional_complex_completed_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Дополнительная комплексная'
    AND exp_status = 'Сдана'
    AND initiator_fio NOT LIKE '%СВО'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    '''
    return query_data(query)

# выполнено повторных комплексных
def get_repeated_complex_completed_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_type = 'Повторная комплексная'
    AND exp_status = 'Сдана'
    AND initiator_fio NOT LIKE '%СВО'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    '''
    return query_data(query)


# НПВ
# выполнено повторных комплексных
def get_npv_count(start_date, end_date, table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE 
    exp_result = 'НПВ'
    AND initiator_fio NOT LIKE '%СВО'
    AND exp_end_date
    BETWEEN '{start_date}' AND '{end_date}'
    '''
    return query_data(query)

# установлено возможно причастных лиц
def get_persons_est_count(start_date, end_date, table_name):
    query = f'''
    WITH get_pers_est AS
        (SELECT exp_number, MAX(crime_persons_est) AS p_est
        FROM {table_name}
        WHERE exp_end_date
        BETWEEN '{start_date}' AND '{end_date}'
        AND initiator_fio NOT LIKE '%СВО'
        GROUP BY exp_number)
    SELECT SUM(p_est)
    FROM get_pers_est    
    '''
    return query_data(query)


def get_facts_est_count(start_date, end_date, table_name):
    query = f'''
    WITH get_facts_est AS
        (SELECT exp_number, MAX(facts_est) AS f_est
        FROM {table_name}
        WHERE exp_end_date 
        BETWEEN '{start_date}' AND '{end_date}'
        AND initiator_fio NOT LIKE '%СВО'
        GROUP BY exp_number)
    SELECT SUM(f_est)
    FROM get_facts_est  
    '''
    return query_data(query)


def get_work_count(table_name):
    query = f'''
    SELECT COUNT(DISTINCT exp_number)
    FROM {table_name}
    WHERE exp_status = 'В производстве'
    AND initiator_fio NOT LIKE '%СВО'
    '''
    return query_data(query)

def get_work_objs_count(table_name):
    query = f'''
    WITH get_uniq_objs_by_exps AS 
        (SELECT DISTINCT exp_number, MAX(objs_count) AS objs_count
        FROM {table_name}
        WHERE exp_status = 'В производстве'
        AND initiator_fio NOT LIKE '%СВО'
        GROUP BY exp_number
        ORDER BY objs_count DESC)
    SELECT SUM(objs_count) FROM get_uniq_objs_by_exps
    '''
    return query_data(query)


# Для выполнения данного запроса требуется Python > 3.8
# сроки в производстве
def get_work_terms(end_date, table_name, status, terms_condition):
    query = f'''SELECT COUNT(month)
        FROM 
          (SELECT
            SUBSTRING(TIMEDIFF({end_date}, exp_in_date), 7, 2) AS month
            FROM  {table_name}
            WHERE {status}
            AND initiator_fio NOT LIKE '%СВО'
            GROUP BY exp_number)
        WHERE {terms_condition}
        GROUP BY month
    '''
    # ??? запрос не работает с fetchone(), если указать условие WHERE 
    # fetchall нужен для отбора всех записей
    res = query_data(query, fetch='all')
    return sum((x[0]) for x in res)