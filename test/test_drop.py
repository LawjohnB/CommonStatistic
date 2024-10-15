import sqlite3

# Функция создания таблицы в соответствии с номером недели и названием листа
def create_table_in_db(table_name):
    # подключаемся к базе
    sqlite_connection = sqlite3.connect('Все журналы.db')
    cursor = sqlite_connection.cursor()
    cursor.execute(f'''
        DROP TABLE IF EXISTS {table_name}''')
    sqlite_connection.commit()
    cursor.execute(f'''
        CREATE TABLE IF NOT EXISTS {table_name}
        ({kte_exp})
        ''')
    sqlite_connection.commit()
    if sqlite_connection:
        sqlite_connection.close()

kte_exp = '''
"id" integer not null primary key unique,
"exp_number" text,
"difficult" text'''

create_table_in_db('test')