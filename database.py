import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()


class Database:

    # Все данные для отображения
    select = """
    SELECT 
    a.type AS activity_type, 
    r.number, 
    r.letter, 
    r.indicator, 
    e.name AS employee_name, 
    e.post AS employee_post
    FROM 
    Responsibilities r
    LEFT JOIN 
    Activities a ON r.activity_id = a.id
    LEFT JOIN 
    Employee_responsibility er ON r.id = er.responsibility_id
    LEFT JOIN 
    Employees e ON er.employee_id = e.id;"""

    # Логика подключения к БД
    _connection = None

    @classmethod
    def get_connection(cls):
        if cls._connection is None:
            try:
                cls._connection = psycopg2.connect(
                    dbname=os.getenv("DB_NAME"),
                    user=os.getenv("DB_USER"),
                    password=os.getenv("DB_PASSWORD"),
                    host=os.getenv("DB_HOST"),
                    port=os.getenv("DB_PORT"),
                )
                print("Соединение с PostgreSQL установлено")
            except Exception as error:
                print("Ошибка при подключении к PostgreSQL", error)
                cls._connection = None
        return cls._connection

    @classmethod
    def close_connection(cls):
        if cls._connection:
            cls._connection.close()
            print("Соединение с PostgreSQL закрыто")
            cls._connection = None

    # Метод получения метрик
    @classmethod
    def get_metrics(cls):
        with Database.get_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(cls.selectMetricsSQL)
                results = cursor.fetchall()
                return results

    # Метод подключения групп
    @classmethod
    def get_sections(cls):
        with Database.get_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(cls.selectSectionsSQL)
                results = cursor.fetchall()
                return results
