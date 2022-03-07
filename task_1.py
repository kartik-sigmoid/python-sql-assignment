# !/usr/bin/python
import psycopg2
from config import config
import xlsxwriter


def write_to_excel(records):
    workbook = xlsxwriter.Workbook('task_1.xlsx')
    worksheet = workbook.add_worksheet()

    row = 0
    column = 0

    # iterating through content list
    for emp_no, emp_name, manager in records:

        # write operation perform
        worksheet.write(row, column, emp_no)
        worksheet.write(row, column + 1, emp_name)
        worksheet.write(row, column + 2, manager)
        row += 1

    workbook.close()


def connect():
    """ Connect to the PostgreSQL database server """
    conn = None
    try:
        # read connection parameters
        params = config()

        # connect to the PostgreSQL server
        print('Connecting to the PostgreSQL database...')
        conn = psycopg2.connect(**params)

        # create a cursor
        cur = conn.cursor()

        # query
        query = 'SELECT empno, ename, mgr from emp;'
        cur.execute(query)
        records = cur.fetchall()

        # writing files to xlsx file
        write_to_excel(records)

        # close the communication with the PostgreSQL
        cur.close()
    except (Exception, psycopg2.DatabaseError) as error:
        print(error)
    finally:
        if conn is not None:
            conn.close()
            print('Database connection closed.')


if __name__ == '__main__':
    connect()
