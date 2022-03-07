# !/usr/bin/python
import psycopg2
from config import config
import xlsxwriter


def write_to_excel(records):
    workbook = xlsxwriter.Workbook('task_2.xlsx')
    worksheet = workbook.add_worksheet()

    row = 0
    column = 0

    # iterating through content list
    for items in records:

        # write operation perform
        worksheet.write(row, column, items[0])
        worksheet.write(row, column + 1, items[1])
        worksheet.write(row, column + 2, items[2])
        worksheet.write(row, column + 3, items[3])
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
        query = 'DROP TABLE IF EXISTS cal_emp;'
        cur.execute(query)

        query = 'CREATE TABLE cal_emp AS TABLE jobhist;'
        cur.execute(query)

        query = 'UPDATE cal_emp SET enddate = CURRENT_DATE WHERE enddate is NULL;'
        cur.execute(query)

        query = 'select cal_emp.empno as Employee_no, ename as Employee, ("enddate"::date - "startdate"::date) / 30 * cal_emp.sal as Total_compensation, dept.dname as Department from cal_emp INNER JOIN emp ON cal_emp.empno = emp.empno INNER JOIN dept on cal_emp.deptno = dept.deptno'

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
