# !/usr/bin/python
import psycopg2
from config import config
import pandas as pd
import xlsxwriter


def write_to_excel(rows):
    c1 = []
    c2 = []
    c3 = []
    c4 = []
    c5 = []
    c6 = []

    for row in rows:
        temp_list = list(row)
        c1.append(temp_list[0])
        c2.append(temp_list[1])
        c3.append(temp_list[2])
        c4.append(temp_list[3])
        c5.append(temp_list[4])
        c6.append(temp_list[5])
    df = pd.DataFrame({'Employee Name': c1, 'Employee No': c2, 'Dept No': c3, 'Dept Name': c4, 'Total Compensation': c5,
                       'Months Spent': c6})
    writer = pd.ExcelWriter(
        '/Users/kartikjaiswal/PycharmProjects/postgres/task_2.xlsx')
    df.to_excel(writer, sheet_name='Q2', index=False)
    writer.save()


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

        cur.execute("UPDATE jobhist SET enddate=CURRENT_DATE WHERE enddate IS NULL;")
        data = cur.execute(
            "SELECT emp.ename, jh.empno, jh.deptno, dept.dname, ROUND((jh.enddate-jh.startdate)/30*jh.sal,0) "
            "AS total_compensation, ROUND((jh.enddate-jh.startdate)/30,0) as months_spent FROM "
            "jobhist as jh INNER JOIN dept ON jh.deptno=dept.deptno INNER JOIN emp ON jh.empno=emp.empno;")
        rows = cur.fetchall()

        # writing files to xlsx file
        write_to_excel(rows)

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
