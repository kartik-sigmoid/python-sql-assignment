import psycopg2
import pandas as pd
import logging


class SQLtoExcel:

    def __init__(self):
        logging.basicConfig(level=logging.INFO)
        self.logger = logging
        try:
            self.connection = psycopg2.connect(database="assignment-python", user="postgres", password="password",
                                          host="localhost", port=5432)
            self.logger.info(msg="Database successfully connected.")
        except:
            self.logger.warning(msg="Could not connect to database.")

    def convert_to_excel(self):
        try:
            cur = self.connection.cursor()

            cur.execute("UPDATE jobhist SET enddate=CURRENT_DATE WHERE enddate IS NULL;")
            data = cur.execute(
                "SELECT emp.ename, jh.empno, jh.deptno, dept.dname, ROUND((jh.enddate-jh.startdate)/30*jh.sal,0) "
                "AS total_compensation, ROUND((jh.enddate-jh.startdate)/30,0) as months_spent FROM "
                "jobhist as jh INNER JOIN dept ON jh.deptno=dept.deptno INNER JOIN emp ON jh.empno=emp.empno;")
            rows = cur.fetchall()
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
            df = pd.DataFrame({'Employee Name': c1, 'Employee No': c2, 'Dept No': c3, 'Dept Name': c4, 'Total Compensation': c5, 'Months Spent': c6})
            writer = pd.ExcelWriter('/Users/kartikjaiswal/PycharmProjects/postgres/task_2.xlsx')
            df.to_excel(writer, sheet_name='Q2', index=False)
            writer.save()

        except:
            self.logger.warning("Execution unsuccessful. Exception occurred.")

        finally:
            self.logger.info("Execution Successful.")
            self.connection.close()
            return 1


if __name__ == "__main__":
    sql_table = SQLtoExcel()
    sql_table.convert_to_excel()

