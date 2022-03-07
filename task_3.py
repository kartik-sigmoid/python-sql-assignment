
import pandas as pd
from sqlalchemy import create_engine
import logging


class ExcelToSQL:
    def __init__(self):
        logging.basicConfig(level=logging.INFO)
        self.logger = logging
        try:
            self.engine = create_engine("postgresql+psycopg2://postgres:password@localhost:5432/assignment-python")
            self.logger.info("Engine created successfully")
        except:
            self.logger.warning("Couldn't create engine")

    def read_sheets(self, data, file):
        try:
            if data == 'Q2':
                df = pd.read_excel(file, 'Q2')
                df.to_sql(name='total_compensation', con=self.engine, if_exists='append', index=False)
        except:
            self.logger.warning("Execution unsuccessful. Exception occurred.")
        finally:
            self.logger.info("Execution Successful.")

    def convert_to_sql(self):
        with pd.ExcelFile('task_2.xlsx') as xls:
            for sheet_name in xls.sheet_names:
                self.read_sheets(sheet_name, xls)


if __name__ == "__main__":
    excel = ExcelToSQL()
    excel.convert_to_sql()
