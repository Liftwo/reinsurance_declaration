import sqlite3
import pandas as pd
from sqlite3 import Error

filename = 'D:\\團傷臨分明細'
con = sqlite3.connect(filename +'.db')
wb = pd.read_excel(filename+'.xlsx', sheetname=None)
for sheet in wb:
    wb['出單'].to_sql(sheet, con, index=False)

con.commit()
con.close()


def create_connection(db_file):
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)
    return None

database_Q = 'M:\\季帳\2019Q1\\2019Q1樞紐-1510(華南賠款保費分開).db'

