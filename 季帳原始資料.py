import sqlite3
import pandas as pd
from sqlite3 import Error


filename = '2019QX原始資料'
con = sqlite3.connect(filename +'.db')
wb = pd.read_excel('M:\季帳\\2019Q1\\2019Q1樞紐-1510(華南賠款保費分開).xlsx', sheet_name='原始資料')
for sheet in wb:
    wb['原始資料'].to_sql(sheet, con, index=False)

con.commit()
con.close()




