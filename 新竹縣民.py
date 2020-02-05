import pandas as pd
import sqlite3

db = sqlite3.connect('D:\\新竹縣民賠款.db')
dfs = pd.read_excel('D:\\2019Q2_新竹縣民團傷Bordereaux(2017&2018).xlsx', sheetname=None)
for table, df in dfs.items():
    df.to_sql(table,db)

cur.execute("SELECT * FROM table where")
cur.fetchall()

cur.execute("SELECT NameInsured FROM fork WHERE NameInsured='張仁昌'")
cur.fetchall()

cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = cur.fetchall()
for t in tables:
    t[0]

for i in table:
    SELECT NameInsured FROM i WHERE NameInsured='張仁昌'
    if exitst:
        print(i)