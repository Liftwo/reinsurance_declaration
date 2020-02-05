import sqlite3

####創建資料庫####
con = sqlite3.connect('path.filename.db')
df = pd.read_excel('excelpath', sheet_name='出單')
df.to_sql('出單', con)

####更新資料庫####
df_insert = df.iloc[-1:]
df_insert.to_sql('出單', con, if_exists='append')


####查詢資料庫####
con = sqlite3.connect('C:\\Users\\3240\季帳.db')
con.row_factory = lambda cursor, row: row[0]
cur = con.cursor()

cur.execute("SELECT 分保合約編號 FROM 出單 WHERE 再保人=='marsh'")
result = cur.fetchall()
print(result)
con.close()

####匯出選擇的資料####



