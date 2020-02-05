import pandas as pd
import os
from functools import reduce

df = pd.read_excel('M:\\保單、名冊\\2019被保人名冊.xlsx', '工作表1')
df.columns = ['a']

df = df[df['a'].str.contains('台灣產物') == False]
df = df[df['a'].str.contains('團體傷害') == False]
df = df[df['a'].str.contains('保單號碼') == False]

df1 = df.iloc[:1]
df2 = df.iloc[1:df.shape[0]]
df2 = df2.reset_index()
df2 = df2.drop('index', axis=1)
for i in df2[df2['a'].str.contains('序號') == True].index.tolist():
    df2 = df2.drop(i)

if df2['a'].str.contains('醫療日額').any():
    for i in df2[df2['a'].str.contains('醫療日額') == True].index.tolist():
        df2 = df2.drop(i)

df1 = df1['a'].str.replace('\u3000\u3000\u3000', '')
df1 = df1.to_frame(name='a')
df1 = df1['a'].str.replace('\u3000\u3000', '')
df1 = df1.to_frame(name='a')
df1 = df1['a'].str.replace('工\u3000作\u3000性\u3000質', '工作性質')
df1 = df1.to_frame(name='a')
df1 = df1['a'].str.replace('\u3000出生日期 ', '\u3000出生日期')
df1 = df1.to_frame(name='a')
df2 = df2['a'].replace('\s+', '\u3000', regex=True)

df2 = df2.to_frame(name='a')
repls = {'經\u3000理': '經理', '廠\u3000長': '廠長', '協\u3000理': '協理', '副\u3000理': '副理',
         '組\u3000長': '組長', '服\u3000務專員': '服務專員', '課\u3000長': '課長',
         '服\u3000務\u3000員': '服務員', '服\u3000務': '服務', '主\u3000任': '主任',
         '\u3000-': '-', '配偶\u3000法定': '配偶法定', '母女\u3000法定': '母女法定',
         '母子\u3000法定': '母子法定', '父子\u3000法定': '父子法定', '\u3000-母子': '-母子',
         '父女\u3000法定': '父女法定', '主\u3000管': '主管', '負\u3000責人': '負責人',
         '銀行員\u3000工': '銀行員工', '自\u3000由業': '自由業', '聘僱\u3000員': '聘僱員', '員\u3000工': '員工',
         '老\u3000板': '老板', '內勤-\u3000': '內勤-', '董\u3000事會': '董事會', '總\u3000經理室': '總經理室', '董\u3000事長':'董事長',
         '監\u3000工':'監工'}
for x, y in repls.items():
    df2['a'] = df2['a'].str.replace(x, y)

df1 = pd.DataFrame(df1.a.str.split('\u3000').tolist())
df2 = pd.DataFrame(df2.a.str.split('\u3000').tolist())
df1 = df1.reset_index()
df1 = df1.drop('index', axis=1)

new_col = []
for i in df2[1]:
    new_col.append(i[:-1])
df2.insert(2, '2.1', new_col)
new_col2 = []
for i in df2[1]:
    new_col2.append(i[-1:])
new_col3 = []
for i in new_col2:
    new_col3.append(i.replace(i, 'O'))

df2.insert(2, '2.2', new_col3)
df2['2t'] = df2['2.1'].astype(str) + df2['2.2']
df2 = df2.drop('2.1', axis=1)
df2 = df2.drop('2.2', axis=1)
df2 = df2.drop(1, axis=1)
df2.insert(1, 1, df2['2t'])
df2 = df2.drop('2t', axis=1)

for i in df2[2]:
    df2[2] = df2[2].str.replace(i[3:9], '')

for i in df2[3]:
    df2[3] = df2[3].str.replace(i[0:7], 'OOOOOOO')

final = pd.concat([df1, df2], axis=0)
final.drop
writer = pd.ExcelWriter('D:\名冊name.xlsx', engine='openpyxl')
final.to_excel(writer, '工作表5', index=False, header=None)
writer.save()
os.startfile('D:\名冊name.xlsx')


                                                     ''