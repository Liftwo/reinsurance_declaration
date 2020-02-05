import pandas as pd
from win32com.client import Dispatch
import numpy as np
import string
import os
from os import walk
from os.path import join
import decimal
from decimal import *
import itertools

xl = Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
thisquarter = input('Sep.2019')
searchfor_1 = input('HSL-1')
df = pd.read_excel('D:\py\\account.xlsx', sheet_name='test')
df = df[(df['帳單類別'] == 'N1') | (df['帳單類別'] == 'N2')]
df = df.loc[df['實際帳年月'] != 201712]
df = df[df['再保公司'].str.contains(searchfor_1)]
searchfor_2 = ['FB', 'QC']
df = df[df['合約號碼'].str.contains('|'.join(searchfor_2))]
df = df.fillna(0)

table1 = pd.pivot_table(df, values='PREMIUM(NT$)', columns='實際帳年月', index='合約號碼', aggfunc=np.sum)
table_loss = pd.pivot_table(df, values='LOSS PAID(原幣)', columns='實際帳年月', index='合約號碼', aggfunc=np.sum)

df_t1 = table1.reset_index()
df_t1 = df_t1.set_index('合約號碼')
df_t1 = df_t1.fillna(0)
df_t1 = df_t1.cumsum(axis=1)
df_t1['number'] = df_t1.index
q = [i for i in df_t1.columns if '20' in str(i)]


df_loss = table_loss.reset_index()
df_loss = df_loss.set_index('合約號碼')
df_loss = df_loss.fillna(0)
df_loss = df_loss.cumsum(axis=1)
df_loss['number'] = df_loss.index

list_re = pd.read_excel('D:\\團傷臨分明細.xlsx', sheet_name='出單')
list_re = list_re.rename(columns={'分保合約編號': 'number'})

li = []
for root, dirs, files in walk('D:\\py\\outstanding'):
    for f in files:
        fullpath = join(root, f)
        do = pd.read_excel(fullpath)
        do['quarter'] = f[0:6]
        li.append(do)


def point_to_int(df, col):
    df[col] = df[col].apply(lambda x: Decimal(x).quantize(Decimal('0'), rounding=decimal.ROUND_HALF_UP))
    return df[col]


outs = pd.concat(li, axis=0, ignore_index=True)
outs = outs[outs['再保模組'].str.contains('FB')]
re_m = pd.read_excel('D:\\再保模組對照表.xlsx', sheet_name='季帳-表頭')
outstanding_0 = outs.merge(re_m, on='再保模組')
outstanding_0 = outstanding_0.rename(columns={'分保合約編號':'number'})
table_outstanding = pd.pivot_table(outstanding_0, values='(A+B)未付合計', columns='quarter', index='number')
df_os = table_outstanding.merge(list_re, on='number')
searchfor_3 = input('再保人')
df_os = df_os[df_os['再保人'].str.contains('|'.join(searchfor_3))]
df_os = df_os.fillna(0)
quarter_os = [c for c in df_os.columns if '20' in str(c)]
for q in quarter_os:
    point_to_int(df_os, q)
writer_out = pd.ExcelWriter('D:\\py\\outstanding.xlsx')
df_os.to_excel(writer_out)
writer_out.save()

result = df_t1.merge(list_re, on='number')
result = result.fillna(0)
quarter = [c for c in result.columns if '20' in str(c)]
result['分保比例'] = (result['分保比例']*100).astype(int)
result['佣金率'] = (result['佣金率']*100).astype(int)
result['分保比例'] = result['分保比例']/100


def gross(a, b):
    return a / b


for i in quarter:
    result[i] = gross(result[i], result['分保比例'])

writer_result = pd.ExcelWriter('D:\\py\\result.xlsx')
result.to_excel(writer_result)
writer_result.save()

result_loss = df_loss.merge(list_re, on='number')
result_loss = result_loss.fillna(0)
result_loss['分保比例'] = (result_loss['分保比例']*100).astype(int)
result_loss['分保比例'] = result_loss['分保比例']/100
for i in quarter:
    result_loss[i] = gross(result_loss[i], result_loss['分保比例'])

writer_result_loss = pd.ExcelWriter('D:\\py\\result_loss.xlsx')
result_loss.to_excel(writer_result_loss)
writer_result_loss.save()


alpha = [i for i in string.ascii_uppercase[2::]]
alpha_all = [i for i in string.ascii_uppercase]
for i in alpha_all:
        for j in string.ascii_uppercase:
            alpha.append(i+j)


def months(q, row, col):
    col_amount = []
    for i in range(len(q)):
        col_amount.append(alpha[i])
    if col == 'C':
        cell = col + row + ':' + col_amount[-1] + row
    else:
        cell = col_amount[-1] + row + ':' + col_amount[-1] + row
    return cell


def period(p, row):
    periods = [i*3 for i in range(1,len(p)+1)]
    value = [str(i)+' '+ 'Months' for i in periods]
    return value


save_path = 'D:\\py\\head'
this_quarter = 'For Ending {}'.format(thisquarter)
xlLeft, xlRight, xlCenter = -4131, -4152, -4108
xlEdgeBottom = 9
xlEdgeTop = 8
xlEdgeRight = 10
xlEdgeLeft = 7

for root, dirs, files in walk('D:\\py\\head'):
    for f in files:
        fullpath = join(root, f)

for i, j, k, l, m, n in zip(result['合約名稱'], result['合約年'], result['佣金率'], result['分保比例'], range(len(result)), range(len(result_loss))):
    wb = xl.Workbooks.Open('D:\\py\\表頭(久津實業)4.xlsx')
    ws = wb.WorkSheets('Trangle -1')
    ws.Range('A4').Value = i
    ws.Range('A6').Value = j
    ws.Range('A11').Value = "Remark: RI" + ' ' + str(l*100) + '%'
    a = result.loc[m, quarter]
    a = a.tolist()
    al = result_loss.loc[n, quarter]
    al = al.tolist()
    for x, y in enumerate(a):
        if y != 0:
            start = a[x::]
            start = [int(g) for g in start]
            start = ['{:,}'.format(g) for g in start]
            start_loss = al[x::]
            start_loss = [int(g) for g in start_loss]
            start_loss = ['{:,}'.format(g) for g in start_loss]
            if len(start) < 3 or len(start) == 3:
                pp = ['Q1', 'Q2', 'Q3', 'Q4']
                ws.Range('C5:F5').Value = period(pp, 5)
                ws.Range('F3').Value = 'Currency:NT$'
                xl.Range('F3').HorizontalAlignment = xlRight
                ws.Range('F4').Value = this_quarter
                xl.Range('F4').HorizontalAlignment = xlRight
                ws.Range(months(start, '6', 'C')).Value = start
                ws.Range(months(start, '7', 'C')).Value = "{0:.2%}".format(k / 100)
            else:
                ws.Range(months(start, '5', 'C')).Value = period(start, 6)
                ws.Range(months(start, '3', 'N')).Value = 'Currency:NT$'
                xl.Range(months(start, '3', 'N')).HorizontalAlignment = xlRight
                ws.Range(months(start, '4', 'N')).Value = this_quarter
                xl.Range(months(start, '4', 'N')).HorizontalAlignment = xlRight
                xl.Range(months(start, '3', 'N')).HorizontalAlignment = xlRight
                xl.Range(months(start, '5', 'N')).HorizontalAlignment = xlRight
                for cell in ws.Range(months(start, '5', 'C')):
                    cell.BorderAround(1, 3)
                    cell.Borders(xlEdgeLeft).LineStyle = 0
                    cell.Borders(xlEdgeLeft).LineStyle = 1
                ws.Range(months(start, '6', 'C')).Value = start
                for cell in ws.Range(months(start, '6', 'C')):
                    cell.BorderAround(1, 2)
                ws.Range(months(start, '7', 'C')).Value = "{0:.2%}".format(k/100)
                for cell in ws.Range(months(start, '7', 'C')):
                    cell.BorderAround(1, 2)

                for cell in ws.Range(months(start, '8', 'C')):
                    cell.BorderAround(1, 2)
                for cell in ws.Range(months(start, '9', 'C')):
                    cell.BorderAround(1, 2)
                for cell in ws.Range(months(start, '10', 'C')):
                    cell.BorderAround(1, 3)
                    cell.Borders(xlEdgeTop).LineStyle = 0
                    cell.Borders(xlEdgeLeft).LineStyle = 0
                    cell.Borders(xlEdgeLeft).LineStyle = 1
                    cell.Borders(xlEdgeTop).LineStyle = 1
                for cell in ws.Range('B6:B9'):
                    cell.Borders(xlEdgeBottom).LineStyle = 0
                    cell.Borders(xlEdgeBottom).LineStyle = 1
                last_col = (months(start, '3', 'C'))[3] + str(5) + ":" + (months(start, '3', 'C'))[3] + str(10)
                ws.Range(last_col).BorderAround(1, 3)
                ws.Range(last_col).Borders(xlEdgeLeft).LineStyle = 0
                ws.Range(last_col).Borders(xlEdgeLeft).LineStyle = 1

            ws.Range(months(start, '8', 'C')).Value = start_loss
            if sum([float(i.replace(',', '')) for i in start_loss]) != 0:
                for y in ws.Range(months(start, '8', 'C')).Value:
                    yy = y
                    yy = list(yy)
                    yy = [0 if g is None else g for g in yy]
                    yy = [int(g) for g in yy]
            else:
                yy = [float(i) for i in start_loss]
            # if sum([float(i.replace(',', '')) for i in start]) != 0:
            #     for x in ws.Range(months(start, '6', 'C')).Value:
            #         xx = x
            #         xx = list(xx)
            #         xx = [0 if g is None else g for g in xx]
            #         xx = [int(g) for g in xx]
            # else:
            xx = [float(i.replace(',', '')) for i in start]
            ws.Range(months(start, '9', 'C')).Value = 0
            try:
                ws.Range(months(start, '10', 'C')).Value = ["{:.2%}".format(s/t) for s, t in zip(yy, xx)]
            except ZeroDivisionError:
                pass
            break

    xl.Rows(5).HorizontalAlignment = xlCenter
    ws.Range('A:ZZ').Font.Name = 'Microsoft JhengHei UI'
    ws.Range('A5:ZZ5').Font.Bold = True
    xl.Rows(5).VerticalAlignment = 2
    ws.Range('A:ZZ').Font.Size = 11
    fn = str(j) + i + ".xlsx"
    ws.SaveAs(Filename=os.path.join(save_path, fn))
    wb.Close()
    xl.Quit()

df_os_result = df_os.merge(result, on='number')
for i, j, k in zip(df_os_result['合約名稱_x'], df_os_result['合約年_x'], range(len(df_os_result))):
    head_file = save_path + '\\' + str(j) + i + '.xlsx'
    print(i)
    if os.path.exists(head_file):
        wb2 = xl.Workbooks.Open(head_file)
        ws2 = wb2.WorkSheets('Trangle -1')
        a = df_os_result.loc[k, quarter_os]
        a = a.tolist()
        b = df_os_result.loc[k, quarter]
        b = b.tolist()
        for x, y in enumerate(b):
            if y != 0:
                start_outstanding = a[x::]
                print(start_outstanding)
                start_outstanding = ['{:,}'.format(g) for g in start_outstanding]
                ws2.Range(months(start_outstanding, '9', 'C')).Value = start_outstanding
                break
        try:
            for x in ws2.Range(months(start_outstanding, '9', 'C')).Value:
                xx = x
                xx = list(xx)
                xx = [0 if g is None else g for g in xx]
                xx = [int(g) for g in xx]
            for y in ws2.Range(months(start_outstanding, '8', 'C')).Value:
                yy = y
                yy = list(yy)
                yy = [0 if g is None else g for g in yy]
                yy = [int(g) for g in yy]
            for z in ws2.Range(months(start_outstanding, '6', 'C')).Value:
                zz = z
                zz = list(zz)
                zz = [0 if h is None else h for h in zz]
                zz = [int(g) for g in zz]
        except TypeError:
            pass
        total = [v+s for v, s in zip(xx, yy)]
        total = [0 if v is None else v for v in total]
        try:
            sum_paid = [k/l for k, l in zip(total, zz)]
        except ZeroDivisionError:
            sum_paid = [0]
        ws2.Range(months(start_outstanding, '10', 'C')).Value = ["{:.2%}".format(i) for i in sum_paid]

        fn = str(j) + i + ".xlsx"
        wb2.SaveAs(Filename=os.path.join(save_path, fn))
        wb2.Close()
        xl.Quit()
    else:
        pass

#有renewal
all_account = []
for root, dirs, files in walk('D:\\py\\head'):
    for f in files:
        fullpath = join(root, f)
        all_account.append(fullpath)
count = 1
group = []
save_path_head = 'D:\\py\\head\\final'


# def renewal():
#     for a in all_account:
#         group = all_account[0]
#         if i[18::] == group[0][18::]:
#             group.append(i)
#             all_account.remove(a)
#             return group
#
# ans = renewal()
# if len(ans) > 0:
#     for i in ans[1::]:
#             wb = xl.Workbooks.Open(group[0])
#             ws = wb.WorkSheets('Trangle -1')
#             wb1 = xl.Workbooks.Open(group[i])
#             ws1 = wb1.WorkSheets('Trangle -1')
#             ws.Range('6:10').Copy(ws.Range('%s:%s') % (6+n*5, 10+n*5))
#             ws.Range('%s:%s' % (6+n*5, 10+n*5)).Value = ' '
#             ws.Range('%s:%s' % (6+n*5, 10+n*5)).Value = ws1.Range('6:10').Value
#             count += 1
#             group.remove(group[i])
#             # wb.SaveAs(Filename=os.path.join(save_path_head, all_account[0][15::]))
#             wb.Save()
#             wb.Close()
#             wb1.Close()
# else:
#     pass
#
#
##聯結續保件
all_account = []
for root, dirs, files in walk('D:\\py\\head'):
    for f in files:
        fullpath = join(root, f)
        all_account.append(fullpath)

all_account = sorted(all_account, key=lambda g:(g[15::], g))
group = [list(g) for _, g in itertools.groupby(all_account, lambda x: x[15::])]
group = [g for g in group if len(g) >= 2]
print(group)

for g in group:
    n = 1
    for s in g[1::]:
        print(g[0])
        wb = xl.Workbooks.Open(g[0])
        ws = wb.WorkSheets('Trangle -1')
        wb1 = xl.Workbooks.Open(s)
        ws1 = wb1.WorkSheets('Trangle -1')
        ws.Range('6:10').Copy(ws.Range('%s:%s' % (6 + n * 5, 10 + n * 5)))
        ws.Range('%s:%s' % (6 + n * 5, 10 + n * 5)).Value = ''
        ws.Range('%s:%s' % (6 + n * 5, 10 + n * 5)).Value = ws1.Range('6:10').Value
        ws.Range('%s:%s' % (11 + n * 5, 16 + n * 5)).Value = ws1.Range('11:16').Value
        n = n+1
        m = str(n) + '.xlsx'
        wb.Save()
        # ws.SaveAs(Filename=os.path.join('D:\\py\\head', m))
        wb1.Save()
        wb.Close(True)
        wb1.Close(True)
