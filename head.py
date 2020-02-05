import pandas as pd
import os.path
from win32com.client import Dispatch

xl = Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
save_path = 'M:\\季帳\\2019Q2\\表頭'
wbs_path = 'D:\表頭.xls'

df = pd.read_excel('M:\\季帳\\2019Q2\\X2019Q2樞紐-1510(華南賠款保費分開).xlsx', sheet_name='原始資料')
df2 = pd.read_excel('D:\\團傷臨分明細.xlsx', sheet_name='出單')


df_1 = df.groupby(['分保合約編號','種類'])['總金額'].sum().reset_index()

result = df_1.merge(df2, on='分保合約編號')

result['佣金率'] = result['佣金率'].astype(float).map("{:.0%}".format)
result['分保比例'] = result['分保比例'].astype(float).map("{:.0%}".format)

result_p = result[result['種類'] == '保費支出']
result_l = result[result['種類'] == '攤回賠款']
writer = pd.ExcelWriter('D:\\hresult.xlsx')
writer_p = pd.ExcelWriter('D:\\hresult_p.xlsx')
writer_l = pd.ExcelWriter('D:\\hresult_l.xlsx')
result.to_excel(writer, index=False)
result_p.to_excel(writer_p, index=False)
result_l.to_excel(writer_l, index=False)
writer.save()
writer_p.save()
writer_l.save()

for i, j, k, l, m in zip(result_p['合約名稱'], result_p['總金額'], result_p['佣金率'], result_p['分保比例'], result_p['合約年']):
    wb1 = xl.Workbooks.Open(wbs_path)
    ws1 = wb1.WorkSheets('Trangle -1')
    ws1.Range('A4').Value = i
    ws1.Range('A6').Value = m
    ws1.Range('C6').Value = j
    ws1.Range('C7').Value = k
    ws1.Range('C8').Value = 0
    ws1.Range('C9').Value = 0
    ws1.Range('C10').Value = '0%'
    ws1.Range('A11').Value = "Remark:RI" + " " + l
    ws1.Range('F4').Value = 'For Ending Jun. 2019'

    fn = i + ".xls"
    ws1.SaveAs(Filename=os.path.join(save_path, fn))
    wb1.Close()
    xl.Quit()

for i, j, k in zip(result_l['總金額'], result_l['佣金率'], result_l['合約名稱']):
    CR_file = save_path + '\\' + k + '.xls'
    print(CR_file)
    if os.path.exists(CR_file):
        wb2 = xl.Workbooks.Open(CR_file)
        ws2 = wb2.WorkSheets('Trangle -1')
        ws2.Range('C8').Value = i
        ws2.Range('C9').Value = 0
        ws2.Range('C10').Value = (int(ws2.Range('C8').Value) + int(ws2.Range('C9').Value)) / int(ws2.Range('C6'))

        fn = k + ".xls"
        ws2.SaveAs(Filename=os.path.join(save_path, fn))
        wb2.Close()
        xl.Quit()
    else:
        pass

