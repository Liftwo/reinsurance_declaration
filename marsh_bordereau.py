import pandas as pd
import numpy as np
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import xlrd

bordereau_all = pd.read_excel('M:\季帳\\2019Q4\marsh\工作檔_2019Q4boardereaux_(公務人員等).xls', sheet_name='公務人員')

bordereau_cs = bordereau_all[bordereau_all['分保合約編號'] == '8HFB20190Q22']
bordereau_ocean1 = bordereau_all[bordereau_all['分保合約編號'] == '8HFB20190Q12']
bordereau_ocean2 = bordereau_all[bordereau_all['分保合約編號'] == '8HFB20190Q55']


sel = []
xls = pd.ExcelFile('M:\\保單、名冊\\2019被保人名冊.xlsx')
for i in xls.sheet_names:
    if '公務人員' in i or '工會' in i:
        sel.append(i)
sel.remove('公務人員共4件')
sel.remove('桃園市桃園市金屬建築結構及組件製造職業工會')
print(sel)
goal_sheet = []
for sheet in sel:
    goal_sheet.append(xls.parse(sheet))

name_list_all = pd.concat(goal_sheet, sort=True)
writer = pd.ExcelWriter('D:\\namename.xlsx')
name_list_all.to_excel(writer, index=False)
writer.save()

# name_list_ocean1 = pd.read_excel('M:\\保單、名冊\\2019被保人名冊.xlsx', '岸巡')
# writer = pd.ExcelWriter('D:\\namename_nameocean1.xlsx')
# name_list_all.to_excel(writer, index=False)
# writer.save()
# name_list_ocean2 = pd.read_excel('M:\\保單、名冊\\2019被保人名冊.xlsx', '洋巡')
# writer = pd.ExcelWriter('D:\\namename_ocean2.xlsx')
# name_list_all.to_excel(writer, index=False)
# writer.save()


def merge(bordereau, name_list):
    def new_column(row):
        if len(str(row['被保險人'])) > 3:
            return str(row['被保險人'])[0:3] + 'O'
        return str(row['被保險人'][0:2]) + 'O'

    bordereau['ID2'] = bordereau.apply(lambda row:'OOOOOOO'+str(row['ID'])[-3:], axis=1)
    bordereau['被保險人2'] = bordereau.apply(new_column, 1)

    name_list = name_list.rename(columns={'被保險人':'被保險人2', '身分證號碼':'ID2'})
    result = pd.merge(bordereau, name_list, how='left', on=['被保險人2','ID2'])
    return result


def pivotable(a):
    table = pd.pivot_table(a, values='應攤總額', index=['被保險人', '賠案號碼', '保單號碼', '選擇方案', '保單生效日', '保單到期日', '事故日', '診斷名稱'], aggfunc=np.sum)
    return table


servant = merge(bordereau_cs, name_list_all)
# ocean1 = merge(bordereau_ocean1, name_list_ocean1)
# ocean2 = merge(bordereau_ocean2, name_list_ocean2)


writer_servant = pd.ExcelWriter('D:\py\marsh_bordereau.xlsx')
# writer_ocean1 = pd.ExcelWriter('D:\py\marsh_bordereau.xlsx', sheet_name='ocean1')
# writer_ocean2 = pd.ExcelWriter('D:\py\marsh_bordereau.xlsx', sheet_name='ocean2')

pivotable(servant).to_excel(writer_servant)
writer_servant.save()
# pivotable(ocean1).to_excel(writer_ocean1, index=False)
# writer_servant.save(writer_ocean1)
# pivotable(ocean2).to_excel(writer_ocean2, index=False)
# writer_servant.save(writer_ocean2)

# wb = xlsxwriter.Workbook('D:\\2019Q2公務人員Bordereaux.xlsx')