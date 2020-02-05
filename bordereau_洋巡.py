import pandas as pd
import numpy as np
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import xlrd

bordereau_ocean2 = pd.read_excel('M:\季帳\\2019Q4\marsh\工作檔_2019Q4boardereaux_(公務人員等).xls', sheet_name='洋巡')
# bordereau_ocean2 = bordereau_ocean2.droop_duplicates()
name_list_ocean2 = pd.read_excel('M:\保單、名冊\\2019被保人名冊.xlsx', sheet_name='洋巡update')

writer = pd.ExcelWriter('D:\\namename.xlsx')
name_list_ocean2.to_excel(writer, index=False)
writer.save()


def merge(bordereau, name_list):
    def new_column(row):
        if len(str(row['被保險人'])) > 3:
            return str(row['被保險人'])[0:3] + 'O'
        return str(row['被保險人'][0:2]) + 'O'

    bordereau['ID2'] = bordereau.apply(lambda row:'OOOOOOO'+str(row['ID'])[-3:], axis=1)
    bordereau['被保險人2'] = bordereau.apply(new_column, 1)
    writer_boarde = pd.ExcelWriter('D:\\boar.xlsx')
    bordereau.to_excel(writer_boarde, index=False)
    writer_boarde.save()
    name_list = name_list.rename(columns={'被保險人':'被保險人2', '身分證號碼':'ID2'})
    writer_name_list = pd.ExcelWriter('D:\\name_list.xlsx')
    name_list.to_excel(writer_name_list, index=False)
    writer_name_list.save()
    result = pd.merge(bordereau, name_list, how='left', on=['被保險人2','ID2'])
    return result


def pivotable(a):
    table = pd.pivot_table(a, values='應攤總額', index=['被保險人', '賠案號碼', '保單號碼', '選擇方案', '保單生效日', '保單到期日', '事故日', '診斷名稱'], aggfunc=np.sum)
    return table


ocean2 = merge(bordereau_ocean2, name_list_ocean2)
writer_ocean2 = pd.ExcelWriter('D:\py\marsh_bordereau_洋巡.xlsx')
# pivotable(ocean2).to_excel(writer_ocean2)
ocean2.to_excel(writer_ocean2, index=False)
writer_ocean2.save()

# final = bordereau_ocean2.merge(ocean2, on='賠案號碼')
# writer_final = pd.ExcelWriter('D:\\py\\bordereau_ocean2_final.xlsx')
# final.to_excel(writer_final, index=False)
# writer_final.save()