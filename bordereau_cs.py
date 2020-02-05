import pandas as pd
import numpy as np

bordereau_cs = pd.read_excel('M:\季帳\\2019Q4\marsh\工作檔_2019Q4boardereaux_(公務人員等).xls', sheet_name='公務人員')

xls = pd.ExcelFile('M:\\保單、名冊\\2019被保人名冊.xlsx')
matchers = ['公務', '教師', '教育']
sel = [s for s in xls.sheet_names if any(xs in s for xs in matchers)]

goal_sheet = []
for sheet in sel:
    goal_sheet.append(xls.parse(sheet))

name_list_all = pd.concat(goal_sheet, sort=True)
writer = pd.ExcelWriter('D:\\namename.xlsx')
name_list_all.to_excel(writer, index=False)
writer.save()


def merge(bordereau, name_list):
    def new_column(row):
        if len(str(row['被保險人'])) > 3:
            return str(row['被保險人'])[0:3] + 'O'
        elif len(str(row['被保險人'])) == 2:
            return str(row['被保險人'])[0:1] + 'O'
        else:
            return str(row['被保險人'])[0:2] + 'O'

    bordereau['ID2'] = bordereau.apply(lambda row:'OOOOOOO'+str(row['ID'])[-3:], axis=1)
    bordereau['被保險人2'] = bordereau.apply(new_column, 1)
    writer_boarde = pd.ExcelWriter('D:\\boar.xlsx')
    bordereau.to_excel(writer_boarde, index=False)
    writer_boarde.save()
    name_list = name_list.rename(columns={'被保險人':'被保險人2', '身分證號碼':'ID2'})
    writer_name_list = pd.ExcelWriter('D:\\name_list.xlsx')
    name_list.to_excel(writer_name_list, index=False)
    writer_name_list.save()
    result = bordereau.merge(name_list, on=['被保險人2', 'ID2'], how='left')
    result = result.drop_duplicates(subset=['合計', '給付項目', '事故日', '被保險人'], keep='first')
    # result = pd.merge(bordereau, name_list, how='left', on=['被保險人2','ID2'])
    return result


def pivotable(a):
    table = pd.pivot_table(a, values='應攤總額', index=['被保險人', '賠案號碼', '保單號碼', '選擇方案', '保單生效日', '保單到期日', '事故日', '診斷名稱'], aggfunc=np.sum)
    return table


servant = merge(bordereau_cs, name_list_all)
writer_servant = pd.ExcelWriter('D:\\servant.xlsx')
servant.to_excel(writer_servant)
writer_servant.save()