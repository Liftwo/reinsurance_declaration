import os
import shutil
import pandas as pd
from os import walk

list_re = pd.read_excel('D:\\團傷臨分明細.xlsx', sheet_name = '出單')

goal = 'D:\\py\\final'
for root, dirs, files in walk(goal):
    for f in files:
        for i, j, k in zip(list_re['分保合約編號'], list_re['合約年'], list_re['合約名稱']):
            if f[0:12] == i:
                shutil.move(os.path.join(goal, f), 'D:\\py\\rename_CR&SOA\\'+str(j)+k+'.pdf')
