from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
import os
import pandas as pd

newlist = []
path1='M:\\季帳\\2019Q2\\starr\\statement\\'
path2='M:\季帳\\2019Q2\starr\CR\\'
save_path='C:\\Users\\3240\\Desktop\\CR\\result\\'

for i in os.listdir(path1):
    for j in os.listdir(path2):
        if i==j:
            newlist.append([i, j])

for i in newlist:
    i[0] = path1+i[0]
    i[1] = path2+i[1]

for i in newlist:
    merger = PdfFileMerger()
    for j in i:
        if not os.path.exists(save_path+'\\'+j):

            merger.append(PdfFileReader(j))
            merger.write(save_path+i[0][-12:])
    merger.close()


for i in os.listdir('C:\\Users\\3240\\Desktop\\CR\\result'):
    if i not in os.listdir('M:\季帳\\2019Q2\starr\CR&Statement'):
        with open('C:\\Users\\3240\\Desktop\\CR\\result\\'+i, 'rb') as file:
            reader = PdfFileReader(file)
            writer = PdfFileWriter()
            try:
                for p in range(1, 3):
                    writer.addPage(reader.getPage(p))
            except IndexError:
                pass

            with open('C:\\Users\\3240\\Desktop\\CR\\'+i, 'wb') as outfile:
                writer.write(outfile)
    else:
        pass


df = pd.read_excel('D:\\團傷臨分明細.xlsx', sheet_name = '出單')
for filename in os.listdir('M:\季帳\\2019Q2\starr\CR&Statement'):
    os.rename('M:\季帳\\2019Q2\starr\CR&Statement\\'+filename, 'M:\季帳\\2019Q2\starr\CR&Statement\\'+'8HFB'+filename[-12::])

import pandas as pd
df = pd.read_excel('D:\\團傷臨分明細.xlsx', sheet_name = '出單')

for filename in os.listdir('M:\季帳\\2019Q2\starr\CR&Statement'):
    for i,j in zip(df['分保合約編號'], df['合約名稱']):
        if filename[0:12]==i:
            os.rename('M:\季帳\\2019Q2\starr\CR&Statement\\'+filename, 'M:\季帳\\2019Q2\starr\CR&Statement\\'+j+'.pdf')









