import re

import pandas as pd
import os.path

import datetime

import pdfplumber
from PyPDF2 import PdfFileReader, PdfFileWriter
from win32com.client import Dispatch
from datetime import datetime
import fitz
from os import walk
from os.path import join
import comtypes.client
import decimal
from decimal import *

xl = Dispatch("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
which_quarter = input('西元+季:')
as_at = input('e.g.Sep 30,2019')
soa_period = input('e.g.201909N1')
AC_period = input('e.g. Apr.01, 2019 to Jun.30, 2019')
save_path = 'M:\\季帳\\{}\\starr\\CR'.format(which_quarter)
wbs_path = 'D:\\creditnote.xlsx'

df = pd.read_excel('M:\\季帳\\{}\\{}樞紐-1510(2020合約要入再Q2).xlsx'.format(which_quarter, which_quarter), sheet_name='原始資料')
df2 = pd.read_excel('D:\\團傷臨分明細.xlsx', sheet_name='出單')

df_1 = df.groupby(['分保合約編號', '種類'])['分出金額'].sum().reset_index()

result = df_1.merge(df2, on='分保合約編號')
result = result[result['保險期間'] != 'N']
result = result[pd.notnull(result['保險期間'])]
result[['起始日', '結束日']] = result['保險期間'].str.split(pat='-', expand=True)
result['起始日'] = result['起始日'].astype(str)
result['結束日'] = result['結束日'].astype(str)
result['起始日'] = result['起始日'].apply(lambda x: datetime.strptime(x, '%Y%m%d').strftime('%b.%d, %Y'))
result['結束日'] = result['結束日'].apply(lambda x: datetime.strptime(x, '%Y%m%d').strftime('%b.%d, %Y'))


def point_to_int(col_name):
    for i in col_name:
        if round(i) == int(i) and isinstance(i, float):
            if i > 0:
                i = i + 0.5
            else:
                i = i - 0.5
        else:
            i = round(i)
        return i


# result['分出金額'] = result.apply(lambda x: point_to_int(result['分出金額']), axis=1)
result['分出金額'] = result['分出金額'].apply(lambda x: Decimal(x).quantize(Decimal('0'), rounding=decimal.ROUND_HALF_UP))
result['分出金額'] = result['分出金額'].astype(int)
result = result.sort_values(by=['種類'])
result['佣金率'] = result['佣金率'].astype(float).map("{:.0%}".format)
result['佣金率(%)'] = result['佣金率'].str.rstrip('%').astype('float') / 100.0
result_p = result[result['種類'] == '保費支出']
result_l = result[result['種類'] == '攤回賠款']
writer = pd.ExcelWriter('D:\\result.xlsx')
writer_p = pd.ExcelWriter('D:\\result_p.xlsx')
writer_l = pd.ExcelWriter('D:\\result_l.xlsx')
result.to_excel(writer, index=False)
result_p.to_excel(writer_p, index=False)
result_l.to_excel(writer_l, index=False)
writer.save()
writer_p.save()
writer_l.save()

for i, j, k, l, m, n, o, p in zip(result_p['分保合約編號'], result_p['合約名稱'], result_p['起始日'],
                                  result_p['佣金率'], result_p['種類'], result_p['分出金額'], result_p['佣金率(%)'],
                                  result_p['結束日']):

    wb1 = xl.Workbooks.Open('D:\\creditnote.xlsx')
    ws1 = wb1.WorkSheets('ABC')
    ws1.Range('B8').Value = "Our Reference : " + i + "-{}".format(soa_period)
    ws1.Range('B7').Value = "Insured : " + j
    ws1.Range('B9').Value = "Insurance Period : " + k + " to " + p
    ws1.Range('X6').Value = "As at {}".format(as_at)
    ws1.Range('B10').Value = "A/C Period : {}".format(AC_period)
    ws1.Range('B15').Value = "R/I Commission" + "(" + l + ")"
    ws1.Range('X14').Value = n
    ws1.Range('J15').Value = n * o
    if round(ws1.Range('J15').Value) == int(ws1.Range('J15').Value):
        if ws1.Range('J15').Value > 0:
            ws1.Range('J15').Value = int(ws1.Range('J15').Value + 0.5)
        else:
            ws1.Range('J15').Value = int(ws1.Range('J15').Value - 0.5)
    else:
        ws1.Range('J15').Value = int(round(n * o))

    ws1.Range('J16').Value = (ws1.Range('X14').Value) * 0.01
    if round(ws1.Range('J16').Value) == int(ws1.Range('J16').Value):
        if ws1.Range('J16').Value > 0:
            ws1.Range('J16').Value = int(ws1.Range('J16').Value + 0.5)
        else:
            ws1.Range('J16').Value = int(ws1.Range('J16').Value - 0.5)
    else:
        ws1.Range('J16').Value = int(round(ws1.Range('J16').Value))

    ws1.Range('X19').Value = n

    fn = i + ".xlsx"
    ws1.SaveAs(Filename=os.path.join(save_path, fn))
    wb1.Close()
    xl.Quit()

for i, j, k, l, m, n, o in zip(result_l['分保合約編號'], result_l['合約名稱'], result_l['起始日'],
                               result_l['佣金率'], result_l['種類'], result_l['分出金額'], result_l['結束日']):
    CR_file = save_path + '\\' + i + '.xlsx'
    if os.path.exists(CR_file):
        wb2 = xl.Workbooks.Open(CR_file)
        ws2 = wb2.WorkSheets('ABC')
        ws2.Range('J14').Value = n
        ws2.Range('J17').Value = int(ws2.Range('X14').Value) - int(ws2.Range('J14').Value) - int(
            ws2.Range('J15').Value) - int(ws2.Range('J16').Value)
        ws2.Range('J19').Value = int(ws2.Range('J14').Value) + int(ws2.Range('J15').Value) + int(
            ws2.Range('J16').Value) + int(ws2.Range('J17').Value)
        fn = i + ".xlsx"
        wb2.SaveAs(Filename=os.path.join(save_path, fn))
        wb2.Close(True)

    else:
        wb2 = xl.Workbooks.Open(wbs_path)
        ws2 = wb2.WorkSheets('ABC')
        ws2.Range('B8').Value = "Our Reference : " + i + "-{}".format(soa_period)
        ws2.Range('B7').Value = "Insured : " + j
        ws2.Range('B9').Value = "Insurance Period : " + k + " to " + o
        ws2.Range('X6').Value = "As at {}".format(as_at)
        ws2.Range('B10').Value = "A/C Period : {}".format(AC_period)
        ws2.Range('B15').Value = "R/I Commission" + "(" + l + ")"
        ws2.Range('J14').Value = n
        ws2.Range('X14').Value = 0
        ws2.Range('J15').Value = 0
        ws2.Range('J16').Value = 0
        ws2.Range('J17').Value = int(ws2.Range('X14').Value) - int(ws2.Range('J14').Value) - int(
            ws2.Range('J15').Value) - int(ws2.Range('J16').Value)
        ws2.Range('J19').Value = int(ws2.Range('J14').Value) + int(ws2.Range('J15').Value) + int(
            ws2.Range('J16').Value) + int(ws2.Range('J17').Value)

        fn = i + ".xlsx"
        ws2.SaveAs(Filename=os.path.join(save_path, fn))

        wb2.Close()
        xl.Quit()

def excel_pdf(path):
    pdf_path = path.replace('xlsx', 'pdf')
    xlApp = comtypes.client.CreateObject("Excel.Application")
    books = xlApp.Workbooks.open(path)
    books.ExportAsFixedFormat(0, pdf_path)
    xlApp.Quit()


for root, dirs, files in walk(save_path):
    for f in files:
        fullpath = join(root, f)
        excel_pdf(fullpath)

for root, dirs, files in walk(save_path):
    for f in files:
        if f.endswith('.pdf'):
            input_file = join(root, f)
            output_file = join('D:\\py\\new_CR', f)
            barcode_mark = 'D:\\py\\title.png'
            barcode_sign = 'D:\\sign2.png'
            doc = fitz.open(input_file)
            page = doc[0]
            rect_mark = fitz.Rect(0, 0, 1024, 50)
            rect_sign = fitz.Rect(250, 545, 800, 1000)
            pix_1 = fitz.Pixmap(barcode_mark)
            pix_2 = fitz.Pixmap(barcode_sign)
            page.insertImage(rect_sign, pixmap=pix_2, overlay=True)
            page.insertImage(rect_mark, pixmap=pix_1, overlay=True)
            doc.save(output_file)

import fitz
for root, dirs, files in walk('D:\\py\\new_CR'):
    for f in files:
        doc = fitz.open(join(root, f))
        width, height = fitz.PaperSize('a4')
        totaling = doc.pageCount
        for pg in range(totaling):
            page = doc[pg]
            zoom = int(100)
            rotate = int(0)
            trans = fitz.Matrix(zoom / 60, zoom / 60).preRotate(rotate)
            pm = page.getPixmap(matrix=trans, alpha=False)
            lurl = 'D:\\py\\new_CR\\{}.jpg'.format(str(f)[0:12])
            pm.writePNG(lurl)
        doc.close()

for root, dirs, files in walk('D:\\py\\new_CR'):
    for f in files:
        if '.jpg' in f:
            doc_pdf = fitz.open()
            imgdoc = fitz.open(join(root, f))
            pdfbytes = imgdoc.convertToPDF()
            imgpdf = fitz.open('pdf', pdfbytes)
            doc_pdf.insertPDF(imgpdf)
            doc_pdf.save('D:\\py\\new_CR\\final\\{}.pdf'.format(str(f)[0:12]))
            doc_pdf.close()

inputpdf = PdfFileReader(open('D:\\py\\8H.pdf', 'rb'))
for i in range(inputpdf.numPages):
    output = PdfFileWriter()
    output.addPage(inputpdf.getPage(i))
    with open('D:\\py\\soa\\page%s.pdf' % i, 'wb') as outputStream:
        output.write(outputStream)
#
soa_path = 'D:\\py\\soa'
for root, dirs, files in walk(soa_path):
    for f in files:
        fullpath = join(root, f)
        pdf_file = pdfplumber.open(fullpath)
        page = pdf_file.pages[0]
        text = page.extract_text()
        position = str(text).find('OUR')
        soa_number = str(text)[position + 52:position + 64] + '.pdf'
        pdf_file.close()
        os.rename(fullpath, join(root, soa_number))

cr_path = 'D:\\py\\new_CR\\final'
merge_CR = []
for root, dirs, files in walk(cr_path):
    for f in files:
        if '.pdf' in f:
            merge_CR.append(join(root, f))

merge_soa = []
for root, dirs, files in walk('D:\\py\\soa'):
    for f in files:
        merge_soa.append(join(root, f))


def merger_pdf(output_path, input_paths):
    pdf_writer = PdfFileWriter()
    for path in input_paths:
        pdf_reader = PdfFileReader(path)
        for page in range(pdf_reader.getNumPages()):
            pdf_writer.addPage(pdf_reader.getPage(page))
            with open(output_path, 'wb') as fh:
                pdf_writer.write(fh)


for pdf_cr in merge_CR:
    for pdf_soa in merge_soa:
        if pdf_cr[19::] == pdf_soa[10::]:
            merge_final = [pdf_cr, pdf_soa]
            merger_pdf(join('D:\\py\\final', pdf_soa[10::]), merge_final)
        else:
            pass

# check CR的金額
balance_amount = []
account = []
for root, dirs, files in walk(save_path):
    for f in files:
        if 'pdf' in f:
            pdffile = join(root, f)
            pdf = pdfplumber.open(pdffile)
            page = pdf.pages[0]
            text = page.extract_text()
            start = text.find('Balance Due to You')
            end = text.find('Total')
            pdf.close()
            balance_amount.append(text[start + 19:end - 1])
            account.append(f[0:12])

account_balance = list(zip(account, balance_amount))
df_balance = pd.DataFrame(account_balance, columns=['account', 'amount_CR'])
df_balance['amount_CR'] = df_balance['amount_CR'].apply(lambda x: x.replace(',', ''))
df_balance['amount_CR'] = df_balance['amount_CR'].astype(str).str.replace('\((.*)\)', '\\1')
df_balance['amount_CR'] = df_balance['amount_CR'].astype(int)
writer_balance = pd.ExcelWriter('D:\\py\\CR_balance.xlsx')
df_balance.to_excel(writer_balance, index=False)
writer_balance.save()

# check SOA的金額
account_soa = []
soa_balance = []
for root, dirs, files in walk(soa_path):
    for f in files:
        pdffile_soa = join(root, f)
        pdf_soa = pdfplumber.open(pdffile_soa)
        page_soa = pdf_soa.pages[0]
        text_soa = page_soa.extract_text()
        duetoyou = text_soa.find('DUE TO YOU')
        duetous = text_soa.find('PPPPll')
        pdf_soa.close()
        account_soa.append(f[0:12])
        soa_balance.append(re.findall(r'[,\d]+.?\d*', text_soa[duetoyou:duetous]))

soa_balance2 = []
for i in soa_balance:
    for j in i:
        soa_balance2.append(j)

account_soa_balance = list(zip(account_soa, soa_balance2))
df_soa_balance = pd.DataFrame(account_soa_balance, columns=['account', 'amount_soa'])
df_soa_balance['amount_soa'] = df_soa_balance['amount_soa'].apply(lambda x: x.replace(',', ''))
df_soa_balance['amount_soa'] = pd.to_numeric(df_soa_balance['amount_soa'])
writer_soa_balance = pd.ExcelWriter('D:\\py\\soa_balance.xlsx')
df_soa_balance.to_excel(writer_soa_balance, index=False)
writer_soa_balance.save()

# 合併CR跟SOA做比對
df_confirm_amount = df_balance.merge(df_soa_balance, on='account')
writer_confirm_amount = pd.ExcelWriter('D:\\py\\confirm_amount.xlsx')
df_confirm_amount.to_excel(writer_confirm_amount, index=False)
writer_confirm_amount.save()

# 針對尾數差檔案重新跑
