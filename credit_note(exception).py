import os
from os.path import join, walk
import fitz
import pdfplumber
from PyPDF2 import PdfFileReader, PdfFileWriter

save_path = 'M:\\季帳\\2019Q4\\starr\\CR\\exception'
for root, dirs, files in walk(save_path):
    for f in files:
        if f.endswith('.pdf'):
            input_file = join(root, f)
            output_file = join('D:\\py\\new_CR\\exception', f)
            barcode_mark = 'D:\\py\\title.png'
            doc = fitz.open(input_file)
            page = doc[0]
            rect_mark = fitz.Rect(0, 0, 1024, 50)
            pix_1 = fitz.Pixmap(barcode_mark)
            page.insertImage(rect_mark, pixmap=pix_1, overlay=True)
            doc.save(output_file)

for root, dirs, files in walk('D:\\py\\new_CR\\exception'):
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

for root, dirs, files in walk('D:\\py\\new_CR\\exception'):
    for f in files:
        if '.jpg' in f:
            doc_pdf = fitz.open()
            imgdoc = fitz.open(join(root, f))
            pdfbytes = imgdoc.convertToPDF()
            imgpdf = fitz.open('pdf', pdfbytes)
            doc_pdf.insertPDF(imgpdf)
            doc_pdf.save('D:\\py\\new_CR\\final\\{}.pdf'.format(str(f)[0:12]))
            doc_pdf.close()

soa_path = 'D:\\py\\soa'
cr_path = 'D:\\py\\new_CR\\final\\exception'
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