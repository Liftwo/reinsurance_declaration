from os import walk
from os.path import join
import fitz

original_path = 'M:\\季帳\\2019Q4\\starr\\CR'


for root, dirs, files in walk(original_path):
    for f in files:
        if f.endswith('.pdf'):
            input_file = join(root, f)
            output_file = join('D:\\py\\new_CR\\print', f)
            barcode_sign = 'D:\\sign5.png'
            doc = fitz.open(input_file)
            page = doc[0]
            rect_sign = fitz.Rect(300, 600, 570, 900)
            pix_1 = fitz.Pixmap(barcode_sign)
            page.insertImage(rect_sign, pixmap=pix_1, overlay=True)
            doc.save(output_file)

