# -*- coding: utf-8 -*-

#https://qiita.com/taashi/items/07bf75201a074e208ae5

import os
import sys
import openpyxl
from googletrans import Translator

translator = Translator()

args = sys.argv

# Excel読み込み
path_in=args[1]
tmpf=os.path.splitext(path_in)
path_out=tmpf[0] + '_trans' +tmpf[1]
print(path_in)

wb = openpyxl.load_workbook(path_in)
sheet = wb.worksheets[0]

# 1行ごとに翻訳
for row in range(1, sheet.max_row + 1):
    en = sheet.cell(row=row, column=1).value
    if en!=None:
        ja = translator.translate(text=en, src="en", dest="ja")
        print(en)
        print("  =>",ja.text)

        sheet.cell(row=row, column=2).value = ja.text
        sheet.cell(row=row, column=2).alignment = openpyxl.styles.Alignment(wrapText=True)

wb.save(path_out)
