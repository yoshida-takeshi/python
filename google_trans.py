# -*- coding: utf-8 -*-

#https://qiita.com/taashi/items/07bf75201a074e208ae5

import sys
import openpyxl
from googletrans import Translator

translator = Translator()

args = sys.argv

# Excel読み込み
path=args[1]
print(path)

wb = openpyxl.load_workbook(path)
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

wb.save(path)
