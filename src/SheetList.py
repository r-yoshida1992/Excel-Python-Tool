# 引数で指定したディレクトリ配下のエクセルファイルからシート名を全て出力する
# --使い方--
# ・「SheetList.py」で実行
# ・指定のディレクトリに存在するxlsxおよびxlsmファイルから、シート名を出力する。

import glob
import openpyxl
import sys

# 第一引数を評価
if len(sys.argv) > 1:
    excelFiles = glob.glob(sys.argv[1] + "\**\*.xlsx", recursive=True)
    excelFiles.extend(glob.glob(sys.argv[1] + "\**\*.xlsm", recursive=True))
else:
    print('warning : 第一引数にフォルダパスを指定してください。')
    exit()

for file in excelFiles:
    wb = openpyxl.load_workbook(file)
    sheets = wb.sheetnames
    for sheet in sheets:
        if len(sys.argv) > 2 :
            if sys.argv[2] in sheet :
                print(file + '\t' + sheet)
        else :
            print(file + '\t' + sheet)
