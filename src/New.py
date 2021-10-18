# 空のExcelファイルを作成する
# --使い方--
# ・「New.py hoge」で実行
# ・hoge.xlsxがカレントディレクトリに作成される。
# ・既に同名ファイルが存在していた場合、ファイルは作成されない。

import sys
import openpyxl
import os

# コマンドラインで引数が渡されているかを検査
if len(sys.argv) > 1:
    fileName = sys.argv[1] + '.xlsx'
    # 既に同名ファイルが存在していた場合、ファイルは作成されない。
    if os.path.exists(fileName):
        print('既に[' + fileName + ']は存在している為、作成できませんでした。')
    else:
        wb = openpyxl.Workbook()
        wb.save(fileName)
        print(fileName + 'を作成しました。')
else:
    # 引数の指定がない
    print('ファイル名を指定して下さい。')
