# 指定した文言を含むシートを持つexcelファイルを抽出する
# --使い方--
# 「py SearchSheet.py rootDirPath searchWord」
# rootDirPath : 検索する元となるフォルダ
# searchWord : 検索文字列 ※指定しない場合は全てのシートを出力する

import glob
import openpyxl
import sys

# 第一引数を評価
if len(sys.argv) > 1:
    # xlsxファイルを再帰的に取得
    files = glob.glob(sys.argv[1] + "\**\*.xlsx", recursive=True)
    # xlsmファイルを再帰的に取得し、リストを結合する
    files.extend(glob.glob(sys.argv[1] + "\**\*.xlsm", recursive=True))
else:
    print("warning : 第一引数にフォルダパスを指定してください。")
    exit()

for file in files:
    try:
        wb = openpyxl.load_workbook(file)
    except PermissionError:
        print(file + "\t" + "error : 権限がありません。")
        continue

    sheets = wb.sheetnames
    for sheet in sheets:
        # 第2引数が含まれている場合は部分一致検索を行う
        if len(sys.argv) > 2:
            if sys.argv[2] in sheet:
                print(file + "\t" + sheet)
        else:
            print(file + "\t" + sheet)
