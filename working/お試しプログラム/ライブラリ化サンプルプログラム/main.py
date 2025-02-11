from xmlcsvfnc import read_excel_large
# from プログラム名 import 関数名
import time

# 実行
excel_file = "input_file_large.xlsx" # ファイル名を指定
sheet_name = "sheet1" # シート名を指定
csv_output_file = "output.csv"

time_start = time.time()
print("<処理開始>")

read_excel_large(excel_file, sheet_name, csv_output_file)
# プログラムの関数名

time_end = time.time() - time_start
print("<処理時間> ", time_end)
