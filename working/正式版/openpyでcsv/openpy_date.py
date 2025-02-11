import openpyxl
import csv
import time
from datetime import datetime

def read_excel_optimized(excel_file, sheet_name, csv_file):
    """Excelファイルを読み込み、指定されたシートのデータをCSVファイルに書き出す（高速化版）。"""

    try:
        wb = openpyxl.load_workbook(excel_file, read_only=True, data_only=True)
        sheet = wb[sheet_name]
    except FileNotFoundError:
        print(f"エラー: ファイル '{excel_file}' が見つかりません。")
        return
    except KeyError:
        print(f"エラー: シート '{sheet_name}' が見つかりません。")
        return
    except Exception as e:
        print(f"エラー: Excelファイルの読み込み中にエラーが発生しました: {e}")
        return

    try:
        with open(csv_file, mode='w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)

            # ヘッダーを一度だけ書き込む
            header = ["作業日", "氏名", "工事番号", "実績工数", "所属"]
            writer.writerow(header)

            # Excelファイルを読み込む（高速化）
            for row in sheet.iter_rows(min_row=2):  # 2行目から開始
                row_data = []
                for cell in row[:5]:  # 必要な列のみ取得
                    value = cell.value
                    if isinstance(value, datetime):
                        value = value.strftime('%Y-%m-%d') # 日付型の場合、フォーマット変換
                    row_data.append(value if value is not None else "")
                writer.writerow(row_data)

        print(f"CSVファイル '{csv_file}' にデータを書き出しました。")

    except Exception as e:
        print(f"エラー: CSVファイルへの書き込み中にエラーが発生しました: {e}")
        return

# 使用例
# excel_file = input("エクセルファイル名を入力してください: ")
# sheet_name = input("シート名を入力してください: ")
excel_file = "input_file_large.xlsx"
sheet_name = "Sheet1"

# csv_file = excel_file.replace('.xlsx', '.csv')
csv_file = 'output.csv'

time_start = time.time()
print("<処理開始>")

read_excel_optimized(excel_file, sheet_name, csv_file)

time_end = time.time() - time_start
print("<処理時間> ", time_end)
