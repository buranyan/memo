import openpyxl
import csv
from lxml import etree

def read_excel_lxml_optimized(excel_file, sheet_name, csv_file):
    """Excelファイルを読み込み、指定されたシートのデータをCSVファイルに書き出す（lxml最適化版）。"""

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
            writer.writerow(["作業日", "氏名", "工事番号", "実績工数", "所属"])

            # lxmlのparserを指定
            parser = etree.XMLParser(recover=True)

            # ExcelファイルをXML形式で読み込む
            for row in sheet.rows:
                row_data = []
                for cell in row[:5]:  # 必要な列のみ取得
                    row_data.append(cell.value if cell.value is not None else "")

                writer.writerow(row_data)

        print(f"CSVファイル '{csv_file}' にデータを書き出しました。")

    except Exception as e:
        print(f"エラー: CSVファイルへの書き込み中にエラーが発生しました: {e}")
        return

# 使用例
excel_file = input("エクセルファイル名を入力してください: ")
sheet_name = input("シート名を入力してください: ")
csv_file = excel_file.replace('.xlsx', '.csv')

read_excel_lxml_optimized(excel_file, sheet_name, csv_file)
