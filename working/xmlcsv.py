import zipfile
import csv
from lxml import etree
import time
from datetime import datetime, timedelta

def read_excel_large(file_path, output_csv):
    with zipfile.ZipFile(file_path, 'r') as zip_data:
        # ZIP内のファイルを確認
        file_list = zip_data.namelist()
        if 'xl/worksheets/sheet1.xml' not in file_list:
            raise FileNotFoundError("sheet1.xml が見つかりません")

        # `sharedStrings.xml` を読み込む（存在する場合のみ）
        shared_strings = []
        if 'xl/sharedStrings.xml' in file_list:
            shared_data = zip_data.read('xl/sharedStrings.xml')
            shared_root = etree.XML(shared_data)
            shared_strings = [s.text for s in shared_root.findall('.//{*}si/{*}t')]

        # `sheet1.xml` をストリーム処理で読み込む
        with zip_data.open('xl/worksheets/sheet1.xml') as f, open(output_csv, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)

            context = etree.iterparse(f, events=('end',), tag='{*}row')
            for _, row in context:
                row_data = []
                for cell in row.findall('.//{*}c'):
                    value_elem = cell.find('{*}v')
                    cell_type = cell.get('t')

                    if value_elem is not None:
                        value = value_elem.text
                        if cell_type == 's' and value.isdigit():
                            # 文字列データの処理（Shared Strings）
                            value = shared_strings[int(value)]
                        else:
                            # 数字（シリアル値）の場合
                            try:
                                # シリアル値を日付に変換
                                excel_date = datetime(1899, 12, 30) + timedelta(days=int(value))

                                # 「作業日」列（1列目）を仮定して変換
                                if len(row_data) == 0:  # 1列目（作業日）の処理
                                    value = excel_date.strftime('%Y-%m-%d')  # YYYY-MM-DD形式に変換
                            except ValueError:
                                pass  # 日付として変換できない場合はそのまま
                    else:
                        value = ''

                    row_data.append(value)

                writer.writerow(row_data)
                row.clear()

    print(f"データを {output_csv} に保存しました。")

# 実行
excel_file = "input_file_large.xlsx"
csv_output_file = "output.csv"

time_start = time.time()
print("<処理開始>")

read_excel_large(excel_file, csv_output_file)

time_end = time.time() - time_start
print("<処理時間> ", time_end)
