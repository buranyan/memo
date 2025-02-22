import zipfile
import csv
from lxml import etree
import time
from datetime import datetime, timedelta

def read_excel_large(file_path, sheet_name, output_csv):
    with zipfile.ZipFile(file_path, 'r') as zip_data:
        # ZIP内のファイルを確認
        file_list = zip_data.namelist()

        # 指定されたシート名に対応するファイルを探す
        sheet_file = None
        for file in file_list:
            if file.startswith('xl/worksheets/') and file.endswith(sheet_name + '.xml'):
                sheet_file = file
                break

        if sheet_file is None:
            raise FileNotFoundError(f"{sheet_name}.xml が見つかりません")

        # `sharedStrings.xml` を読み込む（存在する場合のみ）
        shared_strings = []
        if 'xl/sharedStrings.xml' in file_list:
            shared_data = zip_data.read('xl/sharedStrings.xml')
            shared_root = etree.XML(shared_data)
            shared_strings = [s.text if s is not None else '' for s in shared_root.findall('.//{*}si/{*}t')]

        # 指定されたシートのXMLファイルをストリーム処理で読み込む
        with zip_data.open(sheet_file) as f, open(output_csv, 'w', newline='', encoding='utf-8-sig') as csv_file:
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
                            index = int(value)
                            if 0 <= index < len(shared_strings):
                                value = shared_strings[index]
                            else:
                                value = f"[ERR: {index}]"  # エラー値の可視化
                        # 数字（シリアル値）の場合はそのまま
                    else:
                        value = ''

                    row_data.append(value)
                
                writer.writerow(row_data)
                del row

    print(f"データを {output_csv} に保存しました。")

# 実行
excel_file = "input_file_large.xlsx"
sheet_name = "sheet1"  # シート名を指定
csv_output_file = "out_list.csv"

time_start = time.time()
print("<処理開始>")

read_excel_large(excel_file, sheet_name, csv_output_file)

time_end = time.time() - time_start
print("<処理時間>", time_end)
