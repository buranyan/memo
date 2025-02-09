import zipfile
import csv
from lxml import etree
import time
from datetime import datetime, timedelta

def read_excel_large(file_path, sheet_name, output_csv, selected_columns):
    with zipfile.ZipFile(file_path, 'r') as zip_data:
        file_list = zip_data.namelist()
        sheet_file = None
        for file in file_list:
            if file.startswith('xl/worksheets/') and file.endswith(sheet_name + '.xml'):
                sheet_file = file
                break
        
        if sheet_file is None:
            raise FileNotFoundError(f"{sheet_name}.xml が見つかりません")

        shared_strings = []
        if 'xl/sharedStrings.xml' in file_list:
            shared_data = zip_data.read('xl/sharedStrings.xml')
            shared_root = etree.XML(shared_data)
            shared_strings = [s.text for s in shared_root.findall('.//{*}si/{*}t')]

        with zip_data.open(sheet_file) as f, open(output_csv, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            
            context = etree.iterparse(f, events=('end',), tag='{*}row')
            for _, row in context:
                row_data = {}
                
                for cell in row.findall('.//{*}c'):
                    cell_ref = cell.get('r')
                    col_letter = ''.join(filter(str.isalpha, cell_ref))
                    col_index = ord(col_letter) - ord('A')  # Excelの列を0始まりのインデックスに変換
                    
                    if col_index == 0 or col_index in selected_columns:  # 1列目＋指定列のみ処理
                        value_elem = cell.find('{*}v')
                        cell_type = cell.get('t')

                        if value_elem is not None:
                            value = value_elem.text
                            if cell_type == 's' and value.isdigit():
                                value = shared_strings[int(value)]
                            elif col_index == 0:  # 1列目（作業日）は日付に変換
                                try:
                                    excel_date = datetime(1899, 12, 30) + timedelta(days=int(value))
                                    value = excel_date.strftime('%Y-%m-%d')
                                except ValueError:
                                    pass
                        else:
                            value = ''
                        
                        row_data[col_index] = value
                
                # 1列目（0番目）と指定された列を取得
                output_row = [row_data.get(0, '')] + [row_data.get(col, '') for col in selected_columns]
                writer.writerow(output_row)
                row.clear()
    
    print(f"データを {output_csv} に保存しました。")

# 実行
excel_file = "input_file_large.xlsx"
sheet_name = "sheet1"
csv_output_file = "output.csv"
selected_columns = [1]  # 2, 3, 5列目を選択（0始まりで考える）

time_start = time.time()
print("<処理開始>")

read_excel_large(excel_file, sheet_name, csv_output_file, selected_columns)

time_end = time.time() - time_start
print("<処理時間>", time_end)
