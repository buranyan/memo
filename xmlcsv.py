import zipfile
import csv
from lxml import etree

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
            
            # XML をパースしながら 1行ずつ処理
            context = etree.iterparse(f, events=('end',), tag='{*}row')
            for _, row in context:
                row_data = []
                for cell in row.findall('.//{*}c'):
                    value_elem = cell.find('{*}v')
                    cell_type = cell.get('t')  # 文字列の場合 's' になる
                    
                    if value_elem is not None:
                        value = value_elem.text
                        if cell_type == 's' and value.isdigit():  # 文字列セルの場合
                            value = shared_strings[int(value)]
                    else:
                        value = ''
                    
                    row_data.append(value)
                
                writer.writerow(row_data)
                row.clear()  # メモリを開放

    print(f"データを {output_csv} に保存しました。")

# 実行
excel_file = "input_file.xlsx"
csv_output_file = "output.csv"

read_excel_large(excel_file, csv_output_file)
