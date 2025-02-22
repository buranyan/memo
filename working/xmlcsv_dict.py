import zipfile
import csv
from lxml import etree
import time

def read_excel_large(file_path, sheet_name, output_csv):
    with zipfile.ZipFile(file_path, 'r') as zip_data:
        file_list = zip_data.namelist()

        # 指定されたシート名に対応するファイルを探す
        sheet_file = next((file for file in file_list if file.startswith(f'xl/worksheets/{sheet_name}.xml')), None)
        if sheet_file is None:
            raise FileNotFoundError(f"{sheet_name}.xml が見つかりません")

        # `sharedStrings.xml` を辞書で読み込む
        shared_strings = {}
        if 'xl/sharedStrings.xml' in file_list:
            shared_data = zip_data.read('xl/sharedStrings.xml')
            shared_root = etree.XML(shared_data, etree.XMLParser(recover=True))
            shared_strings = {i: s.text if s is not None else '' for i, s in enumerate(shared_root.findall('.//{*}si/{*}t'))}

        # シートのXMLファイルをストリーム処理で読み込む
        with zip_data.open(sheet_file, 'r') as f, open(output_csv, 'w', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.writer(csv_file, quoting=csv.QUOTE_MINIMAL)
            context = etree.iterparse(f, events=('end',), tag='{*}row')

            buffer = []
            buffer_size = 50000  # メモリ管理のためバッファサイズを調整
            for row_index, (_, row) in enumerate(context, start=1):
                row_data = []
                for cell in row.iterfind('{*}c'):
                    value_elem = cell.find('{*}v')
                    cell_type = cell.get('t')

                    if value_elem is not None:
                        value = value_elem.text
                        if cell_type == 's' and value.isdigit():
                            index = int(value)
                            value = shared_strings.get(index, f"[ERR: {index}]")  # エラー処理追加
                    else:
                        value = ''

                    row_data.append(value)

                buffer.append(row_data)

                # バッファサイズごとに書き出す
                if row_index % buffer_size == 0:
                    writer.writerows(buffer)
                    buffer.clear()  # メモリ解放
                    csv_file.flush()

                row.clear()  # メモリ解放

            # 残りのデータを書き出す
            if buffer:
                writer.writerows(buffer)

    print(f"データを {output_csv} に保存しました。")

# 実行
if __name__ == "__main__":
    excel_file = "input_file_large.xlsx"
    sheet_name = "sheet1"  # シート名を指定
    csv_output_file = "out_dict.csv"

    time_start = time.time()
    print("<処理開始>")

    read_excel_large(excel_file, sheet_name, csv_output_file)

    time_end = time.time() - time_start
    print("<処理時間>", time_end)
