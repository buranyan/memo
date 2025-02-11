import zipfile
import csv
import gc
import time
import numpy as np
from lxml import etree

def load_shared_strings(zip_data):
    """ sharedStrings.xml を NumPy 配列としてロード（超高速アクセス） """
    shared_strings = []
    if 'xl/sharedStrings.xml' in zip_data.namelist():
        shared_data = zip_data.read('xl/sharedStrings.xml')
        shared_root = etree.XML(shared_data, etree.XMLParser(recover=True))
        shared_strings = [s.text if s is not None else '' for s in shared_root.findall('.//{*}si/{*}t')]

    return np.array(shared_strings, dtype=object) if shared_strings else None  # NumPy に変換（超高速化）

def get_namespace(element):
    """XMLツリーの名前空間を取得"""
    match = etree.QName(element.tag)
    return { 'ns': match.namespace } if match.namespace else {}

def read_excel_large(file_path, sheet_name, output_csv):
    """ Excelファイルを解析してCSVに書き出す（最適化版） """
    with zipfile.ZipFile(file_path, 'r') as zip_data:
        file_list = zip_data.namelist()

        # 指定シートのXMLファイル検索
        sheet_file = next((file for file in file_list if file.startswith(f'xl/worksheets/{sheet_name}.xml')), None)
        if sheet_file is None:
            raise FileNotFoundError(f"{sheet_name}.xml が見つかりません")

        # sharedStrings.xml を NumPy 配列でロード
        shared_strings = load_shared_strings(zip_data)

        # シートのXMLを解析
        with zip_data.open(sheet_file, 'r') as f:
            context = etree.iterparse(f, events=('start', 'end'))
            _, root = next(context)  # 最初の要素（worksheet）を取得
            namespaces = get_namespace(root)  # 名前空間を取得

            # タグをデバッグ表示
            # print(f"XML 名前空間: {namespaces}")

            # XPath の事前コンパイル（名前空間対応）
            row_finder = etree.XPath('.//ns:row', namespaces=namespaces)
            cell_finder = etree.XPath('.//ns:c', namespaces=namespaces)
            value_finder = etree.XPath('./ns:v', namespaces=namespaces)

            buffer = []
            buffer_size = 50000  # バッファサイズ
            with open(output_csv, 'w', newline='', buffering=1024*1024, encoding='utf-8') as csv_file:
                writer = csv.writer(csv_file, quoting=csv.QUOTE_MINIMAL)

                for row_index, (event, row) in enumerate(context, start=1):
                    if event == 'end' and row.tag.endswith('row'):
                        # デバッグ: 行データを確認
                        # print(f"処理中: {row_index} 行目")

                        row_data = []
                        for cell in cell_finder(row):
                            value_elements = value_finder(cell)
                            cell_type = cell.get('t')

                            if value_elements:
                                value = value_elements[0].text
                                if cell_type == 's' and value.isdigit():
                                    index = int(value)
                                    value = shared_strings[index] if 0 <= index < len(shared_strings) else ''
                            else:
                                value = ''

                            row_data.append(value)

                        buffer.append(row_data)

                        # 一定行ごとにバッファ書き出し
                        if row_index % buffer_size == 0:
                            writer.writerows(buffer)
                            buffer.clear()
                            csv_file.flush()
                            print(f"{row_index} 行処理完了...")
                            gc.collect()  # メモリ解放

                        row.clear()  # lxml のメモリ管理
                        del row

                # 残りデータを書き出す
                if buffer:
                    writer.writerows(buffer)

    print(f"データを {output_csv} に保存しました。")

# 実行
if __name__ == "__main__":
    excel_file = "input_file_large.xlsx"
    sheet_name = "sheet1"  # シート名
    csv_output_file = "output.csv"

    time_start = time.time()
    print("<処理開始>")

    read_excel_large(excel_file, sheet_name, csv_output_file)

    time_end = time.time() - time_start
    print("<処理時間>", time_end)
