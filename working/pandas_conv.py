import pandas as pd
import time

def read_excel_pandas(file_path, sheet_name, output_csv):
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    df.to_csv(output_csv, index=False, encoding="utf-8")
    print(f"データを {output_csv} に保存しました。")

if __name__ == "__main__":
    excel_file = "input_file_large.xlsx"
    sheet_name = "Sheet1"
    csv_output_file = "output_pandas.csv"

    time_start = time.time()
    print("<処理開始>")

    read_excel_pandas(excel_file, sheet_name, csv_output_file)

    time_end = time.time() - time_start
    print("<処理時間>", time_end)
