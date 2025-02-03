from xlsx2csv import Xlsx2csv
import os

def excel_to_csv_fast(excel_file, sheet_name):
    """xlsx2csvを使ってExcelの指定シートをCSVに変換"""
    # ファイルの存在確認
    if not os.path.isfile(excel_file):
        print(f"❌ エラー: ファイル '{excel_file}' が見つかりません。正しいパスを指定してください。")
        return

    # ファイル形式の確認
    if not excel_file.lower().endswith(".xlsx"):
        print("❌ エラー: xlsx2csv は .xls 形式をサポートしていません。Excel で .xlsx に変換してください。")
        return

    # 出力ファイル名を作成
    base_name, _ = os.path.splitext(excel_file)
    csv_file = f"{base_name}.csv"

    try:
        # 変換処理
        Xlsx2csv(excel_file, outputencoding="utf-8").convert(csv_file, sheetname=sheet_name)
        print(f"✅ {sheet_name} を {csv_file} に変換しました！")
    except Exception as e:
        print(f"❌ エラー: 変換に失敗しました - {e}")

# --- 使い方 ---
excel_file = input("Excelファイル名を入力してください（例: data.xlsx）: ").strip()
sheet_name = input("変換するシート名を入力してください（例: Sheet1）: ").strip()

excel_to_csv_fast(excel_file, sheet_name)
