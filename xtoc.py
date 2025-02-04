import pandas as pd
import os

def excel_to_csv_fast(excel_path, sheet_name):
    try:
        # 出力するCSVファイル名をExcelのファイル名から自動生成
        csv_path = os.path.splitext(excel_path)[0] + ".csv"

        # Excelファイルの指定シートを読み込む（高速化のため openpyxl を使用）
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")

        # CSVファイルとして保存
        df.to_csv(csv_path, index=False, encoding="utf-8")

        print(f"{csv_path} に変換が完了しました。")
    except FileNotFoundError:
        print(f"エラー: ファイル '{excel_path}' が見つかりません。")
    except ValueError:
        print(f"エラー: シート '{sheet_name}' が存在しません。")
    except Exception as e:
        print(f"予期しないエラーが発生しました: {e}")

# ユーザーから入力を受け付ける
excel_path = input("変換するExcelファイルのパスを入力してください: ").strip()
sheet_name = input("読み込むシート名を入力してください: ").strip()

excel_to_csv_fast(excel_path, sheet_name)
