import pandas as pd
import time

def load_data(file_path, sheet_name):
    """Excelファイルを読み込み、DataFrameを返す"""
    try:
        start_time = time.time()
        print("<読み取り開始>")
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        print(f"<読み取り時間>: {time.time() - start_time:.6f} 秒")
        return df
    except FileNotFoundError:
        print(f"エラー: 指定されたExcelファイル '{file_path}' が見つかりません。")
        return None
    except ValueError as e:
        print(f"エラー: Excelの読み込み中に問題が発生しました - {e}")
        return None

def validate_inputs():
    """ユーザー入力のバリデーション"""
    datefmt = "%Y-%m-%d"  # ユーザー入力を省略し、固定フォーマットを使用
    subject = input("出力したい工事番号を入力してください（例：国語、数学、英語など）: ").strip()
    affiliation = input("出力したい所属を入力してください（例：D、A など）: ").strip()

    while True:
        year_input = input("出力したい年度を入力してください（例：2024）: ").strip()
        if year_input.isdigit():
            year = int(year_input)
            if 1900 <= year <= 2100:
                break
        print("無効な年度が入力されました。1900～2100の範囲で数値を入力してください。")
    
    return datefmt, subject, year, affiliation

def preprocess_dates(df, datefmt):
    """日付フォーマットを変換し、NaT の割合をチェック"""
    df['作業日'] = pd.to_datetime(df['作業日'], format=datefmt, errors='coerce')
    nat_ratio = df['作業日'].isna().mean()
    
    if nat_ratio > 0.5:
        print(f"エラー: 指定したフォーマット '{datefmt}' で変換できないデータが50%以上あります。")
        return None
    return df

def filter_data(df, subject, affiliation):
    """工事番号と所属でデータをフィルタリング"""
    df_filtered = df.loc[(df['所属'] == affiliation) & (df['工事番号'] == subject)].copy()
    if df_filtered.empty:
        print(f"指定した工事番号「{subject}」、所属「{affiliation}」のデータは存在しません。")
        return None
    return df_filtered

def analyze_data(df_filtered, year, subject, affiliation):
    """データの年度と月を計算し、集計"""
    df_filtered['年度'] = df_filtered['作業日'].dt.year
    df_filtered['月'] = df_filtered['作業日'].dt.month
    df_filtered['fiscal_month'] = df_filtered['月'] + (df_filtered['月'] < 4) * 12
    df_filtered['fiscal_year'] = df_filtered['年度'] - (df_filtered['月'] < 4)

    df_filtered = df_filtered[(df_filtered['fiscal_year'] == year) & df_filtered['fiscal_month'].between(4, 15)]
    if df_filtered.empty:
        print(f"指定した工事番号「{subject}」および年度「{year}」に対応するデータは存在しません。")
        return None

    grouped = df_filtered.groupby(['氏名', 'fiscal_month'])['実績工数'].sum().unstack(fill_value=0)
    grouped = grouped.reindex(columns=range(4, 16), fill_value=0)
    grouped['年間合計'] = grouped.sum(axis=1)

    total_row = grouped.sum(axis=0)
    total_row.name = '合計'
    grouped = pd.concat([grouped, total_row.to_frame().T])

    grouped['工事番号'] = subject
    grouped['年度'] = year
    grouped['所属'] = affiliation

    return grouped.reset_index()

def save_results(df_result, subject, year, affiliation):
    """結果をExcelに保存"""
    output_file = f'output_{subject}_{year}_{affiliation}.xlsx'
    try:
        df_result.to_excel(output_file, index=False, engine='xlsxwriter')
        print(f"{subject}のデータ（年度：{year}、所属：{affiliation}）が {output_file} に書き出されました。")
    except PermissionError:
        print(f"エラー: ファイル '{output_file}' が開かれています。閉じてから再実行してください。")
    except Exception as e:
        print(f"エラー: ファイルの保存中に問題が発生しました - {e}")

def main():
    file_path = 'input_file_large.xlsx'
    sheet_name = 'Sheet1'

    df = load_data(file_path, sheet_name)
    if df is None:
        return

    datefmt, subject, year, affiliation = validate_inputs()

    df = preprocess_dates(df, datefmt)
    if df is None:
        return

    df_filtered = filter_data(df, subject, affiliation)
    if df_filtered is None:
        return

    start_time = time.time()
    print("<分析開始>")
    df_result = analyze_data(df_filtered, year, subject, affiliation)
    if df_result is None:
        return
    print(f"<分析時間>: {time.time() - start_time:.6f} 秒")

    save_results(df_result, subject, year, affiliation)

if __name__ == "__main__":
    main()
