import pandas as pd

def convert_serial_date(csv_file):
    """
    CSVファイルの日付（シリアル値）をyyyy/mm/dd形式に変換する。

    Args:
        csv_file (str): CSVファイルのパス。

    Returns:
        pandas.DataFrame: 変換後のデータフレーム。
    """

    df = pd.read_csv(csv_file)

    # 日付列の名前を特定する（必要に応じて変更）
    date_column = '作業日'

    # シリアル値をdatetime型に変換
    df[date_column] = pd.to_datetime(df[date_column], unit='D', origin='1899-12-30')

    return df

# CSVファイルのパス
csv_file = 'output.csv'

# 日付変換
df_converted = convert_serial_date(csv_file)

# 変換後のデータフレームを表示
# print(df_converted)

# 変換後のデータフレームを新しいCSVファイルに保存（必要であれば）
df_converted.to_csv('converted_file.csv', index=False)
