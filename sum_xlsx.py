import pandas as pd
import time

# Excelファイルの読み込み（Sheet1から）
start_time_r = time.time()
df = pd.read_excel('/Users/mukaikazuhiro/Documents/kuronyan-sleep/local_file/input_file.xlsx', sheet_name='Sheet1')
end_time_r = time.time()
print(f"ファイルの読み取り時間: {end_time_r - start_time_r:.6f} 秒")
print(df.head(2))

# ユーザー入力
datefmt = input("dateフォーマットを入力してください（例：%m-%d-%y、%Y/%m/%dなど：")
subject = input("出力したい科目を入力してください（例：国語、数学、英語など）: ")
year = input("出力したい年度を入力してください（例：2024）: ")
affiliation = input("出力したい所属を入力してください（例：D、Aなど）: ")

# csvのdateフォーマットを指定
df['年月日'] = pd.to_datetime(df['年月日'], format=datefmt, errors='coerce')
print(df.head(2))

# 年度の入力チェック
if not year.isdigit():
    print("無効な年度が入力されました。数値を入力してください。")
    exit()
year = int(year)

# 分析開始時間を記録
start_time_a = time.time()

# 科目、所属のデータを高速フィルタリング
df_filtered = df.query("所属 == @affiliation and 科目 == @subject").copy()
if df_filtered.empty:
    print(f"指定した科目「{subject}」、所属「{affiliation}」のデータは存在しません。")
    exit()

# 年度と月を計算し、fiscal_month と fiscal_year を作成
df_filtered = df_filtered.assign(
    年度=df_filtered['年月日'].dt.year,
    月=df_filtered['年月日'].dt.month
)
df_filtered['fiscal_month'] = df_filtered['月'] + df_filtered['月'].lt(4) * 12
df_filtered['fiscal_year'] = df_filtered['年度'] - df_filtered['月'].lt(4)

# 指定した年度のデータを抽出
df_filtered = df_filtered[df_filtered['fiscal_year'].eq(year) & df_filtered['fiscal_month'].between(4, 15)]
if df_filtered.empty:
    print(f"指定した科目「{subject}」および年度「{year}」に対応するデータは存在しません。")
    exit()

# 氏名ごとの月別得点を集計
grouped = df_filtered.groupby(['氏名', 'fiscal_month'])['得点'].sum().unstack(fill_value=0)
# 4月～15月のカラムを確保
grouped = grouped.reindex(columns=range(4, 16), fill_value=0)
grouped['年間合計'] = grouped.sum(axis=1)

# 合計行の追加
total_row = grouped.sum(axis=0)
total_row.name = '合計'
grouped = pd.concat([grouped, total_row.to_frame().T])

# 科目、年度、所属の情報を追加
grouped['科目'] = subject
grouped['年度'] = year
grouped['所属'] = affiliation

# 結果の出力
df_result = grouped.reset_index()
columns = ['氏名'] + [f'{i}月' for i in range(4, 13)] + [f'{i}月' for i in range(1, 4)] + ['年間合計', '科目', '年度', '所属']
df_result.columns = columns

# 分析終了時間を記録
end_time_a = time.time()
print(f"データの分析時間: {end_time_a - start_time_a:.6f} 秒")

# Excelに出力（xlsxwriterを使用して高速化）
output_file = f'output_{subject}_{year}_{affiliation}.xlsx'
df_result.to_excel(output_file, index=False, engine='xlsxwriter')
print(f"{subject}のデータ（年度：{year}、所属：{affiliation}）が {output_file} に書き出されました。")
