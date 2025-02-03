import pandas as pd

# サンプルデータの作成
data = {
    '商品': ['A', 'B', 'A', 'B', 'C', 'A', 'C'],
    '売上': [100, 200, 150, 250, 300, 200, 350]
}
df = pd.DataFrame(data)
print("元のデータフレーム:")
import pandas as pd

# サンプルデータの作成
data = {
    '商品': ['A', 'B', 'A', 'B', 'C', 'A', 'C'],
    '売上': [100, 200, 150, 250, 300, 200, 350]
}
df = pd.DataFrame(data)
print("元のデータフレーム:")
print(df)

# '商品'ごとにデータをグループ化する
grouped = df.groupby('商品')
print("'商品'ごとにデータをグループ化する")
print(grouped)

# 各グループの売上合計を計算する
result_sum = grouped.sum()
print("\n商品ごとの売上合計:")
print(result_sum)

# 各グループの売上平均を計算する
result_mean = grouped.mean()
print("\n商品ごとの売上平均:")
print(result_mean)

# 複数の集計関数を同時に適用する例（合計と平均）
result_agg = grouped.agg({'売上': ['sum', 'mean']})
print("\n商品ごとの売上の合計と平均:")
print(result_agg)