# MVVS 0.5sq×2芯 + UL1007 AWG20 Python解析

Excelの差動／3導体同相インピーダンス解析をPythonへ移植したものです。

## セットアップ

```bash
python -m pip install -r requirements.txt
```

## 実行

```bash
python mvvs_ul1007_mtl_analysis.py
```

結果は既定で `mvvs_ul1007_results` フォルダへ出力されます。

## 設定値を変更する

```bash
python mvvs_ul1007_mtl_analysis.py --write-default-config config.json
```

生成された `config.json` を編集後、次で実行します。

```bash
python mvvs_ul1007_mtl_analysis.py --config config.json
```

## 実測値との比較

CSV列名:

```text
frequency_hz,magnitude_ohm,phase_deg
```

実行例:

```bash
python mvvs_ul1007_mtl_analysis.py   --measured-dm measured_dm.csv   --measured-cm measured_cm.csv
```

## 主な出力

- `derived_parameters.csv`
- `differential_results.csv`
- `common_mode_results.csv`
- `analysis_summary.txt`
- 差動／同相の振幅・位相PNG
- 実測比較CSV・PNG（実測CSV指定時）
