import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from scipy.stats import ttest_ind, ttest_rel
import matplotlib

# 日本語フォントを設定（明朝体）
matplotlib.rcParams['font.family'] = 'Hiragino Mincho ProN'

# 出力フォルダを作成
output_dir = "ttestResult"
os.makedirs(output_dir, exist_ok=True)

# Excelファイルの読み込み
file_path = "data/マインドセット.xlsx"  # 適宜変更
input_filename = os.path.basename(file_path).replace(".xlsx", "")

xls = pd.ExcelFile(file_path)

# 各シートのデータを取得
df_esports = pd.read_excel(xls, sheet_name="eスポ通年")
df_liberal = pd.read_excel(xls, sheet_name="リベ通年")
df_all_july = pd.read_excel(xls, sheet_name="全体7月")
df_all_dec = pd.read_excel(xls, sheet_name="全体12月")

# 対象のカラムリスト（必要な尺度を手動で追加）
scale_types = ["GRIT合計スコア", "根気尺度", "一貫性尺度", "外向性", "協調性", "勤勉性", "神経症傾向", "開放性", "マインドセット合計スコア"]

# 有意差水準の定義
def significance_label(p_value):
    p_value = round(p_value, 3)
    if p_value < 0.001:
        return "***"
    elif p_value < 0.01:
        return "**"
    elif p_value < 0.05:
        return "*"
    else:
        return "n.s."

# 結果保存用配列
stats_list = []
ttest_list = []

# 各因子ごとに処理を実行
for scale in scale_types:
    scale_july_esports = f"{scale}7月eスポ"
    scale_dec_esports = f"{scale}12月eスポ"
    scale_july_liberal = f"{scale}7月リベ"
    scale_dec_liberal = f"{scale}12月リベ"
    fig, axes = plt.subplots(2, 2, figsize=(12, 10))
    axes = axes.flatten()

    # 指定したカラムがデータフレームに存在するかチェック
    if all(col in df_all_july.columns for col in [scale_july_esports, scale_july_liberal]) and \
       all(col in df_all_dec.columns for col in [scale_dec_esports, scale_dec_liberal]):

        # データ整形
        data_july_esports = df_all_july[scale_july_esports].dropna().astype(float)
        data_july_liberal = df_all_july[scale_july_liberal].dropna().astype(float)
        data_dec_esports = df_all_dec[scale_dec_esports].dropna().astype(float)
        data_dec_liberal = df_all_dec[scale_dec_liberal].dropna().astype(float)

        # 自由度の計算
        df_july = len(data_july_esports) + len(data_july_liberal) - 2
        df_dec = len(data_dec_esports) + len(data_dec_liberal) - 2
        df_esports = len(data_july_esports) - 1
        df_liberal = len(data_july_liberal) - 1

        # t検定
        t_july, p_july = ttest_ind(data_july_esports, data_july_liberal, equal_var=False)
        t_dec, p_dec = ttest_ind(data_dec_esports, data_dec_liberal, equal_var=False)
        t_esports, p_esports = ttest_rel(data_july_esports, data_dec_esports) if len(data_july_esports) == len(data_dec_esports) else (None, None)
        t_liberal, p_liberal = ttest_rel(data_july_liberal, data_dec_liberal) if len(data_july_liberal) == len(data_dec_liberal) else (None, None)

        stats_list.extend([
            {"尺度": scale, "Group": "eスポーツ(7月)", "平均": round(data_july_esports.mean(), 2), "標準偏差": round(data_july_esports.std(), 2), "Sample Size": len(data_july_esports)},
            {"尺度": scale, "Group": "リベラル(7月)", "平均": round(data_july_liberal.mean(), 2), "標準偏差": round(data_july_liberal.std(), 2), "Sample Size": len(data_july_liberal)},
            {"尺度": scale, "Group": "eスポーツ(12月)", "平均": round(data_dec_esports.mean(), 2), "標準偏差": round(data_dec_esports.std(), 2), "Sample Size": len(data_dec_esports)},
            {"尺度": scale, "Group": "リベラル(12月)", "平均": round(data_dec_liberal.mean(), 2), "標準偏差": round(data_dec_liberal.std(), 2), "Sample Size": len(data_dec_liberal)},
        ])

        # t検定結果
        ttest_list.extend([
            {"尺度": scale, "Test": "対応なしt検定（7月）", "t検定結果": f"t({df_july}) = {t_july:.3f}, p = {p_july:.3f}"},
            {"尺度": scale, "Test": "対応なしt検定（12月）", "t検定結果": f"t({df_dec}) = {t_dec:.3f}, p = {p_dec:.3f}"},
            {"尺度": scale, "Test": "対応ありt検定（eスポーツ）", "t検定結果": f"t({df_esports}) = {t_esports:.3f}, p = {p_esports:.3f}"},
            {"尺度": scale, "Test": "対応ありt検定（リベラル）", "t検定結果": f"t({df_liberal}) = {t_liberal:.3f}, p = {p_liberal:.3f}"},
        ])

        # グラフ情報
        graph_info = [
                (f"{scale}（7月）", ["eスポーツ", "リベラル"], [data_july_esports.mean(), data_july_liberal.mean()],
                    [data_july_esports.std(), data_july_liberal.std()], p_july, axes[0]),

                (f"{scale}（12月）", ["eスポーツ", "リベラル"], [data_dec_esports.mean(), data_dec_liberal.mean()],
                    [data_dec_esports.std(), data_dec_liberal.std()], p_dec, axes[1]),

                (f"{scale}（eスポーツ）", ["7月", "12月"], [data_july_esports.mean(), data_dec_esports.mean()],
                    [data_july_esports.std(), data_dec_esports.std()], p_esports, axes[2]),

                (f"{scale}（リベラル）", ["7月", "12月"], [data_july_liberal.mean(), data_dec_liberal.mean()],
                    [data_july_liberal.std(), data_dec_liberal.std()], p_liberal, axes[3]),
            ]

        # グラフ作成ループ
        for title, x_labels, means, stds, p_value, ax in graph_info:
            bars = ax.bar(x_labels, means, yerr=stds, capsize=5, color=["#4c72b0", "#55a868"])
            ax.set_title(title)
            ax.set_ylabel(scale)
            ax.tick_params(axis="y", direction="in")

            # 有意差表示
            significance = significance_label(p_value)
            y_max = max([b.get_height() for b in bars])
            bar_height = max(stds) * 1.2

            x1, x2 = 0, 1
            y_pos = y_max + bar_height
            ax.plot([x1, x1, x2, x2], [y_pos-0.1, y_pos, y_pos, y_pos-0.1], 'k', lw=1.5)
            ax.text((x1 + x2) / 2, y_pos + 0.05, significance, ha='center', va='bottom', fontsize=12, color="black")

        plt.tight_layout()
        plt.savefig(f"{output_dir}/{scale}_comparison.png")
        plt.show()

# DataFrameを作成し、Excelに保存
df_stats = pd.DataFrame(stats_list)
df_ttest = pd.DataFrame(ttest_list)

output_stats_file = os.path.join(output_dir, f"{input_filename}_statistics.xlsx")

with pd.ExcelWriter(output_stats_file) as writer:
    df_stats.to_excel(writer, sheet_name="Basic Statistics", index=False)
    df_ttest.to_excel(writer, sheet_name="T-Test Results", index=False)

print(f"基礎統計量とt検定の結果を'{output_stats_file}'に保存しました！")

