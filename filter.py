import pandas as pd
import os
from course_list import course_esports, course_liberal, course_all

# ファイルパス
file_path = r""
# シート名
sheet_name = ""

# データの読み込み
df = pd.read_excel(file_path, sheet_name=sheet_name)

# subjectId 列に一致する名前だけを残す
filtered_df = df[df['subjectId'].isin(course_liberal)].copy()

# 並べ替えの順番を指定
filtered_df['subjectId'] = pd.Categorical(filtered_df['subjectId'], categories=course_liberal, ordered=True)

# 並べ替えたデータフレーム
sorted_df = filtered_df.sort_values('subjectId')

# 出力先ディレクトリを作成（存在しない場合）
output_dir = "filteredResult"
os.makedirs(output_dir, exist_ok=True)

# 元のファイル名から拡張子を取り除き、結果ファイル名を作成
original_file_name = os.path.basename(file_path).replace(".xlsx", "")
output_file_name = f"filtered_{sheet_name}_{original_file_name}.xlsx"
output_file = os.path.join(output_dir, output_file_name)

# エクセルファイルに保存
sorted_df.to_excel(output_file, index=False)

print(f"フィルタリングと並べ替えが完了しました。'{output_file}' に保存されました。")
