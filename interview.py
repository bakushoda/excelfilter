import whisper
import json
import os

# Whisper モデルの読み込み
model = whisper.load_model("small")

# 出力フォルダの作成 (存在しない場合)
output_dir = "interviewResult"
os.makedirs(output_dir, exist_ok=True)

# 音声ファイルの変換
result = model.transcribe("interviewData/じぶん.m4a", verbose=True, fp16=False, language="ja")

# 結果をテキストファイルに保存
output_file = os.path.join(output_dir, 'transcription.txt')
with open(output_file, 'w', encoding='UTF-8') as f:
    json.dump(result['text'], f, sort_keys=True, indent=4, ensure_ascii=False)

print(f"Transcription saved to {output_file}")
