import openai
import os
from flask import Flask, request, jsonify, render_template
import pandas as pd

app = Flask(__name__)

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

# チャット履歴を格納するリスト
chat_history = []

# ホームページを表示するルート
@app.route('/')
def index():
    return render_template('index.html', chat_history=chat_history)

# ファイルとプロンプトを処理するルート
@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')

    file_details = ""
    
    # ファイルがあれば処理
    if files:
        for file in files:
            try:
                # ExcelファイルをDataFrameとして読み込む
                df = pd.read_excel(file, engine='openpyxl')
                file_details += f"{file.filename}に{len(df)}行あります。\n"
            except Exception as e:
                return f"エラーが発生しました: {str(e)}"

    # OpenAIに送信するメッセージを構成
    input_data = f"{file_details}\nプロンプト: {prompt}"

    # Azure OpenAI にプロンプトと関連データを送信
    try:
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=[
                {"role": "system", "content": "あなたは有能なアシスタントです。"},
                {"role": "user", "content": input_data}
            ],
            max_tokens=2000
        )
        assistant_reply = response['choices'][0]['message']['content']

        # チャット履歴に追加
        chat_history.append({"user": prompt, "assistant": assistant_reply})

        return render_template('index.html', chat_history=chat_history)

    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
