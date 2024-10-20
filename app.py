import openai
import os
from flask import Flask, request, render_template
import pandas as pd

app = Flask(__name__)

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

# チャット履歴を保持するリスト
chat_history = []

# ホームページを表示するルート
@app.route('/')
def index():
    return render_template('index.html', chat_history=chat_history)

# ファイルとプロンプトを処理するルート
@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    application_list_file = request.files.get('application_list')
    preapproval_list_file = request.files.get('preapproval_list')
    prompt = request.form.get('prompt', '')

    # ファイルの内容を読み込む処理（任意）
    application_data = ""
    preapproval_data = ""

    try:
        if application_list_file:
            application_df = pd.read_excel(application_list_file, engine='openpyxl')
            application_data = application_df.to_string(index=False)
        
        if preapproval_list_file:
            preapproval_df = pd.read_excel(preapproval_list_file, engine='openpyxl')
            preapproval_data = preapproval_df.to_string(index=False)
        
        input_data = f"申請リスト:\n{application_data}\n\n" \
                     f"事前承認リスト:\n{preapproval_data}\n\n" \
                     f"質問:\n{prompt}"

    except Exception as e:
        input_data = f"プロンプトのみ: {prompt}"

    # チャット履歴に今回の入力を追加
    chat_history.append({"role": "user", "content": prompt})

    # Azure OpenAI にプロンプトと関連データを送信
    try:
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=chat_history + [{"role": "user", "content": input_data}],
            max_tokens=2000
        )
        
        # 応答をチャット履歴に追加
        response_content = response['choices'][0]['message']['content']
        chat_history.append({"role": "assistant", "content": response_content})
        
        # 更新されたチャット履歴を表示
        return render_template('index.html', chat_history=chat_history)

    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
