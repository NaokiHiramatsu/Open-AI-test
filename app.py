from flask import Flask, request, jsonify, render_template, session
import openai
import os
import pandas as pd

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # セッション用のシークレットキー

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

@app.route('/')
def index():
    # セッションからチャット履歴を取得
    chat_history = session.get('chat_history', [])
    return render_template('index.html', chat_history=chat_history)

@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')

    # チャット履歴をセッションから取得
    chat_history = session.get('chat_history', [])

    try:
        # ファイルがアップロードされた場合、ファイルの処理を実施
        if files:
            for file in files:
                df = pd.read_excel(file, engine='openpyxl')
                # ここでデータフレームの内容を文字列化して表示に使う（省略）

        # OpenAIの呼び出し
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=[
                {"role": "system", "content": "あなたは有能な生成AIです。"},
                {"role": "user", "content": prompt}
            ],
            max_tokens=2000
        )

        assistant_reply = response['choices'][0]['message']['content']

        # チャット履歴を更新
        chat_history.append({
            'user': prompt,
            'assistant': assistant_reply
        })

        # セッションにチャット履歴を保存
        session['chat_history'] = chat_history

        return render_template('index.html', chat_history=chat_history)

    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
