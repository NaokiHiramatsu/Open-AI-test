import openai
import os
from flask import Flask, request, jsonify, render_template

app = Flask(__name__)

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")  # 環境変数からエンドポイントを取得
openai.api_version = "2023-03-15-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")  # 環境変数からAPIキーを取得
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")  # デプロイ名

# GETリクエストでフォームを表示するルート
@app.route('/')
def index():
    return render_template('index.html')

# POSTリクエストでプロンプトを処理するルート
@app.route('/ask_openai', methods=['POST'])
def ask_openai():
    # フォームからプロンプトを取得
    prompt = request.form.get('prompt', '')

    # Azure OpenAIにリクエストを送信
    response = openai.ChatCompletion.create(
        engine=deployment_name,
        messages=[
            {"role": "system", "content": "You are an assistant."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=50
    )

    # OpenAIからの応答を返す
    return jsonify(response['choices'][0]['message']['content'])

if __name__ == '__main__':
    app.run(debug=True)
