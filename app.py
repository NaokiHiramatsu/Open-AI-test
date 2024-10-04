import openai
import os
from flask import Flask, request, jsonify

app = Flask(__name__)

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")  # 環境変数からエンドポイントを取得
openai.api_version = "2023-03-15-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")  # 環境変数からAPIキーを取得
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")  # デプロイ名

@app.route('/ask_openai', methods=['POST'])
def ask_openai():
    # クライアントからのプロンプトを取得
    data = request.json
    prompt = data.get('prompt', '')

    # Azure OpenAIにリクエストを送信
    response = openai.ChatCompletion.create(
        engine=deployment_name,  # 環境変数から取得したデプロイ名を使用
        messages=[
            {"role": "system", "content": "You are an assistant."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=50
    )

    # OpenAIからの応答をJSONで返す
    return jsonify(response['choices'][0]['message']['content'])

if __name__ == '__main__':
    app.run(debug=True)
