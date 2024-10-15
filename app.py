import openai
import os
from flask import Flask, request, jsonify, render_template
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential

app = Flask(__name__)

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2023-03-15-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

# Azure Cognitive Search の設定
search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT")
search_service_key = os.getenv("AZURE_SEARCH_KEY")
index_name = "hiramatsu2"  # 使用するインデックス名を指定

# SearchClient の設定
search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

# フォームを表示するルート
@app.route('/')
def index():
    return render_template('index.html')

# POSTリクエストでプロンプトを処理するルート
@app.route('/ask_openai', methods=['POST'])
def ask_openai():
    prompt = request.form.get('prompt', '')

    # Azure Search で関連するドキュメントを検索
    search_results = search_client.search(search_text=prompt, top=3)
    relevant_docs = "\n".join([doc['chunk'] for doc in search_results])  # 'content' を 'chunk' に変更

    # Azure OpenAI にプロンプトと関連ドキュメントを送信（日本語対応）
    response = openai.ChatCompletion.create(
        engine=deployment_name,
        messages=[
            {"role": "system", "content": "あなたは有能なアシスタントです。"},
            {"role": "user", "content": f"以下のドキュメントに基づいて質問に答えてください：\n{relevant_docs}\n質問: {prompt}"}
        ],
        max_tokens=500  # 応答のトークン数を増やす
    )

    # 応答を JSON で返す
    return jsonify(response['choices'][0]['message']['content'])

if __name__ == '__main__':
    app.run(debug=True)
