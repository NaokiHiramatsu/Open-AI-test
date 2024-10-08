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
index_name = "vector-1727336068502"  # 使用するインデックス名を指定

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
    search_results = search_client.search(
        search_text=prompt,
        top=5,  # 検索結果を増やす
        query_type="semantic",  # セマンティック検索
        search_fields="content"  # 検索対象フィールドを限定
    )

    # 取得したドキュメントの最初の一部を要約
    relevant_docs = []
    for doc in search_results:
        doc_content = doc.get('content', '')
        relevant_docs.append(doc_content[:500])  # 各ドキュメントの最初の500文字を取得

    # Azure OpenAI にプロンプトと関連ドキュメントを送信
    response = openai.ChatCompletion.create(
        engine=deployment_name,
        messages=[
            {"role": "system", "content": "You are an assistant."},
            {"role": "user", "content": f"Based on the following documents:\n{'\n'.join(relevant_docs)}\nAnswer the question: {prompt}"}
        ],
        max_tokens=100
    )

    # 応答を JSON で返す
    return jsonify(response['choices'][0]['message']['content'])

if __name__ == '__main__':
    app.run(debug=True)
