import openai
import os
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
from flask import Flask, request, jsonify

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
index_name = "vector-1727336068502"

# SearchClient の設定
search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

@app.route('/ask_openai_rag', methods=['POST'])
def ask_openai_rag():
    # クライアントからの質問を取得
    data = request.json
    prompt = data.get('prompt', '')

    # Azure Searchで関連するドキュメントを検索
    search_results = search_client.search(search_text=prompt, top=3)
    relevant_docs = "\n".join([doc['content'] for doc in search_results])

    # Azure OpenAIにプロンプトと関連ドキュメントを送信
    response = openai.ChatCompletion.create(
        engine=deployment_name,
        messages=[
            {"role": "system", "content": "You are an assistant."},
            {"role": "user", "content": f"Based on the following documents:\n{relevant_docs}\nAnswer the question: {prompt}"}
        ],
        max_tokens=100
    )

    # OpenAIからの応答をJSONで返す
    return jsonify(response['choices'][0]['message']['content'])

if __name__ == '__main__':
    app.run(debug=True)
