import openai
import os
from flask import Flask, request, jsonify, render_template
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import pandas as pd

app = Flask(__name__)

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
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

# ホームページを表示するルート
@app.route('/')
def index():
    return render_template('index.html')

# ファイルとプロンプトを処理するルート
@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    application_list_file = request.files.get('application_list')
    preapproval_list_file = request.files.get('preapproval_list')
    prompt = request.form.get('prompt', '')

    # ファイルをDataFrameとして読み込む
    try:
        # 'openpyxl'エンジンを指定してExcelファイルを読み込む
        application_df = pd.read_excel(application_list_file, engine='openpyxl')
        preapproval_df = pd.read_excel(preapproval_list_file, engine='openpyxl')

        # ファイル内容を確認（行数を確認するなど）
        application_count = len(application_df)
        preapproval_count = len(preapproval_df)

        # OpenAIに送信するメッセージを構成
        input_data = f"申請リストには {application_count} 行、事前承認リストには {preapproval_count} 行があります。\nプロンプト: {prompt}"

        # Azure OpenAI にプロンプトと関連データを送信
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=[
                {"role": "system", "content": "あなたは有能なアシスタントです。"},
                {"role": "user", "content": input_data}
            ],
            max_tokens=500  # 応答のトークン数を増やす
        )

        # 応答を返す
        return jsonify(response['choices'][0]['message']['content'])

    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
