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

    try:
        # 'openpyxl'エンジンを指定してExcelファイルを読み込む
        application_df = pd.read_excel(application_list_file, engine='openpyxl')
        preapproval_df = pd.read_excel(preapproval_list_file, engine='openpyxl')

        # DataFrameの内容を文字列として取得
        application_data = application_df.to_string(index=False)
        preapproval_data = preapproval_df.to_string(index=False)

        # OpenAIに送信するメッセージを構成
        input_data = f"以下は申請リストの内容です:\n{application_data}\n\n" \
                     f"以下は事前承認リストの内容です:\n{preapproval_data}\n\n" \
                     f"これに基づいて、次の質問に回答してください:\n{prompt}"

        # Azure OpenAI にプロンプトと関連データを送信
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=[
                {"role": "system", "content": "あなたは有能なアシスタントです。"},
                {"role": "user", "content": input_data}
            ],
            max_tokens=2000  # 応答のトークン数を増やす
        )

        # 応答をテンプレートに渡して表示（改行を含む出力に変更）
        response_content = response['choices'][0]['message']['content'].replace("\n", "<br>")
        return render_template('index.html', response_content=response_content)

    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
