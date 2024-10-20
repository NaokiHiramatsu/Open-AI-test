import openai
import os
from flask import Flask, request, render_template, session
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import pandas as pd

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # セッションの秘密鍵

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

# ホームページを表示するルート
@app.route('/')
def index():
    # セッションに保存されたチャット履歴を取得
    chat_history = session.get('chat_history', [])
    return render_template('index.html', chat_history=chat_history)

# ファイルとプロンプトを処理するルート
@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')  # 複数ファイルの取得
    prompt = request.form.get('prompt', '')

    # セッションからチャット履歴を取得
    chat_history = session.get('chat_history', [])

    # ファイルがアップロードされている場合の処理
    file_data_text = []
    try:
        for file in files:
            if file and file.filename.endswith('.xlsx'):  # ファイルが存在し、かつxlsx形式か確認
                # 'openpyxl'エンジンを使用してExcelファイルを読み込む
                df = pd.read_excel(file, engine='openpyxl')

                # 列名を取得して文字列化
                columns = df.columns.tolist()
                columns_text = " | ".join(columns)  # 列名を区切り文字で結合

                # 各行のデータを行ごとに文字列化
                rows_text = []
                for index, row in df.iterrows():
                    row_text = " | ".join([str(item) for item in row.tolist()])  # 各行のデータを区切り文字で結合
                    rows_text.append(row_text)

                # 列名と行データを連結
                df_text = f"ファイル名: {file.filename}\n{columns_text}\n" + "\n".join(rows_text)
                file_data_text.append(df_text)
            else:
                continue  # 不正なファイル形式の場合、処理をスキップ

        # ファイルがある場合、内容を結合して送信するデータに追加
        if file_data_text:
            file_contents = "\n\n".join(file_data_text)
            input_data = f"アップロードされたファイルの内容は次の通りです:\n{file_contents}\nプロンプト: {prompt}"
        else:
            # ファイルがない場合でも、プロンプトをそのまま使用してOpenAIにリクエストを送信
            input_data = f"プロンプトのみが入力されました:\nプロンプト: {prompt}"

        # Azure OpenAI にプロンプトとファイル内容を送信
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=[
                {"role": "system", "content": "あなたは有能なアシスタントです。"},
                {"role": "user", "content": input_data}
            ],
            max_tokens=2000  # 応答のトークン数を増やす
        )

        # 応答の内容を取得
        response_content = response['choices'][0]['message']['content']

        # ユーザーの入力とアシスタントの応答をチャット履歴に追加
        chat_history.append({'user': prompt, 'assistant': response_content})
        session['chat_history'] = chat_history  # チャット履歴をセッションに保存

        # チャット履歴をテンプレートに渡して表示
        return render_template('index.html', chat_history=chat_history)

    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
