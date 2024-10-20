import openai
import os
from flask import Flask, request, render_template, session
import pandas as pd

app = Flask(__name__)
app.secret_key = os.urandom(24)  # セッションを管理するための秘密鍵

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

# ホームページを表示するルート
@app.route('/')
def index():
    session.clear()  # ブラウザを閉じたらセッションをリセット
    return render_template('index.html', chat_history=[])

# ファイルとプロンプトを処理するルート
@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')  # 複数ファイルの取得
    prompt = request.form.get('prompt', '')

    if 'chat_history' not in session:
        session['chat_history'] = []  # チャット履歴を初期化

    # チャット履歴を基にAIへ送信するメッセージを構成
    messages = [{"role": "system", "content": "あなたは有能なアシスタントです。"}]
    
    # チャット履歴を追加
    for entry in session['chat_history']:
        messages.append({"role": "user", "content": entry['user']})
        messages.append({"role": "assistant", "content": entry['assistant']})

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

        # 最新のプロンプトを追加
        messages.append({"role": "user", "content": input_data})

        # Azure OpenAI にプロンプトとファイル内容を送信
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=messages,  # これまでの履歴も含めて送信
            max_tokens=2000  # 応答のトークン数を増やす
        )

        # 応答内容を取得し、履歴に追加
        response_content = response['choices'][0]['message']['content']
        session['chat_history'].append({
            'user': input_data,
            'assistant': response_content
        })

        # 応答をテンプレートに渡して表示
        return render_template('index.html', chat_history=session['chat_history'])

    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
