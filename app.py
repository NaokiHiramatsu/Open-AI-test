import openai
import os
from flask import Flask, request, render_template, session
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import pandas as pd
from docx import Document
from pptx import Presentation
import pdfplumber  # PDF用ライブラリ

app = Flask(__name__)
app.secret_key = os.urandom(24)

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

# Azure Cognitive Search の設定
search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT")
search_service_key = os.getenv("AZURE_SEARCH_KEY")
index_name = "vector-1730110777868"

search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

# ホームページを表示するルート
@app.route('/')
def index():
    session.clear()
    return render_template('index.html', chat_history=[])

# ファイルとプロンプトを処理するルート
@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')

    if 'chat_history' not in session:
        session['chat_history'] = []

    # チャット履歴を基にAIへ送信するメッセージを構成
    messages = [{"role": "system", "content": "あなたは有能なアシスタントです。"}]
    for entry in session['chat_history']:
        messages.append({"role": "user", "content": entry['user']})
        messages.append({"role": "assistant", "content": entry['assistant']})

    # ファイルがアップロードされている場合の処理
    file_data_text = []
    try:
        for file in files:
            if file and file.filename.endswith('.xlsx'):
                # Excelファイルの処理
                df = pd.read_excel(file, engine='openpyxl')
                columns_text = " | ".join(df.columns.tolist())
                rows_text = [" | ".join(map(str, row.tolist())) for index, row in df.iterrows()]
                df_text = f"ファイル名: {file.filename}\n{columns_text}\n" + "\n".join(rows_text)
                file_data_text.append(df_text)

            elif file and file.filename.endswith('.pdf'):
                # PDFファイルの処理
                pdf_text = ""
                with pdfplumber.open(file) as pdf:
                    for page in pdf.pages:
                        pdf_text += page.extract_text() + "\n"
                file_data_text.append(f"ファイル名: {file.filename}\n{pdf_text}")

            elif file and file.filename.endswith('.docx'):
                # Wordファイルの処理
                doc = Document(file)
                doc_text = "\n".join([para.text for para in doc.paragraphs])
                file_data_text.append(f"ファイル名: {file.filename}\n{doc_text}")

            elif file and file.filename.endswith('.pptx'):
                # PowerPointファイルの処理
                prs = Presentation(file)
                ppt_text = ""
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            ppt_text += shape.text + "\n"
                file_data_text.append(f"ファイル名: {file.filename}\n{ppt_text}")

        file_contents = "\n\n".join(file_data_text)
        input_data = f"アップロードされたファイルの内容は次の通りです:\n{file_contents}\nプロンプト: {prompt}"
        
        search_results = search_client.search(search_text=prompt, top=3)
        relevant_docs = "\n".join([doc['chunk'] for doc in search_results])

        input_data_with_search = f"以下のドキュメントに基づいて質問に答えてください：\n{relevant_docs}\n質問: {input_data}"
        messages.append({"role": "user", "content": input_data_with_search})

        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=messages,
            max_tokens=2000
        )

        response_content = response['choices'][0]['message']['content']
        session['chat_history'].append({
            'user': input_data,
            'assistant': response_content
        })

        return render_template('index.html', chat_history=session['chat_history'])

    except Exception as e:
        return f"エラーが発生しました: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
